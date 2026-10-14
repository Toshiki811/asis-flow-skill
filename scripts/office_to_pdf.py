#!/usr/bin/env python3
"""office_to_pdf.py — PowerPoint・Excel・Word のファイルを、見た目どおりの PDF にする（render_pages.py が使う）。

  python3 scripts/office_to_pdf.py --probe            # どの方法が使えるかを調べる
  python3 scripts/office_to_pdf.py in.pptx out_dir    # 単体で試す

方法は2つ。auto は LibreOffice があればそれを使い、なければ Microsoft Office を使う。
  soffice  LibreOffice を画面なしで動かす。確認ダイアログが出ないので大量処理に向く。Excel は1シート＝1ページで出す
  office   macOS の Microsoft Office を AppleScript で動かす。見た目は最も原本に近い。Excel はシートごとに PDF にする。
           初回は macOS が「〜の制御を許可するか」を尋ねるので、利用者が許可する

マクロは実行しない。マクロ付きのファイル（xlsm・pptm・docm）は、マクロを抜いた写しを作ってから開く（defang）。
旧形式（xls・ppt・doc）のマクロ付きは抜けないので、Microsoft Office では開かない。
"""
import argparse
import os
import pathlib
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

MAC_APPS = {
    "powerpoint": pathlib.Path("/Applications/Microsoft PowerPoint.app"),
    "excel": pathlib.Path("/Applications/Microsoft Excel.app"),
    "word": pathlib.Path("/Applications/Microsoft Word.app"),
}
APP_NAME = {"powerpoint": "Microsoft PowerPoint", "excel": "Microsoft Excel", "word": "Microsoft Word"}
APP_OF_EXT = {".pptx": "powerpoint", ".ppt": "powerpoint", ".pps": "powerpoint", ".ppsx": "powerpoint",
              ".xlsx": "excel", ".xls": "excel", ".docx": "word", ".doc": "word"}
SOFFICE_PATHS = ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice",
                 "/usr/local/bin/soffice", "/opt/homebrew/bin/soffice"]

# マクロ付き・テンプレートの本体の種別 → 普通の文書の種別
MAIN_TYPES = {
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
}
PLAIN_EXT = {".xlsm": ".xlsx", ".xltm": ".xlsx", ".xltx": ".xlsx", ".pptm": ".pptx", ".ppsm": ".pptx",
             ".ppsx": ".pptx", ".potx": ".pptx", ".docm": ".docx"}
MACRO_PARTS = re.compile(r"(vbaProject(Signature\w*)?\.bin|vbaData\.xml)$")


class ConvertError(Exception):
    pass


def find_soffice():
    for p in SOFFICE_PATHS:
        if pathlib.Path(p).exists():
            return p
    return shutil.which("soffice") or shutil.which("libreoffice")


def osa(lines, timeout=900):
    args = ["osascript"]
    for line in lines:
        args += ["-e", line]
    try:
        return subprocess.run(args, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise ConvertError("AppleScript が時間内に終わらなかった")


def app_running(app):
    """アプリが起動しているか（この問い合わせではアプリは起動しない）"""
    try:
        r = osa([f'return (application "{APP_NAME[app]}" is running) as string'], timeout=20)
    except ConvertError:
        return False
    return r.returncode == 0 and r.stdout.strip() == "true"


def is_timeout(err):
    return any(k in str(err) for k in ("-1712", "タイムアウト", "時間内に終わらなかった", "timed out"))


STUCK = ("{app} が応えない。{app} の画面に確認のダイアログ（アクセスの許可・起動時の案内・回復したファイルなど）が出ていれば、"
         "閉じてから再実行する。人が画面を見ていられないときは LibreOffice を入れる（brew install --cask libreoffice）")


def office_ready(app):
    """Microsoft Office のアプリが AppleScript で使える状態か。使えないときは理由を返す。
    起動していないアプリには問い合わせない（文書を渡さずに起動すると、起動時のギャラリーが出て後の操作が止まる）。
    起動していなければ、あとで文書を渡して起動する。"""
    if platform.system() != "Darwin" or not MAC_APPS[app].exists():
        return False, f"{APP_NAME[app]} が入っていない"
    if not app_running(app):
        return True, "未起動（文書を渡して起動する）"
    try:
        r = osa(["with timeout of 20 seconds", f'tell application "{APP_NAME[app]}" to return version as string',
                 "end timeout"], timeout=40)
    except ConvertError:
        r = None
    if r is not None and r.returncode == 0:
        return True, r.stdout.strip()
    return False, STUCK.format(app=APP_NAME[app])


def pick_backend(app, backend="auto"):
    """使う方法を決める。戻り値は 'soffice' か 'office'。どちらも使えなければ ConvertError"""
    so = find_soffice()
    if backend in ("auto", "soffice") and so:
        return "soffice"
    if backend == "soffice":
        raise ConvertError("LibreOffice（soffice）が見つからない")
    ok, why = office_ready(app)
    if ok:
        return "office"
    raise ConvertError(why)


def stage_dir():
    """変換のための一時フォルダ（短いパス）を作って返す。使い終わったら呼び出し側が消す。
    Excel はパスが長い（250字を超えるあたり）と保存に失敗し、Office は場所によって「アクセスを許可」の確認を出して止まる。
    macOS では /private/tmp の下がどちらにも当たらなかったので、そこを既定にする（ASISFLOW_STAGE で変えられる）"""
    base = os.environ.get("ASISFLOW_STAGE")
    if not base:
        base = f"/private/tmp/asisflow-{os.getuid()}" if platform.system() == "Darwin" else os.path.join(tempfile.gettempdir(), "asisflow")
    base = pathlib.Path(base)
    base.mkdir(parents=True, exist_ok=True)
    try:
        os.chmod(base, 0o700)
    except OSError:
        pass
    return pathlib.Path(tempfile.mkdtemp(prefix="c", dir=str(base)))


def defang(src, dst_dir):
    """マクロを抜いた写し（必要なら拡張子を普通の文書に直す）を dst_dir に作って返す。マクロは決して実行させない。
    写しの名前はフォルダ名と同じ短い英数字にする（Office が開いている文書を名前で指すため、ほかの文書と重ならない名前にする）"""
    src = pathlib.Path(src)
    dst_dir = pathlib.Path(dst_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)
    ext = src.suffix.lower()
    out = dst_dir / f"{dst_dir.name}{PLAIN_EXT.get(ext, ext)}"
    if not zipfile.is_zipfile(src):
        shutil.copy(src, out)
        return out
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        drop = {n for n in names if MACRO_PARTS.search(n)}
        if not drop and ext not in PLAIN_EXT:
            shutil.copy(src, out)
            return out
        drop_base = {n.rsplit("/", 1)[-1] for n in drop}
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in drop:
                    continue
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    xml = data.decode("utf-8")
                    for a, b in MAIN_TYPES.items():
                        xml = xml.replace(a, b)
                    xml = re.sub(r'<Override[^>]*PartName="[^"]*(?:vbaProject[^"]*\.bin|vbaData\.xml)"[^>]*/>', "", xml)
                    data = xml.encode("utf-8")
                elif item.filename.endswith(".rels") and drop_base:
                    xml = data.decode("utf-8")
                    for b in drop_base:
                        xml = re.sub(r'<Relationship\b[^>]*Target="[^"]*' + re.escape(b) + r'"[^>]*/>', "", xml)
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
    return out


def legacy_has_macro(path):
    """旧形式（xls・ppt・doc）にマクロがあるか。分からなければ True（安全側）"""
    try:
        from oletools.olevba import VBA_Parser
        vp = VBA_Parser(str(path))
        has = bool(vp.detect_vba_macros())
        vp.close()
        return has
    except Exception:
        return True


# ---- LibreOffice ----

def soffice_pdf(src, out_dir, single_page_sheets=False):
    """LibreOffice で PDF にする。Excel は single_page_sheets で1シート＝1ページ（印刷範囲に左右されない）"""
    so = find_soffice()
    out_dir = pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    fmt = "pdf"
    if single_page_sheets:
        fmt = 'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}'
    profile = out_dir / "_lo_profile"                                # 起動中の LibreOffice とぶつからないよう、専用の設定で動かす
    r = subprocess.run([so, f"-env:UserInstallation=file://{profile}", "--headless", "--norestore",
                        "--convert-to", fmt, "--outdir", str(out_dir), str(src)],
                       capture_output=True, text=True, timeout=900)
    pdf = out_dir / (pathlib.Path(src).stem + ".pdf")
    shutil.rmtree(profile, ignore_errors=True)
    if not pdf.exists():
        raise ConvertError(f"LibreOffice での PDF 化に失敗: {(r.stderr or r.stdout).strip()[:300]}")
    return pdf


# ---- Microsoft Office（macOS・AppleScript）----

def powerpoint_pdf(src, pdf):
    """PowerPoint で開いて PDF に書き出し、開いた写しだけを閉じる（consulting-pptx-skill の export_deck.py と同じ手順）"""
    src, pdf = pathlib.Path(src), pathlib.Path(pdf)
    pdf.unlink(missing_ok=True)
    r = osa([
        "with timeout of 600 seconds",
        'tell application "Microsoft PowerPoint"',
        "set opened to true",
        "try",
        "with timeout of 90 seconds",
        f'open POSIX file "{src}"',
        "end timeout",
        "on error",
        "set opened to false",
        "end try",
        "delay 2",
        "set p to missing value",
        "repeat 60 times",
        f'if (name of active presentation) is "{src.name}" then',
        "set p to active presentation",
        "exit repeat",
        "end if",
        "delay 1",
        "end repeat",
        'if p is missing value then error "開いた写しが前面に無い"',
        f'save p in POSIX file "{pdf}" as save as PDF',
        "if opened then close p saving no",
        "return opened",
        "end tell",
        "end timeout",
    ])
    for _ in range(60):                                              # PowerPoint は PDF を遅れて書くことがある
        if pdf.exists():
            break
        time.sleep(1)
    if r.returncode or not pdf.exists():
        raise ConvertError(f"PowerPoint での PDF 書き出しに失敗: {r.stderr.strip()[:300]}")
    if r.stdout.strip() == "false":
        print(f"  PowerPoint で {src.name} が開いたまま残った（保存せずに閉じてよい）")
    return pdf


def word_pdf(src, pdf):
    """Word で開いて PDF に書き出し、開いた写しだけを閉じる。手順を1つずつ送り、開くのに 90 秒以上かかったら諦める
    （Word が確認のダイアログを出して止まっていると、いつまでも返ってこない）"""
    src, pdf = pathlib.Path(src), pathlib.Path(pdf)
    pdf.unlink(missing_ok=True)

    def word(lines, timeout):
        r = osa(["with timeout of %d seconds" % timeout, 'tell application "Microsoft Word"'] + lines + ["end tell", "end timeout"],
                timeout=timeout + 20)
        if r.returncode:
            raise ConvertError(r.stderr.strip()[:300])
        return r.stdout.strip()

    try:
        word([f'open POSIX file "{src}"'], 90)
    except ConvertError as e:
        raise ConvertError(f"Word で開けない（{e}）。" + STUCK.format(app="Microsoft Word"))
    try:
        ok = ""
        for _ in range(30):
            try:
                ok = word([f'return (name of document "{src.name}")'], 20)
                break
            except ConvertError as e:
                if is_timeout(e):                                    # まだ開いていないのではなく、Word が止まっている
                    raise ConvertError("Word が開いた文書を返さない。" + STUCK.format(app="Microsoft Word"))
                time.sleep(1)
        if not ok:
            raise ConvertError("Word で開いた写しが見つからない")
        word([f'save as document "{src.name}" file name "{pdf}" file format format PDF'], 300)
        wait_file(pdf, 30)
    finally:
        time.sleep(0.5)
        try:
            word(["try", f'close document "{src.name}" saving no', "end try"], 30)
        except ConvertError:
            pass
        try:                                                         # PDF で保存すると、文書の名前が PDF の名前に変わることがある
            word(["try", f'close document "{pdf.name}" saving no', "end try"], 30)
        except ConvertError:
            pass
    if not pdf.exists():
        raise ConvertError("Word での PDF 書き出しに失敗")
    return pdf


def excel(lines, timeout=120):
    """Excel に AppleScript を1手順だけ送り、返ってきた文字を返す。失敗は ConvertError"""
    r = osa(["with timeout of %d seconds" % timeout, 'tell application "Microsoft Excel"', "set display alerts to false"]
            + lines + ["end tell", "end timeout"], timeout=timeout + 30)
    if r.returncode:
        raise ConvertError(r.stderr.strip()[:300])
    return r.stdout.rstrip("\n")


def wait_file(path, seconds=60):
    """ファイルが書き終わる（大きさが変わらなくなる）のを待つ"""
    last = -1
    for _ in range(seconds * 2):
        if path.exists():
            size = path.stat().st_size
            if size > 0 and size == last:
                return True
            last = size
        time.sleep(0.5)
    return path.exists()


def excel_sheet_pdfs(src, out_dir):
    """Excel で開き、表示されているシートを1枚ずつ PDF にする。戻り値は [(シート番号, シート名, PDF のパス or None)]。

    Excel の「PDF で保存」はブック全体を出すので、シートを1枚ずつ新しいブックに写し、その写しを PDF にして閉じる。
    元のブックには手を加えず、最後に保存せずに閉じる。手順を1つずつ送り、PDF が書き終わるのを待ってから閉じる
    （1本のスクリプトで続けて閉じると Excel が落ちることがあった）。変換中は Excel を操作しない。"""
    src, out_dir = pathlib.Path(src), pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    name = src.name
    out = []
    try:
        try:
            excel([f'open POSIX file "{src}"'], timeout=90)
        except ConvertError as e:
            raise ConvertError(f"Excel で開けない（{e}）。" + STUCK.format(app="Microsoft Excel"))
        info = ""
        for _ in range(60):
            try:
                info = excel(["set TB to character id 9", 'set res to ""',
                              f'repeat with i from 1 to (count of sheets of workbook "{name}")',
                              f'set ws to sheet i of workbook "{name}"',
                              'set vis to "1"',
                              "try",
                              'if ((visible of ws) as string) is not "sheet visible" then set vis to "0"',
                              "end try",
                              "set res to res & i & TB & (name of ws) & TB & vis & linefeed",
                              "end repeat", "return res"])
                break
            except ConvertError as e:
                if is_timeout(e):
                    raise ConvertError("Excel が開いたブックを返さない。" + STUCK.format(app="Microsoft Excel"))
                time.sleep(1)
        if not info:
            raise ConvertError("Excel で開いたブックが見つからない")
        for line in info.splitlines():
            parts = line.split("\t")
            if len(parts) < 3:
                continue
            i, sheet, vis = int(parts[0]), parts[1], parts[2] == "1"
            pdf = out_dir / f"s{i}.pdf"
            if vis:
                try:
                    book = excel([f'copy worksheet (sheet {i} of workbook "{name}")', "return name of active workbook"])
                    if book and book != name:
                        try:
                            try:                                     # 写しの側だけ、横を1ページに収める（図が左右のページに割れないように）
                                excel(["set ps to page setup object of active sheet", "set zoom of ps to false",
                                       "set fit to pages wide of ps to 1", "set fit to pages tall of ps to 999"], timeout=30)
                            except ConvertError as e:
                                if is_timeout(e):
                                    raise
                            excel([f'save as (active sheet) filename "{pdf}" file format PDF file format'], timeout=120)
                            wait_file(pdf)
                        finally:
                            time.sleep(0.5)
                            try:
                                excel([f'close workbook "{book}" saving no'], timeout=30)
                            except ConvertError:
                                pass
                except ConvertError as e:
                    if is_timeout(e):                                # ダイアログで止まっている。残りのシートを待ち続けず、このブックを打ち切る
                        raise ConvertError(f"Excel での PDF 書き出しが止まった（シート「{sheet}」）。" + STUCK.format(app="Microsoft Excel"))
                    # 止まったのでなければ、このシートだけ画像にできなかった（呼び出し側が文字だけで残す）
            out.append((i, sheet, pdf if pdf.exists() else None))
    finally:
        time.sleep(1)
        try:
            excel(["try", f'close workbook "{name}" saving no', "end try", "set display alerts to true"])
        except ConvertError:
            pass
    return out


def to_pdf(src, out_dir, backend="auto"):
    """PowerPoint・Word の文書を1本の PDF にして返す（Excel は render_pages.py がシート単位で扱う）"""
    src = pathlib.Path(src)
    app = APP_OF_EXT.get(src.suffix.lower())
    if app is None:
        raise ConvertError(f"PDF にできない形式: {src.suffix}")
    how = pick_backend(app, backend)
    out_dir = pathlib.Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if how == "soffice":
        return soffice_pdf(src, out_dir)
    pdf = out_dir / (src.stem + ".pdf")
    return powerpoint_pdf(src, pdf) if app == "powerpoint" else word_pdf(src, pdf)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src", nargs="?", help="Office ファイル")
    ap.add_argument("out", nargs="?", help="PDF を置くフォルダ")
    ap.add_argument("--probe", action="store_true", help="この Mac で使える変換の手段を表示して終わる")
    a = ap.parse_args()
    if a.probe:
        so = find_soffice()
        print(f"LibreOffice: {so or '無し'}")
        for app in ("powerpoint", "excel", "word"):
            ok, why = office_ready(app)
            print(f"{APP_NAME[app]}: {'使える（' + why + '）' if ok else '使えない — ' + why}")
        return
    if not a.src or not a.out:
        ap.error("変換するファイルと、PDF を置くフォルダを指定する（使える手段を見るだけなら --probe）")
    src, out = pathlib.Path(a.src).resolve(), pathlib.Path(a.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    stage = stage_dir()
    safe = defang(src, stage)
    try:
        is_excel = APP_OF_EXT.get(safe.suffix.lower()) == "excel"
        if is_excel and pick_backend("excel") == "office":
            for i, name, pdf in excel_sheet_pdfs(safe, stage):
                if pdf:
                    shutil.move(str(pdf), str(out / f"{src.stem}_{i}.pdf"))
                print(i, name, "→ PDF" if pdf else "→ 画像にできない")
        else:
            pdf = soffice_pdf(safe, stage, single_page_sheets=True) if is_excel else to_pdf(safe, stage)
            shutil.move(str(pdf), str(out / f"{src.stem}.pdf"))
            print(out / f"{src.stem}.pdf")
    except ConvertError as e:
        sys.exit(str(e))
    finally:
        shutil.rmtree(stage, ignore_errors=True)


if __name__ == "__main__":
    main()
