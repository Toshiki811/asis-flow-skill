#!/usr/bin/env python3
"""extract_vba.py — Office ファイルのマクロ（VBA）を、実行せずにソースとして取り出す（render_pages.py が使う）。

  python3 scripts/extract_vba.py book.xlsm            # モジュールの一覧と、ボタンに割り当てられたマクロを表示
  python3 scripts/extract_vba.py book.xlsm --dump dir # モジュールごとに .bas で書き出す

取り出すもの:
  modules   モジュール名・種別・ソース・手続き（Sub / Function）の一覧・イベント（Workbook_Open など）
  touches   ソースの中で触っている先（シート名・ほかのブック・ファイルの保存・メール・外部コマンド）の手がかり
  buttons   シート上のボタン・図形に割り当てられたマクロ（どの操作でどの処理が動くか）
  sheets    シートのコード名（VBA の中での名前）→ シート名
  shortcuts ショートカットに割り当てられたマクロ
  ribbon    リボンの定義（customUI）で、ボタンなどに割り当てられたマクロ
  book      ブックのコード名（ThisWorkbook。言語によって名前が違う）
  unread    読めていないもの（旧形式 xls のボタンの割り当て、など）。黙って「無い」にしない
  locked    VBA プロジェクトに閲覧のロックが掛かっているか（掛かっていても、ソースはファイルに平文で入っている。読むかどうかは利用者に伺う）
要るもの: pip3 install oletools
"""
import argparse
import html
import pathlib
import re
import struct
import zipfile
import xml.etree.ElementTree as ET

PROC = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)",
                  re.I | re.M)
EVENT = re.compile(r"(?i)^(Auto_?Open|Auto_?Close|Workbook_\w+|Worksheet_\w+|Document_\w+|\w+_(Click|Change|Activate|"
                   r"BeforeSave|BeforeClose|Open|Initialize|Calculate|SelectionChange|DblClick))$")
TOUCH = [
    ("シート", re.compile(r'(?i)(?:Work)?sheets\(\s*"([^"]+)"\s*\)')),
    ("ブックを開く", re.compile(r'(?i)Workbooks\.Open(?:Text)?\s*\(?\s*(?:Filename\s*:=\s*)?([^,\n)]+)')),
    ("保存", re.compile(r'(?i)\.(SaveAs|SaveCopyAs|ExportAsFixedFormat)\b')),
    ("ファイル入出力", re.compile(r'(?i)\b(Open\s+.+\s+For\s+(?:Input|Output|Append|Binary)|FileCopy|Kill|MkDir|Name\s+.+\s+As)\b')),
    ("ファイル選択", re.compile(r'(?i)\b(GetOpenFilename|GetSaveAsFilename|FileDialog)\b')),
    ("メール", re.compile(r'(?i)(Outlook\.Application|\.SendMail|CDO\.Message|MailItem)')),
    ("外部プログラム", re.compile(r'(?i)\b(Shell\s*\(|WScript\.Shell|CreateObject\(\s*"[^"]+"\s*\))')),
    ("データベース・Web", re.compile(r'(?i)(ADODB\.|QueryTables|MSXML2\.|WinHttp|ODBC;)')),
    ("印刷", re.compile(r'(?i)\.(PrintOut|PrintPreview)\b')),
    ("メッセージ", re.compile(r'(?i)\bMsgBox\s*\(?\s*"([^"]{4,})"')),
]
NS = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


SHORTCUT = re.compile(r'(?im)^Attribute\s+([^\s.]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^\s"\\])')   # 割り当てが無いと、キーの所は空白


def clean(code):
    lines = [ln.rstrip() for ln in code.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    lines = [ln for ln in lines if not re.match(r"Attribute\s+[\w.]*VB_", ln)]       # VBA の編集画面には出ない行（行の番号を、編集画面と合わせる）
    while lines and not lines[-1].strip():
        lines.pop()
    while lines and not lines[0].strip():
        lines.pop(0)
    return "\n".join(lines)


def analyse(code):
    procs = []
    for mm in PROC.finditer(code):
        name = mm.group(2)
        procs.append({"kind": re.sub(r"\s+", " ", mm.group(1)).title(), "name": name,
                      "line": code.count("\n", 0, mm.start()) + 1, "event": bool(EVENT.match(name))})
    touches = []
    for label, rx in TOUCH:
        seen = []
        for mm in rx.finditer(code):
            v = (mm.group(1) if mm.groups() else mm.group(0)).strip()
            if v not in seen:
                seen.append(v)
        if seen:
            touches.append({"what": label, "items": seen[:20]})
    return procs, touches


def is_encrypted(path):
    """ファイル全体が暗号化されているか（新形式を、パスワードで保護して保存したもの。入れ物が OLE で、中に EncryptedPackage がある）"""
    try:
        import olefile
        if zipfile.is_zipfile(path) or not olefile.isOleFile(str(path)):
            return False
        ole = olefile.OleFileIO(str(path))
        try:
            return ole.exists("EncryptedPackage")
        finally:
            ole.close()
    except Exception:
        return False


def modules_of(path):
    """(モジュールの一覧, エラー文)。マクロは実行しない（oletools はファイルを読むだけ）"""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        return [], "oletools が無い（pip3 install oletools）"
    if is_encrypted(path):
        return [], "ファイルが暗号化されている（パスワードで保護して保存してある）。中は読んでいない。保護を外した写しをもらう"
    out = []
    try:
        vp = VBA_Parser(str(path))
        if vp.detect_vba_macros():
            for _fn, stream, vba_name, code in vp.extract_macros():
                if isinstance(code, bytes):
                    code = code.decode("utf-8", "replace")
                src = clean(code)
                if not src.strip():
                    continue
                name = re.sub(r"\.(bas|cls|frm)$", "", vba_name or pathlib.PurePosixPath(stream or "module").name)
                kind = ("標準モジュール" if (vba_name or "").endswith(".bas") else
                        "フォーム" if (vba_name or "").endswith(".frm") else "クラス・シート・ブック")
                procs, touches = analyse(src)
                out.append({"name": name, "type": kind, "code": src, "procedures": procs, "touches": touches,
                            "ext": pathlib.PurePosixPath(vba_name or "").suffix.lower() or ".cls",
                            "shortcuts": [{"procedure": a, "key": b} for a, b in SHORTCUT.findall(code)]})
        vp.close()
    except Exception as e:
        return out, f"VBA を読めない: {type(e).__name__}: {e}"
    return out, ""


def xlsb_records(data):
    """バイナリ ブック（xlsb）の部品の、レコードの並び [(番号, 中身)]"""
    i, n = 0, len(data)
    while i < n:
        rid = data[i]
        i += 1
        if rid & 0x80 and i < n:
            rid = (rid & 0x7F) | ((data[i] & 0x7F) << 7)
            i += 1
        size = 0
        for k in range(4):
            if i >= n:
                return
            byte = data[i]
            i += 1
            size |= (byte & 0x7F) << (7 * k)
            if not byte & 0x80:
                break
        yield rid, data[i:i + size]
        i += size


def wide(buf, at):
    """xlsb の文字列（長さ 4 バイト＋UTF-16）→ (文字, 次の位置)。長さが 0xFFFFFFFF なら「無い」"""
    n = struct.unpack_from("<I", buf, at)[0]
    if n == 0xFFFFFFFF:
        return None, at + 4
    return buf[at + 4:at + 4 + 2 * n].decode("utf-16-le", "replace"), at + 4 + 2 * n


def sheet_parts(z):
    """[(シート名, シートの部品のパス, 状態)]"""
    names = set(z.namelist())
    out = []
    if "xl/workbook.xml" in names:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
        for s in wb.findall("m:sheets/m:sheet", NS):
            t = target.get(s.get(f"{{{NS['r']}}}id"), "")
            t = t.lstrip("/") if t.startswith("/") else "xl/" + t
            out.append((s.get("name"), t, s.get("state") or "visible"))
    elif "xl/workbook.bin" in names and "xl/_rels/workbook.bin.rels" in names:      # バイナリ ブック（xlsb）
        rels = ET.fromstring(z.read("xl/_rels/workbook.bin.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
        for rid, body in xlsb_records(z.read("xl/workbook.bin")):
            if rid == 156 and len(body) >= 12:                       # BrtBundleSh（シート1つ）
                state = struct.unpack_from("<I", body, 0)[0]
                rel, at = wide(body, 8)
                name, _ = wide(body, at)
                t = target.get(rel or "", "")
                t = t.lstrip("/") if t.startswith("/") else "xl/" + t
                out.append((name, t, {0: "visible", 1: "hidden", 2: "veryHidden"}.get(state, "visible")))
    return out


def rel_targets(z, part, suffix):
    """部品 part の関係のうち、種類が suffix で終わるものの行き先"""
    p = pathlib.PurePosixPath(part)
    rels = str(p.parent / "_rels" / (p.name + ".rels"))
    if rels not in z.namelist():
        return []
    out = []
    for r in ET.fromstring(z.read(rels)).findall("rel:Relationship", NS):
        if r.get("Type", "").endswith(suffix) and r.get("TargetMode") != "External":
            t = r.get("Target")
            if t.startswith("/"):
                out.append(t.lstrip("/"))
            else:
                parts = []
                for seg in (p.parent / t).parts:
                    if seg == "..":
                        parts.pop()
                    else:
                        parts.append(seg)
                out.append("/".join(parts))
    return out


def macro_name(val):
    """割り当ての書き方（[0]!名前・'ブック.xlsm'!名前・モジュール.名前）→ 名前（モジュール名つきは、そのまま）"""
    return html.unescape(val).split("!")[-1].strip().strip("'")


def strip_tags(text):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", text))).strip()


def buttons_of(path):
    """シート上の図形・ボタンに割り当てられたマクロ: [{sheet, shape, macro}]。shape は、ボタンに書いてある文字（無ければ図形の名前）"""
    out = []
    if not zipfile.is_zipfile(path):
        return out

    def add(sheet, shape, macro):
        item = {"sheet": sheet, "shape": shape, "macro": macro_name(macro)}
        if item["macro"] and item not in out:
            out.append(item)
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for sheet, part, _state in sheet_parts(z):
            caption = {}                                             # 図形の番号 → 書いてある文字
            for dr in rel_targets(z, part, "/drawing"):
                if dr not in names:
                    continue
                for el in ET.fromstring(z.read(dr)).iter():
                    tag = el.tag.split("}")[-1]
                    if tag not in ("sp", "pic", "cxnSp", "graphicFrame", "grpSp"):
                        continue
                    nv = el.find(".//xdr:cNvPr", NS)
                    text = "" if tag == "grpSp" else "".join(t.text or "" for t in el.iter(f"{{{NS['a']}}}t")).strip()
                    if nv is not None and text:
                        caption[nv.get("id")] = text
                    if el.get("macro"):
                        add(sheet, text or (nv.get("name") if nv is not None else ""), el.get("macro"))
            shapes = []                                              # 旧来の描画（VML）の図形: (番号, 書いてある文字, 割り当て)
            for vml in rel_targets(z, part, "/vmlDrawing"):
                if vml not in names:
                    continue
                txt = z.read(vml).decode("utf-8", "ignore")
                for mm in re.finditer(r"<v:shape\b(.*?)</v:shape>", txt, re.S):
                    sid = re.search(r'\bid="_x0000_s(\d+)"', mm.group(1))
                    box = re.search(r"<v:textbox\b[^>]*>(.*?)</v:textbox>", mm.group(1), re.S)
                    mac = re.search(r"<x:FmlaMacro>([^<]+)</x:FmlaMacro>", mm.group(1))
                    shapes.append((sid.group(1) if sid else "", strip_tags(box.group(1)) if box else "", mac.group(1) if mac else ""))
            written = {sid: text for sid, text, _mac in shapes if sid and text}
            done = set()                                             # シートの部品のほうで拾った図形（同じボタンを、2個に数えない）
            if part in names and part.endswith(".xml"):              # フォームのボタン（シートの部品に、割り当てが書いてある）
                txt = z.read(part).decode("utf-8", "ignore")
                for mm in re.finditer(r"<control\b([^>]*)>(.*?)</control>", txt, re.S):
                    mac = re.search(r'<controlPr\b[^>]*\bmacro="([^"]+)"', mm.group(2))
                    if mac:
                        sid = re.search(r'\bshapeId="(\d+)"', mm.group(1))
                        nm = re.search(r'\bname="([^"]*)"', mm.group(1))
                        sid = sid.group(1) if sid else ""
                        done.add(sid)
                        add(sheet, caption.get(sid) or written.get(sid) or (html.unescape(nm.group(1)) if nm else "フォームのボタン"), mac.group(1))
            for sid, text, mac in shapes:
                if not mac or (sid and sid in done):
                    continue
                shape = caption.get(sid) or text or "フォームのボタン"
                if not any(x["sheet"] == sheet and x["macro"] == macro_name(mac) for x in out) or shape != "フォームのボタン":
                    add(sheet, shape, mac)
            for cp in rel_targets(z, part, "/ctrlProp"):
                if cp in names:
                    macro = ET.fromstring(z.read(cp)).get("macro")
                    if macro and not any(x["sheet"] == sheet and x["macro"] == macro_name(macro) for x in out):
                        add(sheet, "フォームのボタン", macro)
    return out


def ribbon_of(path):
    """リボンの定義（customUI）で、マクロを割り当ててある部品: [{label, macro, what}]"""
    out = []
    if not zipfile.is_zipfile(path):
        return out
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if not re.match(r"(?i)customUI/[^/]+\.xml$", n):
                continue
            txt = z.read(n).decode("utf-8", "ignore")
            for mm in re.finditer(r"<(\w+)\b([^>]*)>", txt):
                tag, attrs = mm.group(1), mm.group(2)
                get = lambda k: (re.search(rf'\b{k}="([^"]*)"', attrs) or [None, ""])[1]
                for prop in ("onAction", "onLoad"):
                    if get(prop):
                        item = {"label": html.unescape(get("label") or get("id") or tag), "macro": macro_name(get(prop)),
                                "what": "load" if prop == "onLoad" else tag}
                        if item not in out:
                            out.append(item)
    return out


def xls_names(path):
    """旧形式（xls・xla）の (シートのコード名 → シート名, ブックのコード名)。読めなければ ({}, "")。
    Workbook ストリームのレコード（BOUNDSHEET・CODENAME）を読む。暗号化されていれば読まない"""
    try:
        import olefile
        if zipfile.is_zipfile(path) or not olefile.isOleFile(str(path)):
            return {}, ""
        ole = olefile.OleFileIO(str(path))
        try:
            name = next((x for x in ("Workbook", "Book") if ole.exists(x)), None)
            data = ole.openstream(name).read() if name else b""
        finally:
            ole.close()
    except Exception:
        return {}, ""

    def records(at):
        depth = 0
        while at + 4 <= len(data):
            rid, size = struct.unpack_from("<HH", data, at)
            body = data[at + 4:at + 4 + size]
            if rid == 0x0809:
                depth += 1
            yield rid, body, depth
            if rid == 0x000A:
                depth -= 1
                if depth <= 0:
                    return
            at += 4 + size

    def text(body, at, wide_len):
        n = struct.unpack_from("<H" if wide_len else "<B", body, at)[0]
        at += 2 if wide_len else 1
        if at >= len(body):
            return ""
        high = body[at] & 1
        raw = body[at + 1:at + 1 + n * (2 if high else 1)]
        return raw.decode("utf-16-le" if high else "latin-1", "replace")
    bound, book = [], ""
    for rid, body, depth in records(0):
        if rid == 0x002F:                                            # FILEPASS（暗号化）
            return {}, ""
        if rid == 0x0085 and len(body) > 7:
            bound.append((struct.unpack_from("<I", body, 0)[0], text(body, 6, False)))
        elif rid == 0x01BA and depth == 1 and len(body) > 3:
            book = text(body, 0, True)
    sheets = {}
    for pos, name in bound:
        for rid, body, depth in records(pos):
            if rid == 0x01BA and depth == 1 and len(body) > 3:
                sheets[text(body, 0, True)] = name
                break
    return sheets, book


def sheet_codenames(path):
    """シートのコード名（VBA の中での名前。Sheet1 など）→ シート名"""
    out = {}
    if not zipfile.is_zipfile(path):
        return xls_names(path)[0]
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for sheet, part, _state in sheet_parts(z):
            if part not in names:
                continue
            if part.endswith(".bin"):                                # xlsb: BrtWsProp（シートの性質）の最後がコード名
                for rid, body in xlsb_records(z.read(part)):
                    if rid == 147 and len(body) >= 23:
                        code, _ = wide(body, 19)
                        if code:
                            out[code] = sheet
                        break
                    if rid == 145:                                   # BrtBeginSheetData（ここから先はセル）
                        break
                continue
            mm = re.search(r'<sheetPr\b[^>]*\bcodeName="([^"]+)"', z.read(part).decode("utf-8", "ignore"))
            if mm:
                out[mm.group(1)] = sheet
    return out


def book_codename(path):
    """ブックのコード名（ThisWorkbook。ドイツ語版なら DieseArbeitsmappe、など）。分からなければ空"""
    if not zipfile.is_zipfile(path):
        return xls_names(path)[1]
    with zipfile.ZipFile(path) as z:
        if "xl/workbook.xml" in z.namelist():
            mm = re.search(r'<workbookPr\b[^>]*\bcodeName="([^"]+)"', z.read("xl/workbook.xml").decode("utf-8", "ignore"))
            return mm.group(1) if mm else ""
    return ""


def unread_of(path, mods):
    """読めていないもの（マクロがあるのに、形式のせいで見ていない所）"""
    out = []
    if not mods:
        return out
    if not zipfile.is_zipfile(path):
        out.append("ボタン・図形へのマクロの割り当て（旧形式 xls・xla では読めていない）")
        if not xls_names(path)[0]:
            out.append("シートの名前と、VBA の中での名前の対応")
    else:
        with zipfile.ZipFile(path) as z:
            if not sheet_parts(z):
                out.append("シートの名前・ボタンへのマクロの割り当て（ブックの部品を読めていない）")
    return out


def project_stream(path):
    """VBA プロジェクトの PROJECT ストリームの文字（無ければ空）"""
    import io
    try:
        import olefile
        data = None
        if zipfile.is_zipfile(path):
            with zipfile.ZipFile(path) as z:
                name = next((n for n in z.namelist() if n.endswith("vbaProject.bin")), None)
                data = io.BytesIO(z.read(name)) if name else None
        elif olefile.isOleFile(str(path)):
            data = str(path)
        if data is None:
            return ""
        ole = olefile.OleFileIO(data)
        try:
            hit = next((x for x in ole.listdir() if x[-1] == "PROJECT"), None)
            return ole.openstream(hit).read().decode("cp932", "replace") if hit else ""
        finally:
            ole.close()
    except Exception:
        return ""


def cmg_locked(cmg):
    """PROJECT ストリームの CMG（保護の状態。16 進の文字）→ 閲覧のロックが掛かっているか。解けなければ None。
    仕様（MS-OVBA 2.4.3 データの暗号化）のとおりに戻して、VBE の保護の印（0x4）を見る"""
    try:
        raw = bytes.fromhex(cmg)
        seed, ver_enc, key_enc = raw[0], raw[1], raw[2]
        key = seed ^ key_enc
        un1, enc1, enc2, out = key, key_enc, ver_enc, []
        for b in raw[3:]:
            x = b ^ ((enc2 + un1) & 0xFF)
            enc2, enc1, un1 = enc1, b, x
            out.append(x)
        body = out[(seed & 6) // 2:]
        size = int.from_bytes(bytes(body[:4]), "little")
        data = body[4:4 + size]
        return bool(data and data[0] & 0x4) if size == 4 else None
    except Exception:
        return None


def is_locked(path):
    """閲覧のロック（VBA プロジェクトのプロパティの「プロジェクトを表示用にロックする」）が掛かっているか。分からなければ None"""
    mm = re.search(r'(?m)^CMG="([0-9A-Fa-f]+)"', project_stream(path))
    return cmg_locked(mm.group(1)) if mm else None


def extract(path):
    mods, err = modules_of(path)
    return {"modules": mods, "buttons": buttons_of(path), "sheets": sheet_codenames(path), "locked": is_locked(path), "error": err,
            "shortcuts": [dict(s, module=m["name"]) for m in mods for s in m["shortcuts"]],
            "ribbon": ribbon_of(path), "book": book_codename(path), "unread": unread_of(path, mods)}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", help="Office ファイル（.xlsm・.docm・.pptm など）")
    ap.add_argument("--dump", metavar="DIR", help="モジュールごとに .bas で書き出すフォルダ")
    a = ap.parse_args()
    res = extract(a.file)
    if res["error"]:
        print("ERROR", res["error"])
    if res["locked"]:
        print("VBA プロジェクトに閲覧のロックが掛かっている")
    for m in res["modules"]:
        print(f"[{m['type']}] {m['name']}  {len(m['code'].splitlines())} 行")
        for p in m["procedures"]:
            print(f"    {p['kind']} {p['name']}（{p['line']} 行目）{'  ← イベント' if p['event'] else ''}")
        for t in m["touches"]:
            print(f"    {t['what']}: {', '.join(t['items'][:6])}")
    for b in res["buttons"]:
        print(f"ボタン: シート「{b['sheet']}」の「{b['shape']}」→ {b['macro']}")
    for b in res["ribbon"]:
        print(f"リボン: 「{b['label']}」→ {b['macro']}")
    for s in res["shortcuts"]:
        print(f"ショートカット: Ctrl+{'Shift+' if s['key'].isupper() else ''}{s['key'].upper()} → {s['module']}.{s['procedure']}")
    if res["sheets"]:
        print("シートのコード名: " + "、".join(f"{k}＝「{v}」" for k, v in res["sheets"].items()))
    if res["book"]:
        print(f"ブックのコード名: {res['book']}")
    for u in res["unread"]:
        print(f"※ 読めていない: {u}")
    if a.dump:
        out = pathlib.Path(a.dump)
        out.mkdir(parents=True, exist_ok=True)
        for m in res["modules"]:
            (out / f"{m['name']}.bas").write_text(m["code"], encoding="utf-8")
    if not res["modules"] and not res["error"]:
        print("マクロなし")


if __name__ == "__main__":
    main()
