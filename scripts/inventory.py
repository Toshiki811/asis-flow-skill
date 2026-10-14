#!/usr/bin/env python3
"""inventory.py — 入力の棚卸し。読む前に「何が・どれだけ・どんな形で」あるかを manifest.json に書く。

  python3 scripts/inventory.py <入力フォルダ or ファイル>... -w <作業フォルダ>

資料ごとに: 資料ID（M01…）・形式・ページ数（開かずに分かる範囲）・ハッシュ・フォルダ階層・マクロの有無・
PDF に文字層があるか・同一ファイルの重複・版違いらしい組。
作り直しても資料IDは変わらない（パスで引き継ぐ）。消えた資料の ID は、あとから足した資料に使い回さない。
中身が変わった資料は「要再処理」に戻し、前の中身を読んだカード・読取の報告を stale/<日時>/ に退避する（消さない）。
入力が見つからない・前の資料が全部消える、という結果になるときは、manifest.json を書き換えずに止まる（--force で書き換える）。
"""
import argparse
import datetime
import hashlib
import pathlib
import re
import shutil
import sys
import zipfile

import extract_access
from common import load, save, work_dir
from pages import BUDGET, FIXED_TOKENS, PAGE_TOKENS
from validate import validate

KINDS = {
    "pptx": {".pptx", ".pptm", ".ppsx", ".ppsm", ".potx"},
    "pdf": {".pdf"},
    "xlsx": {".xlsx", ".xlsm", ".xltx", ".xltm"},
    "docx": {".docx", ".docm"},
    "image": {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"},
    "text": {".txt", ".md", ".csv", ".tsv"},
    # プログラムのソース（文字として読む。実行しない）。VBA の書き出し（.bas・.cls・.frm）を含む
    "code": {".bas", ".cls", ".frm", ".vbs", ".wsf", ".vb", ".py", ".sql", ".bat", ".cmd", ".ps1", ".psm1", ".sh", ".js", ".gs", ".ts", ".java", ".cs",
             ".cbl", ".cob", ".cpy", ".php", ".rb", ".pl", ".kt", ".go", ".c", ".cpp", ".h", ".jcl"},
    # シートを開かずに、マクロ（VBA）だけを読む Excel の形式（バイナリ ブック・アドイン）
    "macrobook": {".xlsb", ".xlam", ".xla"},
    # Access。書き出したテキストで受け取るのが基本（references/access-export.md）
    "access": {".accdb", ".mdb", ".accde", ".mde", ".accda", ".mda"},
    # 旧形式は Office で開いて新形式に直してから読む（render_pages.py）
    "legacy": {".ppt", ".pps", ".xls", ".doc"},
}
SKIP_NAMES = {".DS_Store", "Thumbs.db", "desktop.ini", extract_access.TMP_NAME}
VERSION_TOKENS = re.compile(
    r"(?i)([_\-\s（(]*(v(er)?\.?\s*\d+(\.\d+)*|第?\d+版|rev\.?\s*\d+|\d{8}|\d{6}|\d{4}[-_.]\d{1,2}[-_.]\d{1,2}|"
    r"old|new|最新|旧|改訂|修正|コピー|copy|final|fix|確定|draft|案)[)）]*)+")


def kind_of(path):
    ext = path.suffix.lower()
    if ext in (".txt", ".form", ".report", ".mac", ".macro") and extract_access.is_export(path):   # Access を書き出したテキスト（フォーム・レポート・マクロ・定義）
        return "code"
    if extract_access.vcs_root(path) is not None:                    # 版管理アドインで書き出したフォルダの中（読むファイルだけが、ここまで来る）
        return "code"
    for k, exts in KINDS.items():
        if ext in exts:
            return k
    return "unsupported"


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def probe(path, kind):
    """開かずに（または軽く開いて）分かること: ページ数・マクロ・文字層"""
    info = {"pages": None, "has_macro": False, "pdf_text_layer": None, "note": ""}
    try:
        if kind in ("pptx", "xlsx", "docx"):
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
                info["has_macro"] = any(n.endswith("vbaProject.bin") for n in names)
                if kind == "pptx":
                    info["pages"] = sum(1 for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
                elif kind == "xlsx":
                    wb = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                    info["pages"] = len(re.findall(r"<sheet\b", wb))      # シート数（印刷ページ数は描画後に確定）
        elif kind == "pdf":
            import fitz
            d = fitz.open(path)
            info["pages"] = d.page_count
            n = min(d.page_count, 5)
            chars = sum(len(d[i].get_text().strip()) for i in range(n))
            info["pdf_text_layer"] = chars >= 20 * max(n, 1)
            if d.needs_pass:
                info["note"] = "パスワード付き"
            d.close()
        elif kind == "image":
            info["pages"] = 1
        elif kind == "text":
            info["pages"] = 1
        elif kind == "code":
            idx = extract_access.log_index(str(path.parent))
            info["note"] = ("Access の書き出し（ファイルを直接読んだもの）" if idx and idx.get("via") == "direct" else
                            "Access の書き出し" if idx or extract_access.is_export(path) else "プログラムのソース")
        elif kind == "macrobook":
            info["has_macro"] = True
            info["note"] = "マクロ（VBA）だけを読む。シートは読まない"
        elif kind == "access":
            info["note"] = "Access のファイル。書き出したテキスト（references/access-export.md）があれば、そちらを読む。無ければ、直接読む（読める範囲だけ）"
        elif kind == "legacy":
            info["note"] = "旧形式。Office で新形式に直してから読む"
    except Exception as e:                                           # 壊れたファイル・暗号化
        info["note"] = f"開けない: {type(e).__name__}: {e}"
    return info


def version_key(path):
    """版違いを束ねるための名前（版・日付・「コピー」などを落とした幹）"""
    stem = VERSION_TOKENS.sub("", path.stem)
    return re.sub(r"[\s_\-（）()]+", "", stem).lower() + path.suffix.lower()


SKIPPED = []                                                         # Access の書き出しのフォルダで、読まなかったファイル（数を知らせる）


def collect(inputs):
    files, seen = [], set()
    SKIPPED.clear()
    for arg in inputs:
        root = pathlib.Path(arg).expanduser().resolve()
        if not root.exists():
            sys.exit(f"入力が見つからない: {root}（パスを確かめる。manifest.json は書き換えていない）")
        found = [root] if root.is_file() else sorted(root.rglob("*"))
        base = root.parent if root.is_file() else root
        for p in found:
            if not p.is_file() or p.name in SKIP_NAMES or p.name.startswith(("~$", ".")):
                continue
            if any(part.startswith(".") for part in p.relative_to(base).parts[:-1]):
                continue
            if extract_access.skip_in_export(p):                     # Access の書き出しのフォルダの、読まないファイル（画像・テーマ・テーブルの行の中身・設定）
                SKIPPED.append(p)
                continue
            if p not in seen:                                        # フォルダと、その中のファイルを両方渡しても、1回だけ数える
                seen.add(p)
                files.append((base, p))
    return files


DIRECT = []                                                          # 直接読んだ Access のファイル: (名前, 結果の1行)


def direct_read(work, inputs):
    """入力に、書き出したテキストの無い Access のファイルがあれば、ファイルを直接読んで、書き出しと同じ形のテキストを <作業フォルダ>/direct/ に作る。
    → 足す入力（direct のフォルダ。作ったものが無ければ空）。読めないファイル（暗号化・古い形式）は、そのまま（render_pages.py が、書き出しの手順を案内する）"""
    import access_direct
    DIRECT.clear()
    files = [p for _root, p in collect(inputs)]
    done = set()                                                     # 書き出したテキストがある Access のファイルの名前
    for p in files:
        if p.name.lower() in ("access_export.txt", "vbe-project.json"):
            try:
                done.add(extract_access.parse_log(extract_access.read_head(p, 4_000_000).splitlines())["source"].lower())
            except OSError:
                pass
    home = work / "direct"
    keep = set()
    for p in files:
        if p.suffix.lower() not in KINDS["access"] or p.name.lower() in done:
            continue
        out = home / f"{p.stem}_direct"
        k = 1
        while out.name in keep:                                      # 同じ名前の Access のファイルが、別のフォルダにある
            k += 1
            out = home / f"{p.stem}_direct{k}"
        log = extract_access.log_index.__wrapped__(str(out))
        if log and log.get("via") == "direct" and access_direct.fingerprint(p) in (out / "access_export.txt").read_text(encoding="utf-8"):
            keep.add(out.name)                                       # 前に読んだときから、中身が変わっていない
            DIRECT.append((p.name, f"前に直接読んだものを使う → {out}"))
            continue
        try:
            res = access_direct.write_export(p, out)
        except access_direct.Unreadable as e:
            DIRECT.append((p.name, f"直接は読めない（{e}）"))
            continue
        except Exception as e:                                       # 形式に公式の説明が無いので、思わぬ形のファイルがありうる。1件の失敗で止めない
            DIRECT.append((p.name, f"直接は読めない（{type(e).__name__}: {e}）"))
            continue
        keep.add(out.name)
        DIRECT.append((p.name, f"直接読んだ: テーブル {len(res['tables'])}・クエリ {len(res['queries'])}・フォーム {len(res['forms'])}・レポート {len(res['reports'])}・"
                       f"マクロ {len(res['macros'])}・モジュール {len(res['modules'])}／読めていないもの {len(res['failed']) + sum(1 for f in res['forms'] + res['reports'] if not f['layout'])} → {out}"))
    if home.is_dir():                                                # 入力から消えたファイルの分を片づける（このスクリプトが作ったフォルダだけ）
        for d in home.iterdir():
            if d.is_dir() and d.name not in keep and (d / "access_export.txt").is_file() and "via\tdirect" in (d / "access_export.txt").read_text(encoding="utf-8", errors="replace"):
                shutil.rmtree(d)
    extract_access.log_index.cache_clear()
    return [str(home)] if keep else []


ASK_SCOPE_OVER = 300                                                 # これを超えるページ数なら、読む範囲と優先順を利用者と決めてから読む
KEEP = ("rendered", "page_list", "pages", "render_note", "outline", "sections", "sheets", "hidden_sheets", "hidden_slides",
        "vba_modules", "vba_buttons", "vba_sheets", "vba_shortcuts", "vba_locked", "vba_ribbon", "vba_book", "vba_unread", "code", "access", "read_via", "read_direct", "tool")                                # 中身が変わっていない資料で引き継ぐもの（render_pages.py が書いた項目）


def build_docs(inputs, old):
    """入力を調べて資料の一覧を作る。前の manifest があれば、資料 ID と、変わっていない資料の描画結果を引き継ぐ"""
    old_by_path = {d["abs_path"]: d for d in old["docs"]}
    used = {d["doc_id"] for d in old["docs"]} | set(old.get("retired") or [])     # 消えた資料の ID も使い回さない（カード・出典は ID で引く）

    def next_id():
        i = 1
        while f"M{i:02d}" in used:
            i += 1
        used.add(f"M{i:02d}")
        return f"M{i:02d}"

    docs = []
    for root, p in collect(inputs):
        kind = kind_of(p)
        h = sha1(p)
        prev = old_by_path.get(str(p))
        d = {
            "doc_id": prev["doc_id"] if prev else next_id(),
            "name": p.name,
            "abs_path": str(p),
            "rel_path": str(p.relative_to(root)),
            "folder": list(p.relative_to(root).parts[:-1]),
            "kind": kind,
            "ext": p.suffix.lower(),
            "size": p.stat().st_size,
            "sha1": h,
            "version_key": version_key(p),
        }
        d.update(probe(p, kind) if kind != "unsupported" else
                 {"pages": None, "has_macro": False, "pdf_text_layer": None, "note": "未対応の形式"})
        if prev and prev.get("sha1") == h:
            d.update({k: prev[k] for k in KEEP if k in prev})
        else:
            d["rendered"] = False
            d["page_list"] = []
        docs.append(d)
    for d in docs:                                                   # 同じ中身のファイルは1件だけを読む。すでに読んである資料があれば、それを元にする
        same = [x for x in docs if x["sha1"] == d["sha1"]] if not d["note"].startswith("Access の書き出し") else [d]   # 書き出しの中の同じ中身（コードの無いフォームなど）は、別のもの
        first = next((x for x in same if x.get("rendered") and not (old_by_path.get(x["abs_path"]) or {}).get("duplicate_of")), same[0])
        d["duplicate_of"] = None if d is first else first["doc_id"]
    groups = {}                                                      # 同じフォルダで、版の表記を除くと同じ名前になる資料（版違いらしい組）
    for d in docs:
        if not d["duplicate_of"]:
            groups.setdefault((tuple(d["folder"]), d["version_key"]), []).append(d["doc_id"])
    for ids in groups.values():
        if len(ids) > 1:
            for d in docs:
                if d["doc_id"] in ids:
                    d["version_group"] = ids
    return docs


def set_aside(work, old, docs, stamp):
    """前の回から、中身が変わった資料・消えた資料の、古いカードと読取の報告を stale/<日時>/ に退避する。
    残しておくと「カード済み」として読み飛ばされ、別の中身のカードが新しいページのものとして残る。退避した結果を [(資料, 何が起きたか, 枚数)] で返す"""
    now = {d["abs_path"]: d for d in docs}
    out = []
    for o in old["docs"]:
        n = now.get(o["abs_path"])
        what = "消えた" if n is None else "中身が変わった" if n.get("sha1") != o.get("sha1") else None
        if what is None:
            continue
        moved = 0
        for src in [work / "cards" / o["doc_id"]] + sorted((work / "reports").glob(f"read_{o['doc_id']}_*.json")):
            if src.exists():
                dst = work / "stale" / stamp / src.relative_to(work)
                dst.parent.mkdir(parents=True, exist_ok=True)
                moved += len(list(src.glob("p*.json"))) if src.is_dir() else 0
                shutil.move(str(src), str(dst))
        out.append((o, what, moved))
    return out


def print_changes(old, docs, aside, stamp):
    added = [d for d in docs if d["abs_path"] not in {o["abs_path"] for o in old["docs"]}]
    if not old["docs"] or not (aside or added):
        return
    print(f"\n前の回との違い: 足された {len(added)} 件・中身が変わった {sum(w == '中身が変わった' for _o, w, _n in aside)} 件・"
          f"消えた {sum(w == '消えた' for _o, w, _n in aside)} 件")
    for o, what, moved in aside:
        print(f"  {o['doc_id']} {o['rel_path']}: {what}" + (f"（古いカード {moved} 枚を stale/{stamp}/ に退避した）" if moved else ""))
    if aside:
        print("  この資料を出典にしたフロー・割り当てた分類は、ページ化と読取をやり直したあとに見直す（python3 scripts/status.py で残りが出る）")


def print_summary(docs):
    """利用者に範囲と優先順を確認するための要約: 種類ごとの件数、読取の見込み、気をつける資料"""
    by_kind = {}
    for d in docs:
        k = by_kind.setdefault(d["kind"], {"n": 0, "pages": 0, "unknown": 0})
        k["n"] += 1
        if d["pages"] is None:
            k["unknown"] += 1
        else:
            k["pages"] += d["pages"]
    print(f"\n資料 {len(docs)} 件")
    for k, v in sorted(by_kind.items()):
        unit = "シート" if k == "xlsx" else "ページ"
        tail = f"（ページ数未確定 {v['unknown']} 件）" if v["unknown"] else ""
        print(f"  {k:<12}{v['n']:>4} 件  {v['pages']:>5} {unit}{tail}")
    known = sum(d["pages"] or 0 for d in docs if not d["duplicate_of"])
    per = max(1, (BUDGET - FIXED_TOKENS) // PAGE_TOKENS)                # 1束（＝1つのエージェントの文脈に収まる量）に入るページ数
    n_batch = max(1, -(-known // per))
    tokens = known * PAGE_TOKENS + n_batch * FIXED_TOKENS
    print(f"読む対象は 約 {known} ページ（重複を除く。Excel は描画後に増えることがある）。")
    print(f"読取の見込み: 約 {n_batch} 束（1束 約 {per} ページ・予算 {BUDGET:,} トークン。1束を1つのエージェントが読む）・"
          f"合計 約 {tokens / 1e6:.1f} 百万トークン。同時に 4〜6 束ずつ進める。")
    print("  ページ化のあと、python3 scripts/pages.py batches が実際のページ（字の小ささ・文字の量）から束を切り直す")
    if known > ASK_SCOPE_OVER:
        print(f"  ※ {ASK_SCOPE_OVER} ページを超える。読む範囲と優先順を利用者と決めてから読む（階層の判定を先に済ませ、範囲を絞る）")


def print_notes(docs):
    notes = [
        ("マクロあり", [d for d in docs if d["has_macro"]]),
        ("プログラムのソース（実行しない。アクティビティ図にする）", [d for d in docs if d["kind"] == "code"]),
        ("Access のファイル", [d for d in docs if d["kind"] == "access"]),
        ("文字層のない PDF（画像から読む）", [d for d in docs if d["pdf_text_layer"] is False]),
        ("同一ファイルの重複（1回だけ読む）", [d for d in docs if d["duplicate_of"]]),
        ("版違いらしい組（どれが現行かを確認事項に出す）", [d for d in docs if d.get("version_group")]),
        ("旧形式（変換して読む）", [d for d in docs if d["kind"] == "legacy"]),
        ("未対応・開けない", [d for d in docs if d["kind"] == "unsupported" or d["note"].startswith("開けない")]),
    ]
    if DIRECT:
        print("\nAccess のファイルの直接読み（書き出したテキストが入力に無いもの。読めた範囲だけ。読めていないものは、確認事項に自動で載る）:")
        for name, line in DIRECT:
            print(f"  {name}: {line}")
    if SKIPPED:
        print(f"\nAccess の書き出しのフォルダの中で、読まなかったファイル: {len(SKIPPED)} 件（画像・テーマ・メニュー・テーブルの行の中身・設定。資料に数えていない）")
    for title, ds in notes:
        if ds:
            print(f"\n{title}: {len(ds)} 件")
            for d in ds[:20]:
                extra = f" = {d['duplicate_of']}" if d["duplicate_of"] else ""
                print(f"  {d['doc_id']} {d['rel_path']}{extra}")
            if len(ds) > 20:
                print(f"  … ほか {len(ds) - 20} 件")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--force", action="store_true", help="資料が1件も見つからない・前の資料が全部消える結果でも、manifest.json を書き換える")
    ap.add_argument("--tool", default="", metavar="資料ID,…", help="道具のブック（マクロ付きの Excel で、シートに手順が書かれていないもの）。"
                    "シートは参考にするだけで、カードにしない（マクロは、いつもどおり読む）。利用者に確かめてから付ける")
    ap.add_argument("--untool", default="", metavar="資料ID,…", help="--tool を取り消す")
    ap.add_argument("--no-direct", action="store_true", help="Access のファイルを直接読まない（書き出したテキストだけを読む）")
    a = ap.parse_args()
    work = work_dir(a.work)
    old = load(work / "manifest.json", {"docs": []})
    extra = [] if a.no_direct else direct_read(work, a.inputs)
    docs = build_docs(list(a.inputs) + extra, old)
    kept = {d["abs_path"] for d in docs} & {o["abs_path"] for o in old["docs"]}
    if not a.force and (not docs or (old["docs"] and not kept)):
        sys.exit(f"資料が {len(docs)} 件で、前の回の資料 {len(old['docs'])} 件が1つも残らない。入力のパスを確かめる"
                 f"（manifest.json は書き換えていない。これで合っているなら --force を付ける）")
    for ids, on in ((a.tool, True), (a.untool, False)):
        want = {x.strip() for x in ids.split(",") if x.strip()}
        miss = want - {d["doc_id"] for d in docs}
        if miss:
            sys.exit(f"--tool・--untool の資料 ID が見つからない: {', '.join(sorted(miss))}（manifest.json は書き換えていない）")
        for d in docs:
            if d["doc_id"] in want:
                if on:
                    d["tool"] = True
                else:
                    d.pop("tool", None)
    now = datetime.datetime.now().isoformat(timespec="seconds")
    stamp = now.replace(":", "").replace("-", "")
    if old["docs"]:
        shutil.copy(work / "manifest.json", work / "manifest.json.bak")
    aside = set_aside(work, old, docs, stamp)
    gone = sorted(set(old.get("retired") or []) | {o["doc_id"] for o in old["docs"]} - {d["doc_id"] for d in docs})
    m = {"version": 1, "created": old.get("created") or now, "updated": now,
         "inputs": [str(pathlib.Path(x).expanduser().resolve()) for x in a.inputs], "docs": docs, "retired": gone}
    save(work / "manifest.json", m)
    print(f"manifest → {work / 'manifest.json'}")
    print_changes(old, docs, aside, stamp)
    print_summary(docs)
    errs = validate(m, "manifest")
    for e in errs:
        print(f"※ manifest.json の形が崩れている（スクリプトの不具合の疑い）: {e}")
    print_notes(docs)
    if errs:                                                         # 形の崩れた棚卸しのまま先へ進まない
        sys.exit(1)


if __name__ == "__main__":
    main()
