#!/usr/bin/env python3
"""real_samples.py — 実物のサンプル（公開プロジェクトのテスト用データ）で、マクロ付き Excel・Access の書き出し・Access の直接読みを確かめ直す。

  python3 evals/real_samples.py <サンプルのフォルダ>

サンプルは、このリポジトリに入れていない。一覧（入手元・大きさ・SHA-256・ライセンス）と、期待する結果は evals/real_samples.json。
<サンプルのフォルダ> には、一覧の path のとおりに置く（excel/…・access/…・code/…。版管理アドインの書き出しは access/vcs-src/）。
このスクリプトは、取り込みをしない。置いてあるファイルだけを確かめ、無いものは「無い」と出して飛ばす。
どのファイルも、Office・Access では開かない。マクロは実行しない（読むのは、バイトと文字だけ）。
"""
import hashlib
import json
import pathlib
import re
import subprocess
import sys
import tempfile
import urllib.parse
import zipfile

SKILL = pathlib.Path(__file__).resolve().parent.parent
S = SKILL / "scripts"
sys.path.insert(0, str(S))
import access_direct as ad  # noqa: E402
import extract_access as ea  # noqa: E402
import extract_vba  # noqa: E402

OUT = []


def say(name, ok, detail=""):
    OUT.append(bool(ok))
    print(f"{'OK  ' if ok else 'NG  '} {name}" + (f"  {detail}" if detail else ""))


def files_check(root, spec):
    have = {}
    for f in spec["files"]:
        p = root / f["path"]
        if not p.is_file():
            continue
        have[f["path"]] = p
        say(f"ファイル {f['path']}", hashlib.sha256(p.read_bytes()).hexdigest() == f["sha256"], "" if p.stat().st_size == f["size"] else f"大きさ {p.stat().st_size}（一覧は {f['size']}）")
    missing = [f["path"] for f in spec["files"] if f["path"] not in have]
    if missing:
        print(f"無い（飛ばす）: {len(missing)} 件 — " + "、".join(missing[:6]) + (" …" if len(missing) > 6 else ""))
    return have


def excel_check(root, spec, have):
    for name, want in spec["expected"]["excel"].items():
        p = have.get(f"excel/{name}")
        if p is None:
            continue
        r = extract_vba.extract(str(p))
        got = {"modules": len(r["modules"]), "procedures": sum(len(m["procedures"]) for m in r["modules"]), "buttons": len(r["buttons"]), "sheets": len(r["sheets"]),
               "shortcuts": len(r["shortcuts"]), "ribbon": len(r.get("ribbon") or []), "unread": len(r.get("unread") or []), "locked": r["locked"], "error": bool(r["error"]),
               "encrypted": extract_vba.is_encrypted(str(p))}
        say(f"マクロ付き Excel {name}", got == want, "" if got == want else f"{ {k: (got[k], want[k]) for k in want if got.get(k) != want[k]} }")


def struct_of(inputs, tmp, name):
    """入力 → 棚卸し・ページ化・構造の取り出し → 構造（まとまり1つ）"""
    work = tmp / name
    for script, args in (("inventory.py", [*map(str, inputs), "-w", str(work)]), ("render_pages.py", ["-w", str(work)]), ("code_struct.py", ["-w", str(work)])):
        r = subprocess.run([sys.executable, str(S / script), *args], capture_output=True, text=True)
        if r.returncode:
            return None, (r.stdout + r.stderr)[-400:]
    import code_struct as cs
    sts = [st for st in cs.load_structs(work).values() if st.get("access")]
    return (sts[0], "") if len(sts) == 1 else (None, f"まとまりが {len(sts)} 個")


def export_check(root, spec, tmp):
    import code_struct as cs
    out = {}
    for rel, want in spec["expected"]["access_export"].items():
        src = root / "access/vcs-src" / rel
        if not src.is_dir():
            print(f"無い（飛ばす）: access/vcs-src/{rel}")
            continue
        st, why = struct_of([src], tmp, "export_" + re.sub(r"\W", "_", rel))
        if st is None:
            say(f"Access の書き出し {rel}", False, why)
            continue
        got = {"procedures": len(st["procedures"]), "entries": len(st["entries"]), "units": len(st["units"]),
               "layout_events": sum(f["events"] for f in st["access"]["forms"] + st["access"]["reports"])}
        say(f"Access の書き出し {rel}: 手続き・起動点・図にする単位・定義のイベント", got == want and cs.tree_errors(st) == [], "" if got == want else str(got))
        odd = {x["kind"] for p in st["procedures"] for x in p["io"]} - set(cs.IO_KINDS)
        say(f"Access の書き出し {rel}: 入出力の種類が、決まった語（common の io_kind）", not odd, "・".join(sorted(odd)))
        out[rel] = st
    return out


def direct_check(root, spec, tmp, exported):
    want = spec["expected"]["access_direct"]
    acc = root / "access"
    for rel, word in want["refuse"].items():
        if not (root / rel).is_file():
            continue
        try:
            ad.read(root / rel)
            say(f"直接読み: 読まない {rel}", False, "読めてしまった")
        except ad.Unreadable as e:
            say(f"直接読み: 読まない {rel}", word in str(e), str(e))
    java = acc / "jackcess/QueryTest.java"
    dbs = [p for p in (acc / "jackcess/queryTestV2000.mdb", acc / "jackcess/queryTestV2003.mdb", acc / "jackcess/queryTestV2007.accdb", acc / "jackcess/queryTestV2010.accdb") if p.is_file()]
    if java.is_file() and dbs:
        text = java.read_text(encoding="utf-8")
        block = text[text.index("expectedQueries.put("):text.index("try (Database db = TestUtil.open(testDB))")]
        expected = {}
        for chunk in block.split("expectedQueries.put(")[1:]:
            parts = re.findall(r'"((?:[^"\\]|\\.)*)"', chunk)
            expected[parts[0]] = "\n".join(p.replace('\\"', '"').replace("\\0", "").replace("\\\\", "\\") for p in parts[1:])
        ands = lambda sql: re.sub(r"ON \((.*?)\)(?= INNER| LEFT| RIGHT|\n|$)", lambda m: "ON (" + ") AND (".join(sorted(m.group(1).split(") AND ("))) + ")", sql)
        ok = total = exact = 0
        for p in dbs:
            got = {q["name"]: q["sql"] for q in ad.read(p)["queries"]}
            for name, sql in expected.items():
                total += 1
                exact += got.get(name) == sql
                ok += ands(got.get(name, "")) == ands(sql)
        say("直接読み: クエリ（Jackcess のテストの期待値。結合の条件を AND で並べる順は問わない）", [ok, total] == want["queries_jackcess"],
            f"{ok}/{total}（文字まで同じ {exact}）")
    vc_zip, src = acc / "vcs/Version_Control_v5.0.1.zip", acc / "vcs-src/Version Control.accda.src"
    if not vc_zip.is_file() or not src.is_dir():
        print("無い（飛ばす）: 版管理アドインのファイル（access/vcs/Version_Control_v5.0.1.zip）か、その書き出し（access/vcs-src/Version Control.accda.src）")
        return
    vc = tmp / "vc_in" / "Version Control.accda"
    vc.parent.mkdir()
    with zipfile.ZipFile(vc_zip) as z:
        vc.write_bytes(z.read(next(n for n in z.namelist() if n.endswith(".accda"))))
    res = ad.read(vc)
    un = lambda p: urllib.parse.unquote(p.stem)
    truth = {"form": {un(p) for p in src.glob("forms/*.form")}, "macro": {un(p) for p in src.glob("macros/*.macro")},
             "module": {un(p) for p in list(src.glob("modules/**/*.bas")) + list(src.glob("modules/**/*.cls"))},
             "query": {un(p) for p in src.glob("queries/*.sql")}, "table": {un(p) for p in src.glob("tbldefs/*.sql")}}
    got = {"form": {f["name"] for f in res["forms"]}, "macro": {m["name"] for m in res["macros"]}, "module": {m["name"] for m in res["modules"]},
           "query": {q["name"] for q in res["queries"]}, "table": {t["name"] for t in res["tables"]}}
    say("直接読み: オブジェクトの一覧（フォーム・マクロ・モジュール・クエリ・テーブルの名前）", got == truth and {k: len(v) for k, v in got.items()} == want["objects"])
    f_ok = f_all = 0
    for t in res["tables"]:
        ref = ea.parse_tabledef((src / "tbldefs" / f"{t['name']}.sql").read_text(encoding="utf-8-sig").splitlines())
        mine = [(f["name"], f["type"], f["pk"]) for f in t["fields"]]
        f_all += len(ref["fields"])
        f_ok += sum(1 for f in ref["fields"] if (f["name"], f["type"], f["pk"]) in mine)
    say("直接読み: テーブルの項目（名前・型・主キー）", [f_ok, f_all] == want["table_fields"], f"{f_ok}/{f_all}")
    squash = lambda s: re.sub(r"\[(\w+)\]", r"\1", re.sub(r"[\s;]+", "", s).lower())
    q_ok = 0
    for q in res["queries"]:
        ref = (src / "queries" / f"{q['name']}.sql").read_text(encoding="utf-8-sig")
        a, b = ea.sql_tables(q["sql"]), ea.sql_tables(ref)
        q_ok += squash(q["sql"]) == squash(ref) and (set(a[0]), set(a[1])) == (set(b[0]), set(b[1]))
    say("直接読み: クエリ（書き出しの SQL と。空白・大小・1語を囲む角かっこは問わない）", [q_ok, len(res["queries"])] == want["queries_vcs"], f"{q_ok}/{len(res['queries'])}")
    ref = {}
    for f in list(src.glob("modules/**/*.bas")) + list(src.glob("modules/**/*.cls")) + list(src.glob("forms/*.cls")):
        ref[("Form_" if f.parent.name == "forms" else "") + un(f)] = f.read_text(encoding="utf-8-sig").replace("\r\n", "\n")

    def norm(text):
        ls = text.split("\n")
        if ls and ls[0].startswith("VERSION 1.0 CLASS"):             # 書き出すときに付く見出し
            ls = ls[next(i for i, x in enumerate(ls) if x.strip() == "END") + 1:]
        ls = [x.rstrip().lower() for x in ls if not x.startswith("Attribute ")]
        while ls and not ls[-1]:
            ls.pop()
        return ls
    code = {m["name"]: m["code"] for m in res["modules"]}
    for f in res["forms"]:
        if "CodeBehindForm" in f["lines"]:
            code["Form_" + f["name"]] = "\n".join(f["lines"][f["lines"].index("CodeBehindForm") + 1:])
    same = sum(1 for name, text in ref.items() if norm(code.get(name, "")) == norm(text))
    say("直接読み: VBA のソース（行ごと。Attribute の行と、識別子の英字の大小は問わない）", [same, len(ref)] == want["vba_modules"], f"{same}/{len(ref)} モジュール")
    c_ok = c_all = e_ok = e_all = 0
    nl = lambda t: t.replace("\r\n", "\n")
    for f in res["forms"]:
        lay = ea.parse_layout(f["lines"])
        other = ea.parse_layout((src / "forms" / f"{f['name']}.form").read_text(encoding="utf-8-sig").splitlines())
        mine = {(c["name"], c["type"], nl(c["caption"]), nl(c["label"])) for c in lay["controls"]}
        theirs = {(c["name"], c["type"], nl(c["caption"]), nl(c["label"])) for c in other["controls"]}
        c_all += len(theirs)
        c_ok += len(mine & theirs)
        ev = [(e["owner"], e["prop"], e["value"]) for e in lay["events"]]
        for e in other["events"]:
            e_all += 1
            e_ok += (e["owner"], e["prop"], e["value"]) in ev
    say("直接読み: フォームのコントロール（名前・種類・表示の文字・付いたラベル）", [c_ok, c_all] == want["form_controls"], f"{c_ok}/{c_all}")
    say("直接読み: フォームのイベントの割り当て", [e_ok, e_all] == want["form_events"], f"{e_ok}/{e_all}")
    xml = lambda rows: "".join(v[5:] for r in rows for nm, v, _k in r["props"] if nm == "Comment" and v.startswith("_AXL:"))
    mine = xml(ea.macro_rows(res["macros"][0]["lines"])[0]) if res["macros"] else ""
    say("直接読み: マクロの XML の定義", bool(mine) and mine == xml(ea.macro_rows((src / "macros/autoexec.macro").read_text(encoding="utf-8-sig").splitlines())[0]))
    st, why = struct_of([vc.parent], tmp, "direct_vc")
    base = exported.get("Version Control.accda.src")
    if st is None or base is None:
        say("直接読み: 通しで、書き出しと同じ構造になる", False, why or "書き出しの側の構造が無い")
        return
    key = lambda s: ({p["id"] for p in s["procedures"]}, {(e["procedure"], e["label"]) for e in s["entries"]}, {u["procedure"] for u in s["units"]})
    say("直接読み: 通しで、書き出しと同じ構造になる（手続き・起動点とその名前・図にする単位）", key(st) == key(base) and st["access"].get("via") == "direct",
        f"手続き {len(st['procedures'])}・起動点 {len(st['entries'])}・図にする単位 {len(st['units'])}")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    root = pathlib.Path(sys.argv[1]).expanduser().resolve()
    spec = json.loads((SKILL / "evals" / "real_samples.json").read_text(encoding="utf-8"))
    have = files_check(root, spec)
    excel_check(root, spec, have)
    with tempfile.TemporaryDirectory() as t:
        tmp = pathlib.Path(t)
        exported = export_check(root, spec, tmp)
        direct_check(root, spec, tmp, exported)
    print(f"\n{sum(OUT)} / {len(OUT)} 項目" + ("" if all(OUT) else "  NG あり"))
    sys.exit(0 if OUT and all(OUT) else 1)


if __name__ == "__main__":
    main()
