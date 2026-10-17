#!/usr/bin/env python3
"""sc_validate.py — 比較の作業フォルダの JSON の形を、スキーマ（assets/schemas/sc_*.schema.json）で確かめる。

  python3 scripts/sc_validate.py -w <比較の作業フォルダ>                 # 全部
  python3 scripts/sc_validate.py -w <比較の作業フォルダ> --kind match     # 種類を絞る
  python3 scripts/sc_validate.py --file match rows/B01.match.json        # ファイルを直接
  python3 scripts/sc_validate.py --doc                                   # references/schemas.md の項目表を作り直す（--check-doc は確かめるだけ）

検証器は asis-flow-skill の validate.py のもの（余分な項目・型違い・決まった語でない は違反）。比較のスキーマを先に読ませるので、
asis の定義（common#/$defs/flow_id など）も引ける。
"""
import argparse
import json
import pathlib
import re
import sys

from sc_common import ASSETS, SKILL_DIR

import validate as V                                                 # asis の検証器

KINDS = ["sites", "applies", "snap", "terms", "axis", "match", "result", "numbers", "report_match", "report_verify", "review", "versions", "answers"]   # answers は asis と同じ形
FILES = {"sites": "sites.json", "applies": "applies.json", "snap": "snap/*.json", "terms": "terms.json", "axis": "axis.json",
         "match": "rows/*.match.json", "result": "rows/*.result.json", "numbers": "numbers.json", "report_match": "reports/match_*.json",
         "report_verify": "reports/verify_*.json", "review": "review/*.review.json", "versions": "review/versions.json", "answers": "answers.json"}
DOC = SKILL_DIR / "references" / "schemas.md"


def preload():
    """比較のスキーマを、asis の検証器の読み置きに入れる（名前は sc_<種類>）"""
    for p in sorted((ASSETS / "schemas").glob("sc_*.schema.json")):
        V._cache[p.name[:-len(".schema.json")]] = json.loads(p.read_text(encoding="utf-8"))


preload()


def kind_name(kind):
    return f"sc_{kind}"


def validate(data, kind):
    """違反の一覧（文字列）。無ければ空（answers.json は asis と同じ形）"""
    return V.validate(data, "answers" if kind == "answers" else kind_name(kind))


def validate_file(path, kind):
    try:
        data = json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return ["ファイルが無い"]
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        return [f"JSON として読めない（{e}）"]
    return validate(data, kind)


def files_of(cmp, kind):
    return sorted(pathlib.Path(cmp).glob(FILES[kind]))


def render_doc(text):
    for kind in ["common"] + [k for k in KINDS if k != "answers"]:
        name = kind_name(kind)
        pat = re.compile(rf"<!-- BEGIN schema:{name}（.*?<!-- END schema:{name} -->", re.S)
        if not pat.search(text):
            sys.exit(f"schemas.md に {name} の表の囲み（<!-- BEGIN schema:{name} … END -->）が無い")
        text = pat.sub(lambda _m: "\n".join(V.doc_block(name)), text)
    return text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work")
    ap.add_argument("--kind", choices=KINDS)
    ap.add_argument("--file", nargs="+", metavar=("種類", "ファイル"))
    ap.add_argument("--doc", action="store_true")
    ap.add_argument("--check-doc", action="store_true")
    a = ap.parse_args()
    if a.doc or a.check_doc:
        old = DOC.read_text(encoding="utf-8")
        new = render_doc(old)
        if a.check_doc:
            print("schemas.md はスキーマと合っている" if old == new else "schemas.md がスキーマと合っていない（sc_validate.py --doc を実行する）")
            sys.exit(0 if old == new else 1)
        DOC.write_text(new, encoding="utf-8")
        print(f"項目表を書いた → {DOC}")
        return
    if a.file:
        if a.file[0] not in KINDS or len(a.file) < 2:
            sys.exit(f"--file <種類> <ファイル>...（種類: {' / '.join(KINDS)}）")
        targets = [(a.file[0], pathlib.Path(f)) for f in a.file[1:]]
    elif a.work:
        targets = [(k, f) for k in ([a.kind] if a.kind else KINDS) for f in files_of(a.work, k)]
    else:
        ap.error("-w か --file か --doc を指定する")
    bad = 0
    for kind, f in targets:
        errs = validate_file(f, kind)
        bad += bool(errs)
        for e in errs:
            print(f"FAIL  {f.parent.name}/{f.name}: {e}")
    print(f"{len(targets)} ファイルを見た。形の違反があるのは {bad} ファイル")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
