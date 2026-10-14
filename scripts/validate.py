#!/usr/bin/env python3
"""validate.py — LLM が書く JSON の形を、スキーマ（assets/schemas/*.schema.json）で確かめる。

  python3 scripts/validate.py -w <作業フォルダ>                              # 作業フォルダにあるものを全部
  python3 scripts/validate.py -w <作業フォルダ> --kind card --only M01 --pages 13-24   # 自分の束だけ
  python3 scripts/validate.py -w <作業フォルダ> --kind flow --only F03
  python3 scripts/validate.py --file card <ファイル>...                       # ファイルを直接
  python3 scripts/validate.py --doc                                         # references/schemas.md の項目表を作り直す
  python3 scripts/validate.py --check-doc                                   # 項目表がスキーマと合っているか

見るもの: 型・必須の項目・決まった語（enum）・書式（pattern）・**余分な項目**（スキーマに無い項目は違反）。
同じ中身を人によって別の項目に書くと、あとの工程が拾えず、実行のたびに結果が変わる。書く場所は1つに決め、機械で確かめる。
check_extract.py・check_flow.py も同じ検査を通す（違反は FAIL）。外部ライブラリは使わない。
終了コード: 違反があれば 1。
"""
import argparse
import json
import pathlib
import re
import sys

from common import ASSETS, SKILL_DIR, parse_pages

SCHEMA_DIR = ASSETS / "schemas"
KINDS = ["card", "hierarchy", "taxonomy", "glossary", "systems", "code_tree", "flow", "edits", "activity_edits", "activity", "code_terms", "report_read", "report_flow",
         "report_activity", "report_review", "report_code_review", "review", "times", "answers", "questions", "columns", "layout"]
DOC = SKILL_DIR / "references" / "schemas.md"
TYPES = {"object": dict, "array": list, "string": str, "boolean": bool}
TYPE_JA = {"object": "オブジェクト", "array": "配列", "string": "文字", "boolean": "true / false", "integer": "整数",
           "number": "数", "null": "null"}
_cache = {}


def schema(kind):
    if kind not in _cache:
        p = SCHEMA_DIR / f"{kind}.schema.json"
        if not p.exists():
            sys.exit(f"スキーマが無い: {p}")
        _cache[kind] = json.loads(p.read_text(encoding="utf-8"))
    return _cache[kind]


def resolve(sch, root):
    """$ref をたどる。「#/$defs/x」は同じファイル、「common#/$defs/x」は common.schema.json"""
    while "$ref" in sch:
        file, _, ptr = sch["$ref"].partition("#")
        node = schema(file) if file else root
        if file:
            root = node
        for part in [x for x in ptr.split("/") if x]:
            node = node[part]
        sch = dict(node, **{k: v for k, v in sch.items() if k not in ("$ref",)})
    return sch, root


def type_ok(v, t):
    if t == "null":
        return v is None
    if t == "integer":
        return isinstance(v, int) and not isinstance(v, bool)
    if t == "number":
        return isinstance(v, (int, float)) and not isinstance(v, bool)
    return isinstance(v, TYPES[t])


def type_name(v):
    for t in ("null", "boolean", "integer", "number", "string", "array", "object"):
        if type_ok(v, t):
            return TYPE_JA[t]
    return type(v).__name__


def walk(v, sch, root, path, errs):
    sch, root = resolve(sch, root)
    at = path or "（全体）"
    if "type" in sch:
        ts = sch["type"] if isinstance(sch["type"], list) else [sch["type"]]
        if not any(type_ok(v, t) for t in ts):
            errs.append(f"{at}: 型が違う（{type_name(v)}。{' / '.join(TYPE_JA[t] for t in ts)} で書く）")
            return
    if "enum" in sch and v not in sch["enum"]:
        errs.append(f"{at}: 「{str(v)[:30]}」は決まった語でない（{' / '.join(map(str, sch['enum']))}）")
        return
    if isinstance(v, str):
        if "minLength" in sch and len(v.strip()) < sch["minLength"]:
            errs.append(f"{at}: 空")
        if "pattern" in sch and not re.search(sch["pattern"], v):
            errs.append(f"{at}: 「{v[:30]}」は書式が違う（{sch['pattern']}）")
    if "minimum" in sch and isinstance(v, (int, float)) and not isinstance(v, bool) and v < sch["minimum"]:
        errs.append(f"{at}: {v} は小さすぎる（{sch['minimum']} 以上）")
    if isinstance(v, list):
        if "minItems" in sch and len(v) < sch["minItems"]:
            errs.append(f"{at}: {len(v)} 件（{sch['minItems']} 件以上要る）")
        if "maxItems" in sch and len(v) > sch["maxItems"]:
            errs.append(f"{at}: {len(v)} 件（{sch['maxItems']} 件まで）")
        if "items" in sch:
            for i, x in enumerate(v):
                walk(x, sch["items"], root, f"{path}[{i}]", errs)
    if isinstance(v, dict):
        props = sch.get("properties") or {}
        lack = [k for k in sch.get("required") or [] if k not in v]
        if lack:
            errs.append(f"{at}: 必須の項目が無い: {', '.join(lack)}")
        if sch.get("additionalProperties") is False:
            extra = [k for k in v if k not in props]
            if extra:
                errs.append(f"{at}: スキーマに無い項目: {', '.join(extra)}（消すか、決まった項目に書く。references/schemas.md）")
        more = sch.get("additionalProperties")
        for k, x in v.items():
            if k in props:
                walk(x, props[k], root, f"{path}.{k}" if path else k, errs)
            elif isinstance(more, dict):                           # 名前が決まっていない項目（地図）の値の形
                walk(x, more, root, f"{path}.{k}" if path else k, errs)


def validate(data, kind):
    """違反の一覧（文字列）。無ければ空"""
    errs = []
    root = schema(kind)
    walk(data, root, root, "", errs)
    return errs


def validate_file(path, kind):
    try:
        data = json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
    except Exception as e:
        return [f"JSON として読めない（{e}）"]
    return validate(data, kind)


def blocking(work, kinds=("flow", "taxonomy", "hierarchy", "glossary", "systems")):
    """図・報告書を作るスクリプト（と、flow.json を書き換える renumber・apply_edits・activity）が、作り始める前に確かめる。必須の項目の欠け・型の違い・
    読めない JSON があれば、どのファイルのどこかを言って止まる（KeyError の1行で終わらせない。用語の正本・台帳を「無い」として進まない。形の点検そのものは check_flow.py の役目）"""
    bad = [(f, e) for kind in kinds for f in files_of(work, kind) for e in validate_file(f, kind)
           if "必須の項目が無い" in e or "型が違う" in e or "JSON として読めない" in e]
    if bad:
        for f, e in bad[:20]:
            print(f"{f.relative_to(work)}: {e}")
        sys.exit(f"形が崩れたファイルがある（{len(bad)} か所）。先に python3 scripts/check_flow.py -w <作業フォルダ> を通す")


def files_of(work, kind):
    work = pathlib.Path(work)
    pat = {"card": "cards/*/p*.json", "hierarchy": "structure/hierarchy.json", "taxonomy": "taxonomy.json",
           "glossary": "glossary.json", "systems": "structure/systems.json", "code_tree": "code/*.tree.json", "activity": "activities/*.activity.json", "activity_edits": "activities/*.edits.json", "code_terms": "structure/code_terms.json", "flow": "flows/*.flow.json", "edits": "flows/*.edits.json", "report_read": "reports/read_*.json",
           "report_flow": "reports/flow_*.json", "report_activity": "reports/activity_*.json", "report_review": "reports/review_*.json",
           "report_code_review": "reports/code_review_*.json", "review": "review/*.review.json", "times": "times.json", "answers": "answers.json", "questions": "questions.json",
           "columns": "out/*.columns.json", "layout": "out/*.layout.json"}[kind]
    return sorted(work.glob(pat))


# ---------- references/schemas.md の項目表 ----------

def type_text(sch, root):
    sch, root = resolve(sch, root)
    if "enum" in sch:
        return "決まった語"
    t = sch.get("type")
    if isinstance(t, list):
        return " / ".join(TYPE_JA[x] for x in t)
    if t == "array":
        it, _ = resolve(sch.get("items") or {}, root)
        inner = "決まった語" if "enum" in it else TYPE_JA.get(it.get("type"), "") if isinstance(it.get("type"), str) else ""
        return f"配列（{inner}）" if inner and inner != "オブジェクト" else "配列"
    return TYPE_JA.get(t, "")


def table(sch, root):
    sch, root = resolve(sch, root)
    req = set(sch.get("required") or [])
    out = ["| 項目 | 必須 | 型 | 中身 |", "| --- | --- | --- | --- |"]
    for k, p in (sch.get("properties") or {}).items():
        d = p.get("description") or resolve(p, root)[0].get("description") or ""
        r = resolve(p, root)[0]
        enum = r.get("enum") or resolve(r.get("items") or {}, root)[0].get("enum") or []
        words = [x for x in enum if isinstance(x, str) and f"`{x}`" not in d and f"「{x}」" not in d]
        if words:                                                    # 決まった語は、説明に出てこなければ表に並べる（表だけ見て書けるように）
            d = f"{d}（{'／'.join(f'`{x}`' for x in enum if isinstance(x, str))}）" if d else "／".join(f"`{x}`" for x in enum if isinstance(x, str))
        out.append(f"| `{k}` | {'○' if k in req else ''} | {type_text(p, root)} | {d.replace('|', '｜')} |")
    return out


def doc_block(kind):
    root = schema(kind)
    out = [f"<!-- BEGIN schema:{kind}（validate.py --doc が書く。手で直さない） -->"]
    if root.get("properties"):
        out += [""] + table(root, root)
    for name, d in (root.get("$defs") or {}).items():
        if d.get("properties") and (kind != "common" or name == "question"):
            out += ["", f"**{d.get('title', name)}**", ""] + table(d, root)
    return out + ["", f"<!-- END schema:{kind} -->"]


def render_doc(text):
    for kind in KINDS + ["common", "manifest", "page"]:
        pat = re.compile(rf"<!-- BEGIN schema:{kind}（.*?<!-- END schema:{kind} -->", re.S)
        if not pat.search(text):
            sys.exit(f"schemas.md に {kind} の表の囲み（<!-- BEGIN schema:{kind} … END -->）が無い")
        text = pat.sub(lambda _m: "\n".join(doc_block(kind)), text)
    return text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work")
    ap.add_argument("--kind", choices=KINDS, help="見る種類（省くと全部）")
    ap.add_argument("--only", help="資料 ID かフロー ID（カンマ区切り）")
    ap.add_argument("--pages", help="カードのページの範囲（例: 13-24 や 3,5,8-10）")
    ap.add_argument("--file", nargs="+", metavar=("種類", "ファイル"), help="ファイルを直接見る: --file card a.json b.json")
    ap.add_argument("--doc", action="store_true")
    ap.add_argument("--check-doc", action="store_true")
    a = ap.parse_args()
    if a.doc or a.check_doc:
        old = DOC.read_text(encoding="utf-8")
        new = render_doc(old)
        if a.check_doc:
            print("schemas.md はスキーマと合っている" if old == new else "schemas.md がスキーマと合っていない（validate.py --doc を実行する）")
            sys.exit(0 if old == new else 1)
        DOC.write_text(new, encoding="utf-8")
        print(f"項目表を書いた → {DOC}")
        return
    targets = []
    if a.file:
        if a.file[0] not in KINDS + ["manifest", "page"] or len(a.file) < 2:
            sys.exit(f"--file <種類> <ファイル>...（種類: {' / '.join(KINDS + ['manifest', 'page'])}）")
        targets = [(a.file[0], pathlib.Path(f)) for f in a.file[1:]]
    elif a.work:
        if not pathlib.Path(a.work).expanduser().is_dir():
            sys.exit(f"作業フォルダが無い: {a.work}")
        only = {x.strip() for x in a.only.split(",") if x.strip()} if a.only else None
        pages = parse_pages(a.pages)
        for kind in ([a.kind] if a.kind else KINDS):
            for f in files_of(a.work, kind):
                if kind == "card":
                    if only and f.parent.name not in only:
                        continue
                    if pages is not None and int(f.stem[1:]) not in pages:
                        continue
                elif kind in ("flow", "edits") and only and f.name.split(".")[0] not in only:
                    continue
                targets.append((kind, f))
    else:
        ap.error("-w か --file か --doc を指定する")
    bad = 0
    for kind, f in targets:
        errs = validate_file(f, kind)
        bad += bool(errs)
        for e in errs:
            print(f"FAIL  {f.parent.name}/{f.name}: {e}")
    print(f"{len(targets)} ファイルを見た。形の違反があるのは {bad} ファイル")
    if not targets and a.work and (a.kind or a.only or a.pages):     # 絞り込んだのに何にも当たらない。確かめたつもりで、何も見ていない
        sys.exit("指定（--kind・--only・--pages）に当たるファイルが無い。資料 ID・フロー ID・ページの範囲・保存先を確かめる")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
