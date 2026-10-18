#!/usr/bin/env python3
"""sc_verify.py — 対応の照合（工程 7）: 照合シートを作り、照合に出した版を控え、照合の報告から採否表の下書きを作る。

  python3 scripts/sc_verify.py sheet -w <比較の作業フォルダ> [--only B01,B03]          # 全体の照合シート（行ごと）。版を控え、見出しに［V-<n>］
  python3 scripts/sc_verify.py sheet -w <比較の作業フォルダ> --only B03 --claims t2/OS:absent,t4/SG   # 差分照合（照合のあとに変わった主張だけ）
  python3 scripts/sc_verify.py draft -w <比較の作業フォルダ> [--only B01]              # 報告（reports/verify_<行>*.json）から、採否表の下書き

主張（照合する単位）:
  t3/TK          東京支店の、この箱の組が、このタスク（名前）の作業である。ほかの拠点の箱と同じ作業を指している
  t3/OS:absent   大阪支店の資料に、このタスクの作業が無い（候補を見ても、同じ作業が無い。粗く書いた 1 つの箱の中にも含まれていない）
  t3/content:1   内容の差（文）が、両拠点の原典のとおり
照合シートは、主張ごとに、箱の名前・担当・札・帳票・出典・原文の引用（原典の言葉のまま）を並べる。照合の担当は、作り方を知らされない
（このスキルの決まり・タスクを決めた理由は渡さない）。判定は、照合に出した順に重ね、あとのシートの判定が前の判定を置き換える。
照合のあとに、主張の中身（タスクの名前・箱の組・記載なしの候補・内容の差）が変わると、点検（--stage verify）が止め、その主張だけを照合し直させる。
"""
import argparse
import datetime
import json
import re
import sys

from sc_common import digest, load, save, sites_of, try_load, work_dir
from sc_rows import Model
from sc_validate import validate

CLAIM = re.compile(r"^(t[0-9]+)/(?:([A-Z][A-Z0-9]{1,5})(:absent)?|content:([0-9]+))$")


def claims(mdl, rid):
    """主張の番号 → (中身の印, 見せる中身)"""
    out = {}
    for t in mdl.tasks(rid):
        mem = t.get("members") or {}
        others = {s: sorted(ks) for s, ks in mem.items() if ks}
        for s, ks in mem.items():
            if ks:
                body = {"name": t.get("name"), "site": s, "nodes": [[k, (mdl.node(s, k) or {}).get("fp"), (mdl.node(s, k) or {}).get("label")] for k in ks],
                        "others": {o: v for o, v in others.items() if o != s}}
                out[f"{t['id']}/{s}"] = (digest(body), body)
        for s, ab in (t.get("absent") or {}).items():
            body = {"name": t.get("name"), "site": s, "checked": sorted(ab.get("checked") or []), "note": ab.get("note"), "others": others}
            if ab.get("not_here"):                                   # 書いた所だけ（無い主張の印は、前と同じ）
                body["not_here"] = sorted(ab["not_here"])
            out[f"{t['id']}/{s}:absent"] = (digest(body), body)
        for i, c in enumerate(t.get("content") or [], 1):
            body = {"name": t.get("name"), "content": c}
            out[f"{t['id']}/content:{i}"] = (digest(body), body)
    return out


def versions(cmp):
    v = try_load(cmp / "review" / "versions.json")[0]
    return v if isinstance(v, dict) and isinstance(v.get("versions"), list) else {"versions": []}


def node_block(mdl, sid, k, indent="    "):
    n = mdl.node(sid, k) or {}
    lines = [f"{indent}- 「{n.get('label', '')}」（担当: {n.get('lane', '')}" + (f"・札: {'・'.join(n['systems'])}" if n.get("systems") else "")
             + (f"・帳票: {'・'.join(dict.fromkeys(f['name'] for f in n['files']))}" if n.get("files") else "") + (" ・推定" if n.get("inferred") else "") + "）"
             + (f" 補足: {n['detail']}" if n.get("detail") else "")]
    for src, q in zip(n.get("src") or [], n.get("quotes") or []):
        lines.append(f"{indent}  出典 {src}: 「{q}」")
    if not n.get("src"):
        lines.append(f"{indent}  出典なし（推定で置いた箱）")
    return lines


def item_block(mdl, sid, key, indent="    "):
    s = mdl.snaps.get(sid) or {}
    if key in (s.get("nodes") or {}):
        return node_block(mdl, sid, key, indent)
    x = next((x for x in s.get("excluded") or [] if x["key"] == key), None)
    if x:
        lines = [f"{indent}- 「{x['text']}」（フローに描かなかった手順。理由: {x['reason']}）",
                 f"{indent}  出典 {x['src']}" + (f": 「{x['quote']}」" if x.get("quote") else "")]
        lines += [f"{indent}  同じページの注記: 「{q}」" for q in x.get("notes") or []]
        return lines
    return [f"{indent}- {key}（写しに無い）"]


def sheet_text(mdl, rid, vid, only=None):
    r = mdl.row_of[rid]
    out = [f"# ［{vid}］対応の照合シート — {r['name']}", "",
           "業務マニュアルから描いた、拠点ごとの業務フローの箱を、拠点をまたいで「同じ作業」のタスクにまとめました。"
           "下の主張が、原典（出典の引用）から見て正しいかを、1 つずつ判定してください。",
           "箱の名前は日本語に直してあります（日本語でない拠点は、引用が原文です）。",
           "箱の担当・札（システム）・帳票・補足は、拠点ごとに原典と照らして確かめてあります（表の資料では、同じ行のほかの列から付けていて、引用に出ないことがあります）。"
           "この照合で見るのは、箱の組とタスクの対応・記載なし・内容の差です。担当・システム・帳票・順序の違いは、機械が箱から出すので、内容の差に書かなくてかまいません。"
           "1 つの箱は、1 つのタスクにだけ入ります。箱がほかの作業の物も名指ししていれば、主な作業のタスクに入れ、もう一方は内容の差（対象）として出ています。", ""]
    out.append("## 拠点ごとの、この業務の箱（並び順）")
    for sid in mdl.order:
        sp = mdl.site_part(rid, sid) or {}
        out.append(f"### {mdl.names[sid]}（{ {'same': '当てたフローあり', 'part': 'この業務の一部', 'split': 'ほかの業務と 1 本のフローの一部', 'none': '資料に記載なしとした', 'out_of_scope': '対象外'}.get(sp.get('relation'), '') }）")
        if sp.get("reason"):
            out.append(f"  当て方の理由: {sp['reason']}")
        for k in mdl.step_nodes(rid, sid):
            out += node_block(mdl, sid, k, "  ")
    out += ["", "## 主張"]
    cs = claims(mdl, rid)
    by = {t["id"]: t for t in mdl.tasks(rid)}
    for cid, (_mark, body) in cs.items():
        if only and cid not in only:
            continue
        m = CLAIM.match(cid)
        t = by[m.group(1)]
        if m.group(4):
            c = body["content"]
            out += [f"### {cid}　タスク「{t['name']}」の内容の差", f"  主張: {c['text']}（種類: {c['kind']}）"]
            for e in c["evidence"]:
                if "item" in e:                                      # 拠点がフローから外した項目（共通の資料の書き分けなど）
                    out.append(f"  {mdl.names.get(e['site'], e['site'])}の根拠（フローに描かなかった、資料の項目）:")
                    out += item_block(mdl, e["site"], e["item"])
                else:
                    out.append(f"  {mdl.names.get(e['site'], e['site'])}の根拠の箱:")
                    out += node_block(mdl, e["site"], e["node"])
        elif m.group(3):
            sid = m.group(2)
            ab = t["absent"][sid]
            if ab.get("not_here"):
                out += [f"### {cid}　{mdl.names[sid]}では、タスク「{t['name']}」の作業は、この拠点の手順ではない"
                        f"（資料に、ほかの拠点が行う・ほかの拠点だけの手順とある）",
                        f"  主張の理由: {ab['note']}", "  この拠点の手順ではないとした根拠:"]
                for k in ab["not_here"]:
                    out += item_block(mdl, sid, k)
                out.append("  見ること: 根拠の文から、この作業をほかの拠点が行う（この拠点の手順ではない）と読めるか。この拠点の箱（上の一覧）が、この作業をしていないか")
                out.append("  ほかの拠点の、このタスクの箱:")
            else:
                out += [f"### {cid}　{mdl.names[sid]}に、タスク「{t['name']}」の作業は書かれていない",
                        f"  主張の理由: {ab['note']}", "  ほかの拠点の、このタスクの箱:"]
            for s, ks in (t.get("members") or {}).items():
                for k in ks:
                    out += [f"    {mdl.names[s]}"] + node_block(mdl, s, k, "      ")
            out.append("  確かめた候補（字が似ている箱・描かなかった手順）:" if ab.get("checked") else "  字が似ている候補は無かった")
            for k in ab.get("checked") or []:
                out += item_block(mdl, sid, k)
            if not ab.get("not_here"):
                out.append("  見ること: 候補のほかにも、同じ作業がこの拠点の箱（上の一覧）や、粗く書いた 1 つの箱の中に含まれていないか")
        else:
            sid = m.group(2)
            out += [f"### {cid}　{mdl.names[sid]}の次の箱は、タスク「{t['name']}」の作業である"]
            for k in (t.get("members") or {}).get(sid) or []:
                out += node_block(mdl, sid, k)
            oth = [(s, ks) for s, ks in (t.get("members") or {}).items() if s != sid and ks]
            if oth:
                out.append("  同じタスクにした、ほかの拠点の箱:")
                for s, ks in oth:
                    out.append(f"    {mdl.names[s]}: " + "／".join(f"「{(mdl.node(s, k) or {}).get('label', '')}」" for k in ks))
        out.append("")
    return "\n".join(out) + "\n"


def cmd_sheet(cmp, mdl, only, claim_ids):
    vs = versions(cmp)
    rows = only or [r["id"] for r in mdl.rows() if mdl.match(r["id"])]
    if claim_ids and len(rows) != 1:
        sys.exit("--claims は、--only に行を 1 つだけ書いて使う")
    for rid in rows:
        if not mdl.match(rid):
            sys.exit(f"{rid}: タスクの対応が無い")
        cs = claims(mdl, rid)
        sel = set(claim_ids) if claim_ids else None
        if sel and sel - set(cs):
            sys.exit(f"{rid}: 主張 {', '.join(sorted(sel - set(cs)))} が無い（sc_check.py --stage verify が挙げた番号を写す）")
        vid = f"V-{len(vs['versions']) + 1}"
        vs["versions"].append({"id": vid, "date": datetime.date.today().isoformat(), "kind": "差分" if sel else "全体", "row": rid,
                               "claims": {c: m for c, (m, _b) in cs.items() if not sel or c in sel}})
        name = f"対応の照合シート_{rid}" + (f"_{vid}" if sel else "") + ".md"
        (cmp / "out").mkdir(exist_ok=True)
        (cmp / "out" / name).write_text(sheet_text(mdl, rid, vid, sel), encoding="utf-8")
        n = len(vs["versions"][-1]["claims"])
        rep = f"reports/verify_{rid}.json" if not sel else f"reports/verify_{rid}_差分{sum(1 for v in vs['versions'] if v['row'] == rid and v['kind'] == '差分')}.json"
        print(f"{rid}: 照合シート out/{name}（［{vid}］主張 {n}）→ 照合の担当の報告の置き場 {rep}")
    save(cmp / "review" / "versions.json", vs)


def reports_of(cmp, rid):
    return sorted((cmp / "reports").glob(f"verify_{rid}.json")) + sorted((cmp / "reports").glob(f"verify_{rid}_差分*.json"),
                                                                        key=lambda p: int(re.sub(r"\D", "", p.stem.split("差分")[-1]) or 0))


def cmd_draft(cmp, mdl, only):
    vs = versions(cmp)
    order = {v["id"]: i for i, v in enumerate(vs["versions"])}
    rows = only or [r["id"] for r in mdl.rows() if reports_of(cmp, r["id"])]
    for rid in rows:
        reps = []
        for p in reports_of(cmp, rid):
            d, err = try_load(p)
            errs = [err] if err else validate(d, "report_verify")
            if errs:
                sys.exit(f"{p.name}: 報告の形が崩れている（{errs[0]}）。照合の担当に直させる")
            if d["sheet"] not in order or vs["versions"][order[d["sheet"]]]["row"] != rid:
                sys.exit(f"{p.name}: シートの番号 {d['sheet']} が、この行の照合に出した記録に無い")
            shown = vs["versions"][order[d["sheet"]]]["claims"]
            extra = [v["claim"] for v in d["verdicts"] if v["claim"] not in shown]
            if extra:
                sys.exit(f"{p.name}: シートに載せていない主張を判定している（{', '.join(extra[:5])}）")
            reps.append(d)
        reps.sort(key=lambda d: order[d["sheet"]])
        verdicts = {}
        for d in reps:
            for v in d["verdicts"]:
                verdicts[v["claim"]] = dict(v, sheet=d["sheet"])
        old = try_load(cmp / "review" / f"{rid}.review.json")[0] or {"findings": []}
        prev = {(f["sheet"], f["kind"], f["target"], f["text"]): f for f in old.get("findings") or []}
        finds = []
        for d in reps:
            for f in d["findings"]:
                p = prev.get((d["sheet"], f["kind"], f["target"], f["text"]), {})
                finds.append({"id": f"f{len(finds) + 1}", "sheet": d["sheet"], "kind": f["kind"], "target": f["target"], "text": f["text"],
                              "decision": p.get("decision", ""), "action": p.get("action", "")})
        rv = {"row": rid, "verdicts": [verdicts[c] for c in sorted(verdicts)], "findings": finds}
        save(cmp / "review" / f"{rid}.review.json", rv)
        open_ = sum(1 for f in finds if not f["decision"])
        print(f"{rid}: 採否表 review/{rid}.review.json（判定 {len(verdicts)}・指摘 {len(finds)}・採否の無い指摘 {open_}）")


def review_md(cmp, mdl):
    """照合の採否表（作業の記録）: 行ごとの支持率と、指摘ごとの採否と対応"""
    out = ["# 対応の照合 採否表", "", "| 業務種別 | 主張 | 支持 | 一部支持 | 不支持 |", "| --- | --- | --- | --- | --- |"]
    finds = []
    for r in mdl.rows():
        rv = try_load(cmp / "review" / f"{r['id']}.review.json")[0] or {}
        vs = [v["verdict"] for v in rv.get("verdicts") or []]
        out.append(f"| {r['name']} | {len(vs)} | {vs.count('支持')} | {vs.count('一部支持')} | {vs.count('不支持')} |")
        finds += [(r["name"], f) for f in rv.get("findings") or []]
    out += ["", "| # | 業務種別 | 対象 | 指摘（原文のまま） | 種別 | 採否 | 対応／理由 |", "| --- | --- | --- | --- | --- | --- | --- |"]
    for i, (name, f) in enumerate(finds, 1):
        cell = lambda s: str(s).replace("|", "｜").replace("\n", " ")
        out.append(f"| {i} | {cell(name)} | {cell(f['target'])} | {cell(f['text'])} | {cell(f['kind'])} | {cell(f['decision'])} | {cell(f['action'])} |")
    if not finds:
        out.append("| — | — | — | 指摘なし | — | — | — |")
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["sheet", "draft"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only")
    ap.add_argument("--claims")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    mdl = Model(cmp, sites_of(cmp))
    only = [x.strip() for x in a.only.split(",") if x.strip()] if a.only else None
    for rid in only or []:
        if rid not in mdl.row_of:
            sys.exit(f"行 {rid} が axis.json に無い")
    if a.cmd == "sheet":
        cmd_sheet(cmp, mdl, only, [x.strip() for x in a.claims.split(",") if x.strip()] if a.claims else None)
    else:
        cmd_draft(cmp, mdl, only)


if __name__ == "__main__":
    main()
