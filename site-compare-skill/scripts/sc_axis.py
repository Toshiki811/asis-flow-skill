#!/usr/bin/env python3
"""sc_axis.py — 比べる行（業務分類 › 業務種別。axis.json）の下書き・材料・見渡し（工程 4）。

  python3 scripts/sc_axis.py draft -w <比較の作業フォルダ>       # 下書き（axis.json が無いときだけ。--force で作り直す）
  python3 scripts/sc_axis.py material -w <比較の作業フォルダ> [--group G01]   # 業務種別の対応を決める材料（拠点ごとのフローの要約）
  python3 scripts/sc_axis.py plan -w <比較の作業フォルダ>        # 行ごとの材料の大きさ（タスクの対応で 1 体に渡せるか）
  python3 scripts/sc_axis.py show -w <比較の作業フォルダ>        # 行 × 拠点の当て方と、まだどの行にも当てていないフロー

下書き: 基準あり（mode: baseline）は基準の拠点の業務分類から、基準なし（side_by_side）は 1 つ目の拠点から作る。
  フローを描く分類のうち、親にフローの無い分類を 1 行にし（下位フローは開いて含める: expand true）、フローの無い束ねる分類を業務分類にする。
  ほかの拠点は「none・理由なし」で置く（決めるまで点検が FAIL）。行の名前は、もとの分類の名前（日本語でなければ訳す）。
"""
import argparse
import json
import sys

from sc_common import MATERIAL_CAP, STEP_TYPES, is_jp, load_snap, ordered_sites, save, site_names, sites_of, try_load, work_dir
from sc_rows import Model


def draft(cmp, sites):
    base = sites.get("baseline") or ordered_sites(sites)[0]
    s = load_snap(cmp, base)
    if not s:
        sys.exit(f"{site_names(sites)[base]} の写しが無い（python3 scripts/sc_collect.py）")
    cats = {c["id"]: c for c in s["categories"]}
    flows = s["flows"]

    def ancestors(cid):
        out, seen = [], set()
        while cats.get(cid, {}).get("parent") and cid not in seen:
            seen.add(cid)
            cid = cats[cid]["parent"]
            out.append(cid)
        return out
    groups, gid_of, rows = [], {}, []
    for c in s["categories"]:
        if c["id"] not in flows or any(a in flows for a in ancestors(c["id"])):
            continue
        grp = next((a for a in ancestors(c["id"]) if a not in flows), None)
        if grp and grp not in gid_of:
            gid_of[grp] = f"G{len(groups) + 1:02d}"
            groups.append({"id": gid_of[grp], "name": cats[grp]["name"], "name_from": {"site": base, "category": grp}})
        others = [{"site": x, "relation": "none", "parts": [], "reason": ""} for x in ordered_sites(sites) if x != base]
        rows.append({"id": f"B{len(rows) + 1:02d}", "group": gid_of.get(grp), "parent": None, "name": c["name"],
                     "name_from": {"site": base, "flow": c["id"]}, "origin": "baseline" if sites.get("baseline") else "merged",
                     "sites": [{"site": base, "relation": "same", "parts": [{"flow": c["id"], "expand": True, "nodes": None}], "reason": ""}] + others})
    return {"mode": "baseline" if sites.get("baseline") else "side_by_side", "groups": groups, "rows": rows, "not_compared": [], "questions": []}


def flow_lines(mdl, sid):
    """拠点のフローの要約（業務種別の対応を決める材料）"""
    s = mdl.snaps.get(sid) or {}
    cats = {c["id"]: c for c in s.get("categories") or []}
    out = []
    for c in s.get("categories") or []:
        f = (s.get("flows") or {}).get(c["id"])
        if not f:
            out.append(f"  {c['id']} {c['name']}（フロー無し{'・対象外: ' + c['out_of_scope'] if c.get('out_of_scope') else '・束ねる分類'}）")
            continue
        steps = [k for k in f["nodes"] if s["nodes"][k]["type"] in STEP_TYPES]
        labs = [s["nodes"][k]["label"] for k in steps]
        par = f" 親 {f['parent']}" if f.get("parent") and f["parent"] in cats else ""
        out.append(f"  {c['id']} {c['name']}{par}｜開始「{f['start_condition']}」→ 終了「{f['end_condition']}」｜担当 {'・'.join(l['name'] for l in f['lanes'])}"
                   f"｜作業 {len(steps)}（{labs[0] if labs else '—'} … {labs[-1] if labs else '—'}）")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft", "material", "plan", "show"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--cap", type=int, default=MATERIAL_CAP)
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    names = site_names(sites)
    if a.cmd == "draft":
        if (cmp / "axis.json").exists() and not a.force:
            sys.exit("axis.json はもうある（作り直すときは --force。LLM が書いた当て方は消える）")
        ax = draft(cmp, sites)
        save(cmp / "axis.json", ax)
        print(f"比べる行 {len(ax['rows'])}・業務分類 {len(ax['groups'])} の下書き → {cmp / 'axis.json'}。ほかの拠点のフローを当てる（references/axis-guide.md）")
        return
    mdl = Model(cmp, sites)
    if a.cmd == "material":
        total = 0
        for sid in mdl.order:
            lines = flow_lines(mdl, sid)
            total += sum(len(x) for x in lines)
            print(f"== {names[sid]}{'【基準】' if sid == sites.get('baseline') else ''}（言葉 {next(x for x in sites['sites'] if x['id'] == sid)['language']}）")
            print("\n".join(lines))
        print(f"\n材料 {total:,} 字" + (f"（上限 {a.cap:,} 字を超える。業務分類ごとに分けて決める）" if total > a.cap else ""))
        return
    if a.cmd == "plan":
        big = 0
        for r in mdl.rows():
            sizes = []
            for sid in mdl.order:
                ks = mdl.row_nodes(r["id"], sid)[0]
                chars = sum(len(json.dumps(mdl.node(sid, k), ensure_ascii=False)) // 3 for k in ks)
                sizes.append((sid, len(ks), chars))
            tot = sum(x[2] for x in sizes)
            big += tot > a.cap
            print(f"{r['id']} {r['name']}｜" + "・".join(f"{names[s]} 箱 {n}" for s, n, _c in sizes) + f"｜材料 約 {tot:,} 字"
                  + ("｜★大きすぎる。下位フローの切れ目で子の行（B03-1…）に分ける" if tot > a.cap else ""))
        print(f"\n行 {len(mdl.rows())}・材料が上限（{a.cap:,} 字）を超える行 {big}。行 1 つ（同じフローを分け合う行はまとめて）を 1 体のサブエージェントに渡す")
        return
    placed = {}
    for r in mdl.rows():
        cells = []
        for sid in mdl.order:
            sp = mdl.site_part(r["id"], sid) or {}
            fl = [p["flow"] + ("（開く）" if p.get("expand") else "") + (f"［箱 {len(p['nodes'])}］" if p.get("nodes") else "") for p in sp.get("parts") or []]
            for p in sp.get("parts") or []:
                placed.setdefault(sid, set()).add(p["flow"])
            cells.append(f"{names[sid]}: {sp.get('relation', '—')} {' '.join(fl)}")
        print(f"{r['id']} {r['name']}{'' if is_jp(r['name']) else '（日本語でない）'}｜" + "｜".join(cells))
    for sid in mdl.order:
        s = mdl.snaps.get(sid) or {}
        cov = set()
        for f in placed.get(sid, set()):
            cov |= {f} | set((s.get("families") or {}).get(f, {}).get("children") or [])
        nc = {x["flow"] for x in mdl.axis.get("not_compared") or [] if x.get("site") == sid}
        left = [f for f in (s.get("flows") or {}) if f not in cov and f not in nc]
        if left:
            print(f"まだ当てていないフロー（{names[sid]}）: {', '.join(left)}")


if __name__ == "__main__":
    main()
