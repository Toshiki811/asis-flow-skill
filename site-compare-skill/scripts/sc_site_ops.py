#!/usr/bin/env python3
"""sc_site_ops.py — 拠点のフロー化（asis の工程 6）で、共通の資料のうち、ほかの拠点だけの手順を外す直しを出す。

  python3 scripts/sc_site_ops.py -w <比較の作業フォルダ> --site OS               # その拠点の全部
  python3 scripts/sc_site_ops.py -w <比較の作業フォルダ> --site OS --flow F02    # その分類のページの分だけ

applies.json の決めに従い、その拠点に当てはまらない項目ごとに、asis の直しの一覧の 1 行
  {"op": "exclude", "item": "M01.2.s3", "reason": "ほかの拠点だけの手順（大阪支店では行わない）"}
を出す（資料 ID は、その拠点の作業フォルダのもの）。フロー化の担当は、これを直しの一覧（flows/<分類>.edits.json）の ops に貼る。
外した分岐の行き先は、つなぎ直す（その拠点の側の手順へ。references/site-guide.md）。
"""
import argparse
import json
import sys

from sc_common import OTHER_SITE, load, site_ids, site_names, site_work, sites_of, try_load, work_dir
from sc_applies import applies_to


def reason(item, sid, names):
    ev = str(item.get("evidence") or "").strip()
    ev = f"。書き分け: {ev}" if ev else ""                            # 原典照合の担当が、外した理由を原典で確かめられるように（本部では別の担当が行う、など）
    if item["scope"] == "only":
        return f"{OTHER_SITE}（{'・'.join(names[s] for s in item['sites'])}だけ{ev}）"
    return f"{OTHER_SITE}（{names[sid]}では行わない{ev}）"


def ops_for(cmp, sites, sid, flow=None):
    """[(直しの 1 行, 項目の文)]。未定の項目があれば止まる"""
    ap = try_load(cmp / "applies.json")[0] or {"items": []}
    names = site_names(sites)
    w = site_work(cmp, sites, sid)
    m = load(w / "manifest.json")
    by_sha = {d["sha1"]: d["doc_id"] for d in m["docs"] if not d.get("duplicate_of")}
    pages = None
    if flow:
        tax = load(w / "taxonomy.json")
        c = next((c for c in tax["categories"] if c["id"] == flow), None)
        if c is None:
            sys.exit(f"{sid}: 分類 {flow} が taxonomy.json に無い")
        pages = {(ps["doc_id"], p) for ps in c.get("pages") or [] for p in ps.get("pages") or []}
    out = []
    for it in ap["items"]:
        ok = applies_to(it, sid)
        if ok is None:
            sys.exit(f"applies.json に未定の項目がある（{it['key']}）。先に、当てはまる拠点を決める")
        if ok or it["doc"] not in by_sha:
            continue
        _d, page, item = it["key"].split(".")
        doc = by_sha[it["doc"]]
        if pages is not None and (doc, int(page)) not in pages:
            continue
        out.append(({"op": "exclude", "item": f"{doc}.{page}.{item}", "reason": reason(it, sid, names)}, it["text"]))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--site", required=True)
    ap.add_argument("--flow")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    if a.site not in site_ids(sites):
        sys.exit(f"拠点 {a.site} が sites.json に無い")
    ops = ops_for(cmp, sites, a.site, a.flow)
    if not ops:
        print("外す項目は無い")
        return
    print(f"# {site_names(sites)[a.site]}: 直しの一覧の ops に貼る（{len(ops)} 件）")
    for op, text in ops:
        print(json.dumps(op, ensure_ascii=False) + f"    # {text}")


if __name__ == "__main__":
    main()
