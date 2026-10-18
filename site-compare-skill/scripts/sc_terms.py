#!/usr/bin/env python3
"""sc_terms.py — 拠点間の用語の対応（terms.json）の下書きと、見渡し（工程 3）。

  python3 scripts/sc_terms.py draft -w <比較の作業フォルダ>    # 拠点の写しに出る用語を、名前が同じものだけまとめて terms.json に足す
  python3 scripts/sc_terms.py list -w <比較の作業フォルダ>     # 用語ごとに、拠点・使われている箱の数・例（LLM がまとめるときの材料）

比べる用語: 箱のレーン（担当。システムのレーンはシステム）・箱の札（使うシステム）・箱に付いた帳票。拠点の写し（snap/）から拾う。
下書きは、名前が同じ（表記をそろえて同じ）ものだけを 1 つの組にし（basis: same_name）、ほかは 1 つずつの組にする。
LLM は、訳・言い換えで同じものを、理由（reason）つきで 1 つの組にまとめ（translation・equivalent）、担当・帳票の組の名前を日本語にする。
決めた組には、下書きをやり直しても触らない（まだどの組にも無い用語だけを、1 つずつの組で足す）。
"""
import argparse
import sys

from sc_common import load_snap, norm, ordered_sites, save, site_named, site_names, sites_of, try_load, work_dir

PREFIX = {"actor": "R", "system": "Y", "document": "D"}


def used_terms(cmp, sites):
    """拠点ごとに使われている用語: {(拠点, 種類, 名前): 箱の数}（拠点の並びは、基準が先頭）"""
    out = {}
    for sid in ordered_sites(sites):
        s = load_snap(cmp, sid)
        if not s:
            continue
        for n in s["nodes"].values():
            if n["type"] in ("start", "end"):
                continue
            if n.get("lane"):
                kind = "system" if n.get("lane_kind") == "system" else "actor"
                out[(sid, kind, n["lane"])] = out.get((sid, kind, n["lane"]), 0) + 1
            for y in n.get("systems") or []:
                out[(sid, "system", y)] = out.get((sid, "system", y), 0) + 1
            for f in n.get("files") or []:
                out[(sid, "document", f["name"])] = out.get((sid, "document", f["name"]), 0) + 1
    return out


def group_of(terms):
    """(拠点, 種類, 名前) → 組。terms.json の組から"""
    out = {}
    for g in (terms or {}).get("groups") or []:
        for m in g.get("members") or []:
            out[(m.get("site"), g.get("kind"), m.get("term"))] = g
    return out


def next_id(groups, kind):
    used = [int(g["id"][1:]) for g in groups if g.get("id", "").startswith(PREFIX[kind]) and g["id"][1:].isdigit()]
    return f"{PREFIX[kind]}{(max(used) if used else 0) + 1:02d}"


def draft(cmp, sites):
    old, err = try_load(cmp / "terms.json")
    if err:
        sys.exit(f"terms.json が読めない（{err}）")
    data = old or {"groups": [], "questions": []}
    have = group_of(data)
    used = used_terms(cmp, sites)
    added = 0
    for kind in ("actor", "system", "document"):
        names = {}
        for (sid, k, name) in used:
            if k == kind and (sid, kind, name) not in have:
                names.setdefault(norm(name), []).append((sid, name))
        for _key, mem in names.items():
            g = {"id": next_id(data["groups"], kind), "kind": kind, "name": mem[0][1], "members": [{"site": s, "term": n} for s, n in mem],
                 "basis": "same_name", "reason": ""}
            owner = {site_named(n, s, sites) for s, n in mem} if kind == "actor" else {None}
            if len(owner) == 1 and None not in owner:                # ほかの拠点の担当の名前（「本部拠点」「本部の固定資産管理責任者」）: 案。理由は LLM が書く
                g["site"] = owner.pop()
            data["groups"].append(g)
            added += 1
    save(cmp / "terms.json", data)
    return added, len(data["groups"])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft", "list"])
    ap.add_argument("-w", "--work", required=True)
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    if a.cmd == "draft":
        added, total = draft(cmp, sites)
        print(f"用語の組 {total}（新しく足した {added}）→ {cmp / 'terms.json'}。訳・言い換えで同じものを、理由つきでまとめる（references/terms-guide.md）")
        return
    names = site_names(sites)
    used = used_terms(cmp, sites)
    gof = group_of(try_load(cmp / "terms.json")[0])
    for kind in ("actor", "system", "document"):
        print(f"== {dict(actor='担当', system='システム', document='帳票')[kind]}")
        for (sid, k, name), n in sorted(used.items(), key=lambda x: (x[0][1], ordered_sites(sites).index(x[0][0]), x[0][2])):
            if k == kind:
                g = gof.get((sid, kind, name))
                own = f"（{names[g['site']]}の担当）" if g and g.get("site") else ""
                print(f"  {names[sid]:<10} {name:<24} 箱 {n:>3}  → {g['id'] + ' ' + g['name'] + own if g else '（どの組にも無い）'}")


if __name__ == "__main__":
    main()
