#!/usr/bin/env python3
"""sc_collect.py — 拠点の asis の作業フォルダから、比較に使う写し（snap/<拠点>.json）を取る（工程 2）。

  python3 scripts/sc_collect.py -w <比較の作業フォルダ>                # 全拠点
  python3 scripts/sc_collect.py -w <比較の作業フォルダ> --site OS
  python3 scripts/sc_collect.py -w <比較の作業フォルダ> --no-gate      # 拠点の門を見ない（試験・下見だけ。門の通らない写しは、納品の門が止める）

拠点の門: その拠点の asis の点検（check_flow.py --stage flow と --stage review。原典照合まで）が FAIL 0 でなければ、写さない。
写すもの: 分類の木・フロー（箱ごとに、種類・レーン・名前・補足・札・付いた帳票・出典の表示「資料名 p.N」・推定の印・指紋）・
フローから外した項目（記載なしの候補・not_here・内容の差の根拠に使う。原文の引用と、同じページの拠点の名前の出る注記つき）・用語。
開いたフロー（families）: 下位フローへのリンク（その分類の子孫へのものだけ）を、子のフローの中身で置き換えた 1 本。子のふつうの終了だけを
親の次へつなぐ（ほかの終了は、終わりのまま）。同じ子を 2 回指せば、2 回目はリンクのまま。ほかの業務へ続くリンクは、そのまま。
箱の鍵は「F03-1:P-004」。拠点を直して番号が振り直されても、出典の指紋（種類と、引いたカードの項目）が 1 つに決まる箱は、
比較のファイル（axis.json・rows/*.match.json）の鍵を、機械が付け直す（中身は変えない）。決まらない箱は、点検が挙げる。
"""
import argparse
import datetime
import json
import pathlib
import re
import subprocess
import sys

from sc_common import (ASIS_SCRIPTS, A, digest, load, match_path, mentions, node_key, save, site_ids, site_work, sites_of, snap_path,
                       split_key, try_load, work_dir)

GATE = (("check_flow.py", "--stage", "flow"), ("check_flow.py", "--stage", "review"))


def gate(w):
    """拠点の門: [(点検, FAIL の数, 動かなかった理由)]"""
    out = []
    for script, *args in GATE:
        r = subprocess.run([sys.executable, str(ASIS_SCRIPTS / script), "-w", str(w), *args], capture_output=True, text=True)
        mt = re.search(r"(\d+) FAIL / (\d+) WARN", r.stdout)
        out.append((" ".join([script] + args), int(mt.group(1)) if mt else 1, None if mt else (r.stderr.strip().splitlines() or ["出力なし"])[-1][:80]))
    return out


def descendants(cats):
    """分類 → その子孫の分類の集合"""
    kids = {}
    for c in cats:
        if c.get("parent"):
            kids.setdefault(c["parent"], []).append(c["id"])
    out = {}

    def walk(cid, seen):
        acc = set()
        for k in kids.get(cid, []):
            if k not in seen:
                acc |= {k} | walk(k, seen | {k})
        return acc
    for c in cats:
        out[c["id"]] = walk(c["id"], {c["id"]})
    return out


def expand(flows, desc, fid):
    """開いたフロー: {"nodes": [鍵], "edges": [[元, 先, ラベル, 推定]], "children": [開いた子のフロー], "boundary": [開かなかったリンクの箱]}"""
    used, notes = set(), []

    def one(f, top):
        fl = flows[f]
        nodes = [k for k in fl["nodes"]]
        edges = [list(e) for e in fl["edges"]]
        for k in list(fl["nodes"]):
            n = fl["node"][k]
            c = n.get("link_flow")
            if n["type"] != "subprocess" or c not in desc.get(f, set()) or c not in flows:
                continue
            if c in used:
                notes.append(f"{k} は、すでに開いた子のフロー {c} を、もう一度指す（リンクのまま残した）")
                continue
            used.add(c)
            sub = one(c, False)
            cs, ce = node_key(c, "start"), node_key(c, "end")
            firsts = [e[1] for e in sub["edges"] if e[0] == cs]
            lasts = [e for e in sub["edges"] if e[1] == ce]
            ins = [e for e in edges if e[1] == k]
            outs = [e for e in edges if e[0] == k]
            edges = [e for e in edges if k not in (e[0], e[1])]
            inner = [e for e in sub["edges"] if e[0] != cs and e[1] != ce]
            edges += [[a[0], b, a[2], a[3]] for a in ins for b in firsts]
            edges += [[x[0], y[1], x[2] or y[2], bool(x[3] or y[3])] for x in lasts for y in outs]
            edges += inner
            at = nodes.index(k)
            nodes[at:at + 1] = [x for x in sub["nodes"] if x not in (cs, ce)]
        return {"nodes": nodes, "edges": edges}
    top = one(fid, True)
    return {"nodes": top["nodes"], "edges": top["edges"], "children": sorted(used), "notes": notes}


def fingerprint(n, items):
    if items:
        return digest([n["type"], sorted(f"{x['doc']}|{x['page']}|{x['item']}" for x in items)])
    return digest([n["type"], A.norm(n.get("label", "")), n.get("lane", "")])


def collect(cmp, sites, sid):
    w = site_work(cmp, sites, sid)
    m = load(w / "manifest.json")
    lab = A.Labels(w, m)
    sha = {d["doc_id"]: d.get("sha1", "") for d in m["docs"]}
    tax = load(w / "taxonomy.json")
    gl = load(w / "glossary.json", {"terms": []})
    cats = [{k: c.get(k) for k in ("id", "number", "name", "parent", "level", "make_flow", "out_of_scope")} for c in tax["categories"]]
    desc = descendants(tax["categories"])
    nodes, flows, excluded = {}, {}, []
    items_text = {}
    for fl in A.load_flows(w):
        fid = fl["flow_id"]
        lanes = {l["id"]: l for l in fl.get("lanes") or []}
        files = {}
        for n in fl["nodes"]:
            if n["type"] == "file" and n.get("attach_to"):
                files.setdefault(n["attach_to"], []).append({"name": n["label"], "direction": n.get("direction", "in")})
        keys, node_of = [], {}
        for n in fl["nodes"]:
            if n["type"] == "file":
                continue
            k = node_key(fid, n["id"])
            items = [{"doc": sha.get(s["doc_id"], ""), "page": s["page"], "item": s["item"], "key": f"{s['doc_id']}.{s['page']}.{s['item']}"}
                     for s in n.get("sources") or [] if s.get("item")]
            srcs = list(dict.fromkeys(lab.src(s) for s in n.get("sources") or []))
            ln = lanes.get(n.get("lane"), {})
            rec = {"type": n["type"], "lane": ln.get("name", ""), "lane_kind": ln.get("kind", ""), "label": n.get("label", ""),
                   "detail": n.get("detail", ""), "action": n.get("action", ""), "systems": n.get("systems") or [], "files": files.get(n["id"], []),
                   "src": srcs, "quotes": [s.get("quote", "") for s in n.get("sources") or []], "items": items,
                   "inferred": bool(n.get("inferred")), "link_flow": n.get("link_flow", ""), "flow": fid}
            rec["fp"] = fingerprint(dict(n, lane=rec["lane"]), items)
            nodes[k] = rec
            keys.append(k)
            node_of[n["id"]] = k
        flows[fid] = {"name": fl["name"], "number": fl.get("number", ""), "parent": fl.get("parent"), "level": fl.get("level", 1),
                      "start_condition": fl.get("start_condition", ""), "end_condition": fl.get("end_condition", ""),
                      "lanes": [{"name": l["name"], "kind": l["kind"]} for l in fl.get("lanes") or []],
                      "nodes": keys, "node": {k: nodes[k] for k in keys},
                      "edges": [[node_of[e["from"]], node_of[e["to"]], e.get("label", ""), bool(e.get("inferred"))]
                                for e in fl.get("edges") or [] if e["from"] in node_of and e["to"] in node_of]}
        for x in fl.get("excluded_items") or []:
            c = A.load_card(w, x["doc_id"], x["page"]) or {}
            it = next((y for y in (c.get("steps") or []) + (c.get("decisions") or []) if y.get("id") == x.get("item")), {})
            notes = [n.get("quote") or n.get("text") for n in c.get("notes") or []          # 同じページの、拠点の名前の出る注記（書き分けの原文のことが多い）
                     if mentions(" ".join(str(v or "") for v in (n.get("text"), n.get("quote"))), sites)]
            excluded.append({"key": f"{x['doc_id']}.{x['page']}.{x['item']}", "flow": fid, "text": it.get("action") or it.get("question") or "",
                             "reason": x.get("reason", ""), "src": lab.src({"doc_id": x["doc_id"], "page": x["page"]}), "doc": sha.get(x["doc_id"], ""),
                             "quote": it.get("quote", ""), "notes": [q for q in notes if q]})
    families = {}
    for fid in flows:
        families[fid] = expand(flows, desc, fid)
    for f in flows.values():
        del f["node"]
    snap = {"site": sid, "categories": cats, "flows": flows, "families": families, "nodes": nodes, "excluded": excluded,
            "terms": [{"kind": t["kind"], "name": t["name"]} for t in gl.get("terms") or []],
            "questions": sum(len(fl.get("questions") or []) for fl in A.load_flows(w))}
    snap["stamp"] = digest(snap, 16)
    snap["made"] = datetime.datetime.now().isoformat(timespec="seconds")
    return snap


def rekey(cmp, sid, old, new):
    """箱の番号が振り直されたとき、指紋が 1 つに決まる箱の鍵を、比較のファイルで付け直す。→ (付け直した数, 決まらない鍵)"""
    if not old:
        return 0, []
    gone = [k for k in old["nodes"] if k not in new["nodes"]]
    by_fp = {}
    for k, n in new["nodes"].items():
        by_fp.setdefault(n["fp"], []).append(k)
    old_fp = {}
    for k, n in old["nodes"].items():
        old_fp.setdefault(n["fp"], []).append(k)
    mp, lost = {}, []
    for k in gone:
        fp = old["nodes"][k]["fp"]
        cand = [x for x in by_fp.get(fp, []) if x not in old["nodes"] or old["nodes"][x]["fp"] != fp]
        if len(cand) == 1 and len(old_fp.get(fp, [])) == 1:
            mp[k] = cand[0]
        else:
            lost.append(k)
    if not mp:
        return 0, lost
    fix = lambda k: mp.get(k, k)
    ax = try_load(cmp / "axis.json")[0]
    if isinstance(ax, dict):
        for r in ax.get("rows") or []:
            for s in r.get("sites") or []:
                if s.get("site") == sid:
                    for p in s.get("parts") or []:
                        if isinstance(p.get("nodes"), list):
                            p["nodes"] = [fix(k) for k in p["nodes"]]
        for x in ax.get("not_compared") or []:
            if x.get("site") == sid and isinstance(x.get("nodes"), list):
                x["nodes"] = [fix(k) for k in x["nodes"]]
        save(cmp / "axis.json", ax)
    for f in sorted((cmp / "rows").glob("*.match.json")):
        mt = try_load(f)[0]
        if not isinstance(mt, dict):
            continue
        for t in mt.get("tasks") or []:
            if isinstance(t.get("members"), dict) and isinstance(t["members"].get(sid), list):
                t["members"][sid] = [fix(k) for k in t["members"][sid]]
            ab = (t.get("absent") or {}).get(sid)
            if isinstance(ab, dict):
                for f2 in ("checked", "not_here"):
                    if isinstance(ab.get(f2), list):
                        ab[f2] = [fix(k) for k in ab[f2]]
            for c in t.get("content") or []:
                for e in c.get("evidence") or []:
                    if e.get("site") == sid and "node" in e:                 # 外した項目（item）の根拠は、箱の番号と関わらない
                        e["node"] = fix(e["node"])
        save(f, mt)
    return len(mp), lost


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--site")
    ap.add_argument("--no-gate", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    todo = [a.site] if a.site else site_ids(sites)
    bad = 0
    for sid in todo:
        if sid not in site_ids(sites):
            sys.exit(f"拠点 {sid} が sites.json に無い")
        w = site_work(cmp, sites, sid)
        if not (w / "manifest.json").exists():
            print(f"{sid}: asis の作業フォルダができていない（{w}）。拠点の as-is を先に（references/site-guide.md）")
            bad += 1
            continue
        if not a.no_gate:
            g = [x for x in gate(w) if x[1]]
            if g:
                print(f"{sid}: 拠点の門を通らない（" + "・".join(f"{n} {f} FAIL" + (f"〔{why}〕" if why else "") for n, f, why in g)
                      + "）。拠点の as-is を原典照合まで済ませてから写す")
                bad += 1
                continue
        old = try_load(snap_path(cmp, sid))[0]
        new = collect(cmp, sites, sid)
        n, lost = rekey(cmp, sid, old if isinstance(old, dict) else None, new)
        (cmp / "snap").mkdir(exist_ok=True)
        if isinstance(old, dict) and old.get("stamp") == new["stamp"]:
            print(f"{sid}: 写しは変わらない（フロー {len(new['flows'])}・箱 {len(new['nodes'])}）")
            continue
        save(snap_path(cmp, sid), new)
        print(f"{sid}: 写した（フロー {len(new['flows'])}・箱 {len(new['nodes'])}・外した項目 {len(new['excluded'])}）"
              + (f"。番号の振り直しで、比較のファイルの鍵を {n} 件付け直した" if n else "")
              + (f"。付け直せない箱 {len(lost)}（{', '.join(lost[:5])}。点検が挙げる）" if lost else ""))
        for f in new["families"].values():
            for note in f["notes"]:
                print(f"  {sid}: {note}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
