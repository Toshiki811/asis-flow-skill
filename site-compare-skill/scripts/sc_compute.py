#!/usr/bin/env python3
"""sc_compute.py — 比べる行ごとに、タスクの並び・区分・記号・差の種類・粒度・複雑性・差の大きさ・追加の資料の依頼を計算する（工程 6。機械だけ）。

  python3 scripts/sc_compute.py -w <比較の作業フォルダ>              # 全部の行（タスクの対応 rows/<行>.match.json のある行）
  python3 scripts/sc_compute.py -w <比較の作業フォルダ> --row B03

書くもの: rows/<行>.result.json（手で直さない）・numbers.json（タスクの通し番号 T-001。行とタスクの ID で引き継ぐ）。
決め方（references/compare-rules.md の 3 章）:
  並び      基準あり: 基準の拠点のタスクを、基準の箱の並びで。ほかのタスクは、それを持つ拠点で線をたどって前にある基準のタスクのあと
            （無ければ、あとにある基準のタスクの前）。基準なし: 前後を支持する拠点の数で並べる（輪は、平均の位置で切る）
  区分      記載のある拠点が 1 つ → 拠点固有／比べる拠点に全部記載があり、差が無い → 共通／ほか → 差異あり
  記号      ○ 差なし／△ その拠点が差に関わる／— 記載なし／対象外。比べる相手は、基準（無ければ、記載のある最初の拠点）
  差の種類  担当（レーンの組）・システム・帳票は「同じ／違う／片方だけ記載」（片方だけ記載は差に数えない）。1 つの箱で書いた拠点（粗い側）の
            担当・システム・帳票が、詳しい拠点のものに含まれていれば同じ。判断は、行き先が別のタスクに分かれる判断の数（粗い側とは比べない）。
            順序は、線でたどれる前後が拠点で逆。内容は LLM が書いたもの。記載なし
  複雑性    原典: 分岐＝判断ごとの（行き先の数−1）の合計・受け渡し＝レーンが変わる線（システムの箱は縮める。開始・終了・帳票・別のフローへ進む箱は数えない）。
            そろえた: 分岐＝行き先のタスクが分かれる判断（タスクの数−1）・受け渡し＝タスクの境をまたぎ、担当の組が変わる線。
            点数と大中小・差の大きさ（基準との点数の差）は assets/sc_complexity.json。差のあるタスクの割合は、数として出すだけ
"""
import argparse
import json
import sys

from sc_common import ASSETS, STEP_TYPES, load, result_path, save, sites_of, split_key, work_dir
from sc_rows import Model, back_edges

MARK_SAME, MARK_DIFF, MARK_NONE, MARK_OUT = "○", "△", "—", "対象外"
REAL = ("担当", "システム", "帳票", "判断", "順序", "内容", "記載なし")


def config():
    c = load(ASSETS / "sc_complexity.json")
    return c


def level_of(score, cfg):
    for lv in cfg["levels"]:
        if lv["max"] is None or score <= lv["max"]:
            return lv["name"]
    return cfg["levels"][-1]["name"]


# ---------- 並び ----------

def reach_sets(nodes, edges):
    """箱 → 線でたどれる箱の集合（戻りの線は除く）"""
    back = back_edges(nodes, edges)
    out_of = {k: [] for k in nodes}
    for a, b, *_ in edges:
        if (a, b) not in back and a in out_of and b in out_of:
            out_of[a].append(b)
    reach = {}
    for k in nodes:
        seen, todo = set(), list(out_of[k])
        while todo:
            v = todo.pop()
            if v not in seen:
                seen.add(v)
                todo += out_of[v]
        reach[k] = seen
    return reach


def before(reach, mem, a, b):
    """拠点の中で、タスク a のどれかの箱から b のどれかの箱へたどれ、逆はたどれない"""
    A, B = mem.get(a) or [], mem.get(b) or []
    if not A or not B:
        return False
    ab = any(y in reach.get(x, ()) for x in A for y in B)
    ba = any(x in reach.get(y, ()) for x in A for y in B)
    return ab and not ba


def order_tasks(mdl, rid, tasks, graphs):
    """タスクの並び（ID の並び）と、順序の差 {タスク: {拠点}}"""
    ids = [t["id"] for t in tasks]
    mem = {sid: {t["id"]: (t.get("members") or {}).get(sid) or [] for t in tasks} for sid in graphs}
    pos = {}
    for sid, (nodes, _e, _r) in graphs.items():
        at = {k: i for i, k in enumerate(nodes)}
        for t in ids:
            ks = [at[k] for k in mem[sid][t] if k in at]
            if ks:
                pos.setdefault(t, []).append(min(ks) / max(1, len(nodes)))
    mean = {t: sum(pos.get(t, [1.0])) / len(pos.get(t, [1.0])) for t in ids}
    base = mdl.baseline if mdl.baseline in graphs else None
    if base:
        at = {k: i for i, k in enumerate(graphs[base][0])}
        bt = sorted([t for t in ids if mem[base][t]], key=lambda t: (min(at.get(k, 1e9) for k in mem[base][t]), t))
        order = list(bt)
        for t in [t for t in ids if not mem[base][t]]:
            sid = next((s for s in mdl.order if s in graphs and mem[s][t]), None)
            reach = graphs[sid][2] if sid else {}
            prev = [p for p in order if p in bt and sid and before(reach, mem[sid], p, t)]
            nxt = [q for q in order if q in bt and sid and before(reach, mem[sid], t, q)]
            if prev:
                i = order.index(prev[-1]) + 1
                while i < len(order) and order[i] not in bt:
                    i += 1
                order.insert(i, t)
            elif nxt:
                order.insert(order.index(nxt[0]), t)
            else:
                order.append(t)
    else:
        w = {}
        for sid, (_n, _e, reach) in graphs.items():
            for a in ids:
                for b in ids:
                    if a != b and before(reach, mem[sid], a, b):
                        w[(a, b)] = w.get((a, b), 0) + 1
        left, order = list(ids), []
        while left:
            free = [t for t in left if not any(w.get((p, t), 0) > w.get((t, p), 0) for p in left if p != t)]
            pick = min(free or left, key=lambda t: (mean[t], ids.index(t)))
            order.append(pick)
            left.remove(pick)
    conflicts = {}
    sids = list(graphs)
    for a in ids:
        for b in ids:
            if a >= b:
                continue
            ab = [s for s in sids if before(graphs[s][2], mem[s], a, b)]
            ba = [s for s in sids if before(graphs[s][2], mem[s], b, a)]
            if ab and ba:
                ref = base if base in ab + ba else None
                odd = (ba if ref in ab else ab) if ref else ab + ba
                for t in (a, b):
                    conflicts.setdefault(t, set()).update(odd)
    return order, conflicts


# ---------- 差 ----------

def compare_sets(mine, ref, mine_coarse, ref_coarse):
    if mine == ref:
        return "同じ"
    if not mine or not ref:
        return "片方だけ記載"
    if (mine_coarse and mine <= ref) or (ref_coarse and ref <= mine):
        return "同じ"
    return "違う"


def cross_decisions(mdl, sid, keys, task_of, edges):
    """行き先が別のタスクに分かれる判断の数"""
    n = 0
    for k in keys:
        if (mdl.node(sid, k) or {}).get("type") != "decision":
            continue
        tgt = {task_of.get(b, "END" if split_key(b)[1].startswith("end") else f"OUT:{b}") for a, b, *_ in edges if a == k}
        n += len(tgt) > 1
    return n


def task_diffs(mdl, rid, t, present, compared, conflicts, edges_of):
    """差: [{kind, state, sites, text}]・片方だけ記載: [...]・記号"""
    mem = t.get("members") or {}
    ref = mdl.baseline if mdl.baseline in present else (present[0] if present else None)
    diffs, onesided = [], []
    cnt = {s: len(mem.get(s) or []) for s in present}
    for kind, fn in (("担当", lambda s, ks: {mdl.role(s, k) for k in ks}),
                     ("システム", lambda s, ks: set().union(*[mdl.systems(s, k) for k in ks]) if ks else set()),
                     ("帳票", lambda s, ks: set().union(*[mdl.documents(s, k) for k in ks]) if ks else set())):
        bad, one = [], []
        rset = fn(ref, mem.get(ref) or []) if ref else set()
        for s in present:
            if s == ref:
                continue
            sset = fn(s, mem.get(s) or [])
            st = compare_sets(sset, rset, cnt[s] == 1 and cnt[ref] > 1, cnt[ref] == 1 and cnt[s] > 1)
            text = f"{mdl.names[s]}: {'・'.join(sorted(mdl.group_name(x) for x in sset)) or '記載なし'}／{mdl.names[ref]}: {'・'.join(sorted(mdl.group_name(x) for x in rset)) or '記載なし'}"
            if st == "違う":
                bad.append((s, text))
            elif st == "片方だけ記載":
                one.append((s, text))
        if bad:
            diffs.append({"kind": kind, "state": "違う", "sites": [s for s, _ in bad], "text": "／".join(x for _, x in bad)})
        if one:
            onesided.append({"kind": kind, "state": "片方だけ記載", "sites": [s for s, _ in one], "text": "／".join(x for _, x in one)})
    if ref:
        dref = cross_decisions(mdl, ref, mem.get(ref) or [], mdl.task_of(rid, ref), edges_of[ref])
        bad = []
        for s in present:
            if s == ref or cnt[s] == 1 and cnt[ref] > 1 or cnt[ref] == 1 and cnt[s] > 1:
                continue
            d = cross_decisions(mdl, s, mem.get(s) or [], mdl.task_of(rid, s), edges_of[s])
            if d != dref:
                bad.append((s, f"{mdl.names[s]}: 判断 {d}／{mdl.names[ref]}: 判断 {dref}"))
        if bad:
            diffs.append({"kind": "判断", "state": "違う", "sites": [s for s, _ in bad], "text": "／".join(x for _, x in bad)})
    if conflicts.get(t["id"]):
        ss = [s for s in mdl.order if s in conflicts[t["id"]]]
        diffs.append({"kind": "順序", "state": "違う", "sites": ss, "text": "前後のタスクとの順が、" + "・".join(mdl.names[s] for s in ss) + " で逆"})
    for c in t.get("content") or []:
        diffs.append({"kind": "内容", "state": c.get("kind", "その他"), "sites": [s for s in c.get("sites") or [] if s in mdl.names], "text": c.get("text", "")})
    absent = [s for s in compared if s not in present]
    if absent and present:
        diffs.append({"kind": "記載なし", "state": "記載なし", "sites": absent, "text": "・".join(mdl.names[s] for s in absent) + " の資料に記載なし"})
    return diffs, onesided, ref


# ---------- 複雑性 ----------

def contract(mdl, sid, nodes, edges):
    """システムの箱を縮めた線: [(元, 先)]（元・先とも、システムでない箱）"""
    ty = {k: (mdl.node(sid, k) or {}).get("type") for k in nodes}
    out_of = {}
    for a, b, *_ in edges:
        out_of.setdefault(a, []).append(b)
    res = []
    for u in nodes:
        if ty[u] == "system":
            continue
        seen, todo = set(), list(out_of.get(u, []))
        while todo:
            v = todo.pop()
            if v in seen:
                continue
            seen.add(v)
            if ty.get(v) == "system":
                todo += out_of.get(v, [])
            else:
                res.append((u, v))
    return res


def complexity(mdl, sid, nodes, edges, task_of, cfg):
    """拠点 × 行の複雑性: {raw, aligned, counted}"""
    ty = {k: (mdl.node(sid, k) or {}).get("type") for k in nodes}
    lab = lambda k: (mdl.node(sid, k) or {}).get("label", "")
    counted = []
    raw_br = al_br = 0
    for d in [k for k in nodes if ty[k] == "decision"]:
        outs = [(b, lbl) for a, b, lbl, *_ in edges if a == d]
        n = len({b for b, _ in outs}) - 1
        if n > 0:
            raw_br += n
            counted.append({"basis": "raw", "kind": "分岐", "at": d, "text": f"{lab(d)}（{'／'.join(l or '—' for _, l in outs)}）"})
        tg = {task_of.get(b, "END" if ty.get(b) == "end" else f"OUT:{b}") for b, _ in outs}
        if len(tg) > 1:
            al_br += len(tg) - 1
            counted.append({"basis": "aligned", "kind": "分岐", "at": d, "text": f"{lab(d)}（行き先のタスク {len(tg)}）"})
    raw_h = al_h = 0
    skip = ("start", "end", "subprocess")
    for u, v in contract(mdl, sid, nodes, edges):
        if ty.get(u) in skip or ty.get(v) in skip:
            continue
        lu, lv = (mdl.node(sid, u) or {}).get("lane"), (mdl.node(sid, v) or {}).get("lane")
        if lu != lv:
            raw_h += 1
            counted.append({"basis": "raw", "kind": "受け渡し", "at": f"{u}→{v}", "text": f"{lu} → {lv}（{lab(u)} → {lab(v)}）"})
        if task_of.get(u) != task_of.get(v) and mdl.role(sid, u) != mdl.role(sid, v):
            al_h += 1
            counted.append({"basis": "aligned", "kind": "受け渡し", "at": f"{u}→{v}",
                            "text": f"{mdl.role_name(sid, u)} → {mdl.role_name(sid, v)}（{lab(u)} → {lab(v)}）"})
    w = cfg["weights"]
    mk = lambda br, h: {"br": br, "hand": h, "score": w["branch"] * br + w["handoff"] * h, "level": level_of(w["branch"] * br + w["handoff"] * h, cfg)}
    return {"raw": mk(raw_br, raw_h), "aligned": mk(al_br, al_h), "counted": counted}


def spread_level(d, cfg):
    sc = cfg["spread"]["score"]
    return "大" if d >= sc[1] else "中" if d >= sc[0] else "小"


# ---------- 追加の資料 ----------

DOC_KIND = (("decisions", "判断の基準が書かれた資料（規程・基準表・チェックリスト）"), ("handoffs", "担当の分担が分かる資料（業務分掌・権限表）"),
            ("systems", "システムの操作の手順書"), ("files", "帳票の様式・記入例"))


def requests(mdl, rid, res, tasks, edges_of):
    out = []
    by = {t["id"]: t for t in tasks}
    row = mdl.row_of[rid]
    for sid in mdl.participants(rid):
        sp = mdl.site_part(rid, sid) or {}
        if not sp.get("parts"):
            if sp.get("relation") == "none" and tasks:
                out.append({"site": sid, "coarse": [], "absent": [x["no"] for x in res], "folded": {}, "kinds": [f"「{row['name']}」の手順書（無ければ、この拠点では行っていないことの確かめ）"],
                            "gain": len(res), "text": f"{mdl.names[sid]}の資料に、{row['name']}の記載が無い"})
            continue
        coarse, absent = [], []
        folded = {"nodes": 0, "decisions": 0, "handoffs": 0, "systems": 0, "files": 0}
        for x in res:
            t = by[x["id"]]
            cnt = x["grain"]["counts"]
            if cnt.get(sid) == 1 and max(cnt.values()) > 1:
                coarse.append(x["no"])
                fine = max((s for s in cnt if cnt[s] > 1), key=lambda s: (cnt[s], -mdl.order.index(s)))
                ks = (t.get("members") or {}).get(fine) or []
                folded["nodes"] += len(ks) - 1
                folded["decisions"] += sum(1 for k in ks if (mdl.node(fine, k) or {}).get("type") == "decision")
                kk = set(ks)
                folded["handoffs"] += sum(1 for a, b, *_ in edges_of[fine] if a in kk and b in kk and
                                          (mdl.node(fine, a) or {}).get("lane") != (mdl.node(fine, b) or {}).get("lane"))
                folded["systems"] += len(set().union(*[mdl.systems(fine, k) for k in ks]) - set().union(*[mdl.systems(sid, k) for k in (t.get("members") or {}).get(sid) or []]))
                folded["files"] += len(set().union(*[mdl.documents(fine, k) for k in ks]) - set().union(*[mdl.documents(sid, k) for k in (t.get("members") or {}).get(sid) or []]))
            elif x["marks"].get(sid) == MARK_NONE:
                absent.append(x["no"])
        if not coarse and not absent:
            continue
        kinds = [text for key, text in DOC_KIND if folded[key]]
        if coarse and not kinds:
            kinds.append("詳しい手順書（作業の順と担当が分かるもの）")
        if absent:
            kinds.append("記載なしの作業の手順書（無ければ、この拠点では行っていないことの確かめ）")
        parts = []
        if coarse:
            parts.append(f"{len(coarse)} つのタスクを 1 つの作業で書いている（ほかの拠点の作業 {folded['nodes']}・判断 {folded['decisions']}・受け渡し {folded['handoffs']} と比べられない）")
        if absent:
            parts.append(f"{len(absent)} つのタスクの記載が無い")
        out.append({"site": sid, "coarse": coarse, "absent": absent, "folded": folded, "kinds": kinds, "gain": len(coarse) + len(absent),
                    "text": f"{mdl.names[sid]}: " + "。".join(parts)})
    return out


# ---------- 行 ----------

def numbers(cmp, mdl):
    nb = load(cmp / "numbers.json", None) or {"next": 1, "assigned": {}}
    return nb


def compute_row(mdl, rid, cfg, nb):
    row = mdl.row_of[rid]
    tasks = mdl.tasks(rid)
    graphs, edges_of = {}, {}
    for sid in mdl.order:
        nodes, edges = mdl.row_nodes(rid, sid)
        edges_of[sid] = edges
        if nodes:
            graphs[sid] = (nodes, edges, reach_sets(nodes, edges))
    order, conflicts = order_tasks(mdl, rid, tasks, graphs)
    compared = mdl.participants(rid)
    by = {t["id"]: t for t in tasks}
    res = []
    for pos, tid in enumerate(order, 1):
        t = by[tid]
        mem = t.get("members") or {}
        present = [s for s in mdl.order if mem.get(s)]
        diffs, onesided, ref = task_diffs(mdl, rid, t, present, compared, conflicts, edges_of)
        real = [d for d in diffs if d["kind"] in REAL and d["kind"] != "記載なし"]
        involved = {s for d in real for s in d["sites"]}
        marks = {}
        for s in mdl.order:
            if (mdl.site_part(rid, s) or {}).get("relation") == "out_of_scope":
                marks[s] = MARK_OUT
            elif s not in present:
                marks[s] = MARK_NONE
            else:
                marks[s] = MARK_DIFF if s in involved else MARK_SAME
        cls = "拠点固有" if len(present) == 1 and len(compared) > 1 else "共通" if set(present) == set(compared) and not real else "差異あり"
        cnt = {s: len(mem[s]) for s in present}
        mx = max(cnt.values()) if cnt else 0
        key = f"{rid}|{tid}"
        if key not in nb["assigned"]:
            nb["assigned"][key] = f"T-{nb['next']:03d}"
            nb["next"] += 1
        members = {s: [{"key": k, "label": (mdl.node(s, k) or {}).get("label", ""), "type": (mdl.node(s, k) or {}).get("type", ""),
                        "role": mdl.role_name(s, k), "systems": sorted(mdl.group_name(y) for y in mdl.systems(s, k) if not y.startswith("R")),
                        "files": sorted(mdl.group_name(d) for d in mdl.documents(s, k)), "src": (mdl.node(s, k) or {}).get("src", []),
                        "inferred": bool((mdl.node(s, k) or {}).get("inferred"))} for k in mem[s]] for s in present}
        res.append({"id": tid, "no": nb["assigned"][key], "pos": pos, "name": t.get("name", ""), "class": cls, "marks": marks, "ref": ref,
                    "kinds": [k for k in REAL if any(d["kind"] == k for d in diffs)], "diffs": diffs, "onesided": onesided,
                    "grain": {"counts": cnt, "coarse": [s for s in present if cnt[s] == 1 and mx > 1], "fine": [s for s in present if cnt[s] > 1]},
                    "joined": t.get("joined", ""), "members": members, "inferred": any(m["inferred"] for ms in members.values() for m in ms)})
    sites = {}
    for sid in mdl.order:
        sp = mdl.site_part(rid, sid) or {}
        entry = {"relation": sp.get("relation", "none"), "flows": [], "complexity": None, "grain": any(sid in x["grain"]["coarse"] for x in res)}
        snap = mdl.snaps.get(sid) or {}
        for p in sp.get("parts") or []:
            f = (snap.get("flows") or {}).get(p["flow"], {})
            kids = (snap.get("families") or {}).get(p["flow"], {}).get("children") if p.get("expand") else []
            entry["flows"].append(p.get("title") or f.get("name", p["flow"]) + (f"（{'・'.join(snap['flows'][k]['name'] for k in kids)}を含む）" if kids else "")
                                  + ("（一部の箱）" if isinstance(p.get("nodes"), list) else ""))
        if sid in graphs:
            nodes, edges, _r = graphs[sid]
            tof = mdl.task_of(rid, sid)
            entry["complexity"] = complexity(mdl, sid, nodes, edges, tof, cfg)
            no = {x["id"]: x["no"] for x in res}
            for it in entry["complexity"]["counted"]:                # 箱の鍵は納品物に出さない: タスクの番号で書く
                ends = it.pop("at").split("→")
                it["where"] = " → ".join(no.get(tof.get(k), "—") for k in ends)
        sites[sid] = entry
    basis = cfg["basis"]
    ref = mdl.baseline if (sites.get(mdl.baseline) or {}).get("complexity") else None
    scores = {s: e["complexity"][basis]["score"] for s, e in sites.items() if e["complexity"]}
    if ref is None and scores:
        ref = min(scores, key=lambda s: (scores[s], mdl.order.index(s)))
    for s, e in sites.items():
        if not e["complexity"] or s == ref:
            e["spread"] = None
            continue
        ratio = sum(1 for x in res if x["marks"].get(s) in (MARK_DIFF, MARK_NONE)) / max(1, len(res))
        e["spread"] = {"diff": scores[s] - scores[ref], "ratio": round(ratio, 2), "level": spread_level(abs(scores[s] - scores[ref]), cfg)}
    lv = [e["spread"]["level"] for e in sites.values() if e.get("spread")]
    grp = next((g for g in mdl.axis.get("groups") or [] if g.get("id") == row.get("group")), None)
    return {"row": rid, "name": row["name"], "group": row.get("group"), "group_name": grp["name"] if grp else "", "parent": row.get("parent"),
            "mode": mdl.axis.get("mode"), "basis": basis, "ref": ref, "sites": sites, "tasks": res,
            "spread": max(lv, key=["小", "中", "大"].index) if lv else None,
            "requests": requests(mdl, rid, res, tasks, edges_of)}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--row")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    mdl = Model(cmp, sites_of(cmp))
    cfg = config()
    nb = numbers(cmp, mdl)
    todo = [a.row] if a.row else [r["id"] for r in mdl.rows()]
    done = 0
    for rid in todo:
        if rid not in mdl.row_of:
            sys.exit(f"行 {rid} が axis.json に無い")
        if not mdl.match(rid):
            print(f"{rid}: タスクの対応（rows/{rid}.match.json）が無いので、計算しない")
            continue
        r = compute_row(mdl, rid, cfg, nb)
        save(result_path(cmp, rid), r)
        done += 1
        cls = {}
        for x in r["tasks"]:
            cls[x["class"]] = cls.get(x["class"], 0) + 1
        print(f"{rid} {r['name']}: タスク {len(r['tasks'])}（" + "・".join(f"{k} {v}" for k, v in cls.items()) + "）・複雑性 "
              + "・".join(f"{mdl.names[s]} {e['complexity'][r['basis']]['level']}" for s, e in r["sites"].items() if e["complexity"]))
    save(cmp / "numbers.json", nb)
    print(f"計算した行 {done}")


if __name__ == "__main__":
    main()
