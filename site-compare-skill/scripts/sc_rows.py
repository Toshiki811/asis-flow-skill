#!/usr/bin/env python3
"""sc_rows.py — 比較のデータ（拠点の写し・用語の対応・比べる行・タスク）を読み、行ごとの箱・並び・役割を引く部品。

対応（sc_axis・sc_match）・計算（sc_compute）・点検（sc_check）・書き出し（sc_book・sc_drawio）が、同じ引き方を使う。
"""
from sc_common import (STEP_TYPES, load, load_snap, match_path, norm, ordered_sites, site_ids, site_names, split_key, try_load)


def linear_order(nodes, edges):
    """箱の並び（開始から線をたどり、前の箱がすべて並んだ箱から置く。戻りの線は数えない。枝は線の順）。asis の step_order_by_branch と同じ考え方"""
    nodes = list(nodes)
    ns = set(nodes)
    out_of = {k: [] for k in nodes}
    for a, b, *_ in edges:
        if a in ns and b in ns:
            out_of[a].append(b)
    depth, state = {}, {}
    starts = [k for k in nodes if split_key(k)[1] == "start"] or nodes[:1]
    stack = [(s, 0, iter(out_of[s])) for s in starts]
    for s in starts:
        depth[s], state[s] = 0, 1
    while stack:                                                     # 深さ優先（再帰を使わない）。たどっている途中の箱へ戻る線は数えない
        u, d, it = stack[-1]
        v = next(it, None)
        if v is None:
            state[u] = 2
            stack.pop()
            continue
        if state.get(v) == 1:
            continue
        if state.get(v) == 2 and depth.get(v, -1) >= d + 1:
            continue
        depth[v] = max(depth.get(v, -1), d + 1)
        state[v] = 1
        stack.append((v, depth[v], iter(out_of[v])))
    last = max(depth.values(), default=0)
    for k in nodes:
        depth.setdefault(k, last + 1)
    succ, pred = {k: [] for k in nodes}, {k: [] for k in nodes}
    for a in nodes:
        for b in out_of[a]:
            if depth[b] > depth[a] and b not in succ[a]:
                succ[a].append(b)
                pred[b].append(a)
    done, out, todo = set(), [], starts[::-1]
    while todo:
        u = todo.pop()
        if u in done or any(p not in done for p in pred[u]):
            continue
        done.add(u)
        out.append(u)
        todo += succ[u][::-1]
    return out + [k for k in nodes if k not in done]


def back_edges(nodes, edges):
    """戻りの線（前に置いた箱へ戻る線）の集合 {(元, 先)}"""
    pos = {k: i for i, k in enumerate(linear_order(nodes, edges))}
    return {(a, b) for a, b, *_ in edges if a in pos and b in pos and pos[b] <= pos[a]}


class Model:
    """比較の作業フォルダのデータ一式"""

    def __init__(self, cmp, sites):
        self.cmp, self.sites = cmp, sites
        self.ids = site_ids(sites)
        self.order = ordered_sites(sites)
        self.names = site_names(sites)
        self.baseline = sites.get("baseline")
        self.snaps = {sid: load_snap(cmp, sid) for sid in self.ids}
        self.terms = try_load(cmp / "terms.json")[0] or {"groups": []}
        self.axis = try_load(cmp / "axis.json")[0] or {"rows": [], "groups": [], "not_compared": []}
        self.row_of = {r["id"]: r for r in self.axis.get("rows") or [] if isinstance(r, dict) and "id" in r}
        self.group_of_term = {}
        for g in self.terms.get("groups") or []:
            for m in g.get("members") or []:
                self.group_of_term[(m.get("site"), g.get("kind"), m.get("term"))] = g
        self._matches = {}

    # ---- 用語 ----
    def term(self, sid, kind, name):
        """(拠点, 種類, 名前) → 組（無ければ None）"""
        return self.group_of_term.get((sid, kind, name))

    def node(self, sid, key):
        s = self.snaps.get(sid)
        return (s or {}).get("nodes", {}).get(key)

    def role(self, sid, key):
        """箱の担当の組の ID（レーンがシステムなら Y、担当なら R。組が無ければ、レーンの名前を「?」つきで）"""
        n = self.node(sid, key) or {}
        kind = "system" if n.get("lane_kind") == "system" else "actor"
        g = self.term(sid, kind, n.get("lane", ""))
        return g["id"] if g else f"?{n.get('lane', '')}"

    def role_name(self, sid, key):
        n = self.node(sid, key) or {}
        kind = "system" if n.get("lane_kind") == "system" else "actor"
        g = self.term(sid, kind, n.get("lane", ""))
        return g["name"] if g else n.get("lane", "")

    def systems(self, sid, key):
        n = self.node(sid, key) or {}
        out = {(self.term(sid, "system", y) or {"id": f"?{y}"})["id"] for y in n.get("systems") or []}
        if n.get("lane_kind") == "system":
            out.add(self.role(sid, key))
        return out

    def documents(self, sid, key):
        n = self.node(sid, key) or {}
        return {(self.term(sid, "document", f["name"]) or {"id": f"?{f['name']}"})["id"] for f in n.get("files") or []}

    def group_name(self, gid):
        g = next((g for g in self.terms.get("groups") or [] if g.get("id") == gid), None)
        return g["name"] if g else str(gid).lstrip("?")

    # ---- 行 ----
    def rows(self):
        return [r for r in self.axis.get("rows") or [] if isinstance(r, dict) and "id" in r]

    def site_part(self, rid, sid):
        r = self.row_of.get(rid) or {}
        return next((x for x in r.get("sites") or [] if x.get("site") == sid), None)

    def participants(self, rid):
        """その行で比べる拠点（対象外を除く）。並びは、基準が先頭"""
        return [sid for sid in self.order if (self.site_part(rid, sid) or {}).get("relation") != "out_of_scope"]

    def with_parts(self, rid):
        """その行に箱を当てた拠点"""
        return [sid for sid in self.order if (self.site_part(rid, sid) or {}).get("parts")]

    def part_graph(self, sid, part):
        """当てたフローの箱と線: (箱の鍵の並び, 線)。expand: 開いた 1 本／しない: そのフローだけ。nodes があれば、その箱と開始・終了だけ"""
        s = self.snaps.get(sid) or {}
        f = part.get("flow")
        if part.get("expand"):
            fam = (s.get("families") or {}).get(f) or {"nodes": [], "edges": []}
            nodes, edges = list(fam["nodes"]), [tuple(e) for e in fam["edges"]]
        else:
            fl = (s.get("flows") or {}).get(f) or {"nodes": [], "edges": []}
            nodes, edges = list(fl["nodes"]), [tuple(e) for e in fl["edges"]]
        if isinstance(part.get("nodes"), list):
            keep = set(part["nodes"]) | {k for k in nodes if (s.get("nodes", {}).get(k) or {}).get("type") in ("start", "end")}
            nodes = [k for k in nodes if k in keep]
        return nodes, edges

    def row_nodes(self, rid, sid):
        """その行のその拠点の箱（並び順。開始・終了・別のフローへ進む箱も含む）と線"""
        sp = self.site_part(rid, sid)
        nodes, edges = [], []
        for p in (sp or {}).get("parts") or []:
            n, e = self.part_graph(sid, p)
            nodes += [k for k in n if k not in nodes]
            edges += [x for x in e if x not in edges]
        ns = set(nodes)
        edges = [e for e in edges if e[0] in ns and e[1] in ns]
        return linear_order(nodes, edges), edges

    def step_nodes(self, rid, sid):
        s = self.snaps.get(sid) or {}
        return [k for k in self.row_nodes(rid, sid)[0] if (s.get("nodes", {}).get(k) or {}).get("type") in STEP_TYPES]

    # ---- タスク ----
    def match(self, rid):
        if rid not in self._matches:
            m, err = try_load(match_path(self.cmp, rid))
            self._matches[rid] = m if isinstance(m, dict) else None
        return self._matches[rid]

    def tasks(self, rid):
        return [t for t in (self.match(rid) or {}).get("tasks") or [] if isinstance(t, dict) and "id" in t]

    def task_of(self, rid, sid):
        """箱 → タスクの ID（その行のその拠点）"""
        out = {}
        for t in self.tasks(rid):
            for k in (t.get("members") or {}).get(sid) or []:
                out.setdefault(k, t["id"])
        return out
