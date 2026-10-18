#!/usr/bin/env python3
"""sc_rows.py — 比較のデータ（拠点の写し・用語の対応・比べる行・タスク）を読み、行ごとの箱・並び・役割を引く部品。

対応（sc_axis・sc_match）・計算（sc_compute）・点検（sc_check）・書き出し（sc_book・sc_drawio）が、同じ引き方を使う。
"""
from sc_common import (OTHER_SITE, STEP_TYPES, load, load_snap, match_path, mentions, norm, ordered_sites, site_ids, site_names, split_key,
                       try_load)


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
        ap = try_load(cmp / "applies.json")[0]
        self.applies = {}                                            # (資料の中身のハッシュ, ページ, 項目の id) → 書き分けの項目
        for it in (ap.get("items") if isinstance(ap, dict) else None) or []:
            parts = str(it.get("key", "")).split(".") if isinstance(it, dict) else []
            if len(parts) == 3 and it.get("doc"):
                self.applies[(it["doc"], parts[1], parts[2])] = it
        self.shared_docs = {k[0] for k in self.applies}

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

    def lane_label(self, sid, gid):
        """拠点の塊のレーンの名前: 組の名前と、その拠点の用語が違えば「出納職員（事務員）」"""
        g = next((g for g in self.terms.get("groups") or [] if g.get("id") == gid), None)
        if not g:
            return str(gid).lstrip("?")
        own = [m["term"] for m in g.get("members") or [] if m.get("site") == sid and norm(m.get("term", "")) != norm(g["name"])]
        return f"{g['name']}（{'・'.join(own)}）" if own else g["name"]

    def role_label(self, sid, key):
        return self.lane_label(sid, self.role(sid, key))

    # ---- ほかの拠点の作業 ----
    def owner(self, sid, key):
        """箱が、ほかの拠点の担当（用語の組の site）のレーンにあれば、その拠点の ID"""
        n = self.node(sid, key) or {}
        if n.get("lane_kind") == "system":
            return None
        o = (self.term(sid, "actor", n.get("lane", "")) or {}).get("site")
        return o if o and o != sid and o in self.names else None

    def done_by(self, sid, keys):
        """ほかの拠点の担当のレーンの箱の、行う人の言い方（「本部拠点」「本部の固定資産管理責任者（本部拠点）」）"""
        out = []
        for k in keys:
            o = self.owner(sid, k)
            if not o:
                continue
            g = self.term(sid, "actor", (self.node(sid, k) or {}).get("lane", "")) or {}
            x = next((s for s in self.sites.get("sites") or [] if s.get("id") == o), {})
            plain = {norm(n) for n in [x.get("name", "")] + list(x.get("aliases") or []) if n}
            out.append(self.names[o] if norm(g.get("name", "")) in plain else f"{g.get('name', '')}（{self.names[o]}）")
        return "・".join(dict.fromkeys(out))

    def excluded(self, sid):
        """その拠点のフローから外した項目: {鍵: 項目}（同じ項目が何本かのフローにあれば、はじめのもの）"""
        out = {}
        for x in (self.snaps.get(sid) or {}).get("excluded") or []:
            out.setdefault(x.get("key"), x)
        return out

    def applies_of(self, doc, key):
        """資料の中身のハッシュと項目の鍵（資料 ID.ページ.項目）から、書き分けの項目（無ければ None）"""
        parts = str(key).split(".")
        return self.applies.get((doc, parts[1], parts[2])) if len(parts) == 3 else None

    def not_here_key(self, sid, key):
        """not_here の鍵 1 つ: (理由の文, 誤りの文)。どちらか一方だけが入る"""
        from sc_applies import applies_to
        n = (self.snaps.get(sid) or {}).get("nodes", {}).get(key)
        if n is not None:
            who = [o for o in mentions(f"{n.get('label', '')} {n.get('detail', '')}", self.sites) if o != sid]
            if not who:
                return None, f"箱 {key} の名前・補足に、ほかの拠点の名前が無い（ほかの拠点が行うと書いた箱だけを書く）"
            return f"資料に、{'・'.join(self.names[o] for o in who)}が行うとある", None
        x = self.excluded(sid).get(key)
        if x is None:
            return None, f"{key} は、この拠点がフローから外した項目にも、この拠点の箱にも無い"
        if not str(x.get("reason", "")).startswith(OTHER_SITE):
            return None, f"{key} の外した理由が「{OTHER_SITE}」でない（{str(x.get('reason', ''))[:30]}）"
        it = self.applies_of(x.get("doc"), key)
        if it is None:
            return "資料に、ほかの拠点だけの手順とある", None
        if applies_to(it, sid):
            return None, f"{key} は、書き分け（applies.json）で{self.names[sid]}に当てはまる手順"
        only = [s for s in self.order if applies_to(it, s)]
        return (f"共通の資料の書き分けで、{'・'.join(self.names[s] for s in only)}だけの手順" if only else "共通の資料の書き分けで、この拠点の手順ではない"), None

    def derivable(self, t, sid):
        """機械で分かる not_here の鍵: ほかの拠点のこのタスクの箱が出典にする共通の資料の手順が、書き分けでこの拠点に当てはまらず、
        当てはまる共通の資料の手順を出典にした箱は無く、この拠点がその手順を「ほかの拠点だけの手順」で外しているとき"""
        from sc_applies import applies_to
        no, yes = [], False
        for s, ks in (t.get("members") or {}).items():
            if s == sid:
                continue
            for k in ks or []:
                for it in (self.node(s, k) or {}).get("items") or []:
                    if it.get("doc") not in self.shared_docs:
                        continue
                    a = self.applies.get((it["doc"], str(it.get("page")), it.get("item")))
                    ok = True if a is None else applies_to(a, sid)          # 書き分けに拾われていない共通の資料の手順は、どの拠点にも当たる
                    if ok:
                        yes = True
                    elif ok is False:
                        no.append(a)
        if yes or not no:
            return []
        want = {(a["doc"], *a["key"].split(".")[1:]) for a in no}
        return sorted(k for k, x in self.excluded(sid).items()
                      if str(x.get("reason", "")).startswith(OTHER_SITE) and (x.get("doc"), *str(k).split(".")[1:]) in want)

    def task_split(self, t):
        """(自拠点の箱 {拠点: [鍵]}, このタスクを、その拠点の作業として比べない拠点 {拠点: 理由の文})。
        箱が全部ほかの拠点の担当のレーンにある拠点と、記載なしに not_here（この拠点の手順ではない根拠）を書いた拠点が、対象外。
        どの拠点にも自拠点の箱が無いタスクは、ほかの拠点の箱を、そのまま記載ありにする（比べる物が無くなる）"""
        mem = {s: list(ks or []) for s, ks in (t.get("members") or {}).items()}
        own = {s: [k for k in ks if not self.owner(s, k)] for s, ks in mem.items()}
        if not any(own.values()):
            own = mem
        out = {}
        for s in self.order:
            if mem.get(s) and not own.get(s):
                out[s] = f"資料では、{self.done_by(s, mem[s])}が行う"
            ab = (t.get("absent") or {}).get(s)
            if not mem.get(s) and isinstance(ab, dict) and ab.get("not_here"):
                why = [self.not_here_key(s, k)[0] for k in ab["not_here"]]
                out[s] = "／".join(dict.fromkeys(w for w in why if w)) or "資料に、この拠点の手順ではないとある"
        return own, out

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
