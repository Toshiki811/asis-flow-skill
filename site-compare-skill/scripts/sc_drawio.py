#!/usr/bin/env python3
"""sc_drawio.py — 拠点比較フロー（out/拠点比較フロー.drawio）を書き出す（工程 8。機械だけ。LLM は XML を書かない）。

  python3 scripts/sc_drawio.py -w <比較の作業フォルダ>

1 ページ目  業務種別 × 拠点の表（複雑性の大中小の色・分岐と受け渡しの数・粒度差の印・差の大きさ）。セルから各ページへ
各ページ    業務種別（比べる行）ごとに 1 ページ。拠点のレーンの塊を横に並べ（基準が左。上に拠点の見出し）、左の「タスク」の列に T 番号とタスク名。
            タスクごとの帯で、同じタスクの箱を同じ高さにそろえる（帯の高さは、そのタスクの箱がいちばん多い拠点の段の数）。
            箱の色は区分（共通・差異あり・拠点固有）。記載なしは点線の空き枠。推定は赤い破線（asis と同じ）。ほかの拠点が 1 つの箱で書いた
            タスクの、詳しい拠点の箱は、点線の枠（内訳）で囲む。箱の中は「T 番号: 作業の名前」と、差のある拠点には「△ 差の種類」。
            出典・担当・差の内容は、マウスを乗せたときの説明に（図には出さない）。レーンは比べる用語の名前（日本語）で、役割の順にそろえる。
            5 拠点以上は、基準（無ければ 1 つ目の拠点）と 3 拠点ずつのページに分ける。タスクが 25 を超える行は、タスクの境でページを分ける。
線の通し方・戻りの線・配置の点検は、asis-flow-skill の業務フローの図と同じ部品（FlowDrawing・check_layout.page_problems）。
並べた配置で配置の点検に当たるページは、レーンを広げて描き直し、それでも当たればページの番号を出す（sc_layout.py が FAIL にする）。
"""
import argparse
import copy
import json
import sys

from sc_common import ASIS, ASSETS, A, FIG, load, result_path, save, sites_of, split_key, try_load, work_dir
from sc_rows import Model

import build_drawio as BD                                            # asis の図の部品
from check_layout import page_problems

MAX_SITES, MAX_TASKS = 4, 25


class _Plain:
    """出典を図に出さないので、asis の出典の書き方（labels）は使わない。node_box が呼ぶ所だけを満たす"""


def styles():
    S = json.loads((ASIS / "assets" / "template_styles.json").read_text(encoding="utf-8"))
    X = load(ASSETS / "sc_styles.json")
    return S, X


def esc(t):
    return BD.esc(t)


class CompareDrawing(BD.FlowDrawing):
    """比べる行 1 つ（の拠点の組）を 1 ページに。asis の FlowDrawing に、帯でそろえた行を渡し、レーン・箱・凡例の描き方を比較用にする"""

    def __init__(self, fl, ctx, rows, meta, page_id, page_name):
        self._rows, self.meta = rows, meta
        super().__init__(fl, ctx)
        self.page = BD.Page(page_id, page_name[:60], self.S["model"])

    def split_cols(self):
        return {}, self._rows

    def measure(self):
        X = self.meta["X"]
        self.geo, self.label, self.cap = {}, {}, {}
        for n in self.main:
            m = self.meta["node"][n["id"]]
            if m["kind"] == "task_label":
                lines = BD.n_lines(m["text"], X["lane"]["task_width"] - 12, 10) + 1
                self.geo[n["id"]] = [0, 0, X["lane"]["task_width"] - 16, 14 + 13 * lines]
                self.label[n["id"]] = f"<b>{esc(m['no'])}</b><br>{esc(m['text'])}"
                continue
            show = dict(n, id=m.get("no") or "⇒", sources=[])
            if m["kind"] in ("placeholder", "out"):
                show["label"] = m["text"]
            w, h, html, _cite = BD.node_box(show, self.S)
            if not m.get("no"):                                      # タスクに入らない箱（別のフローへ進む箱）: 番号の代わりに「⇒」
                html = html.replace(esc("⇒: "), "⇒ ", 1)
            if m.get("mark"):
                html += "<br>" + f'<font style="font-size:9px" color="{X["mark"]["△"]}">{esc(m["mark"])}</font>'
                h += 13
            if m.get("coarse"):
                html += "<br>" + f'<font style="font-size:9px" color="{X["mark"]["粗"]}">{esc(m["coarse"])}</font>'
                h += 13
            self.geo[n["id"]] = [0, 0, w, h]
            self.label[n["id"]] = html
        self.files_of, self.fgeo = {}, {}

    def place(self):
        super().place()
        tl = self.meta["task_lane"]
        if tl is not None:                                           # タスクの列: 名前を、帯の上にそろえて左に寄せる
            for n in self.main:
                if self.meta["node"][n["id"]]["kind"] == "task_label":
                    g = self.geo[n["id"]]
                    g[0] = self.lane_x(tl) + 8
                    g[1] = self.row_y[self._rows[n["id"]]]

    def node_style(self, n):
        X, m = self.meta["X"], self.meta["node"][n["id"]]
        if m["kind"] == "task_label":
            return X["task_label"]
        if m["kind"] == "placeholder":
            return X["placeholder"]
        if m["kind"] == "out":
            return X["out_of_scope"]
        style = super().node_style(n)
        if m.get("cls") and n["type"] not in ("start", "end"):
            style = BD.add_style(style, X["class"][m["cls"]])
        if n.get("inferred"):
            style = BD.add_style(style, self.ext["inferred"])
        return style

    def draw_lanes(self):
        page, S, L, X = self.page, self.S, self.L, self.meta["X"]
        for i, l in enumerate(self.lanes):
            if i == self.meta["task_lane"]:
                page.vertex(f"lane-{l['id']}-h", esc(l["name"]), X["task_lane"], self.lane_x(i), L["header_y"], self.lane_w(i), L["header_h"])
                page.vertex(f"lane-{l['id']}-c", "", X["task_lane_body"], self.lane_x(i), L["content_y"], self.lane_w(i), self.bottom - L["content_y"])
                continue
            page.vertex(f"lane-{l['id']}-h", esc(l["name"]), self.pal[i]["header_style"], self.lane_x(i), L["header_y"], self.lane_w(i), L["header_h"])
            page.vertex(f"lane-{l['id']}-c", "", self.pal[i]["container_style"], self.lane_x(i), L["content_y"], self.lane_w(i), self.bottom - L["content_y"])
        for sid, name, first, last, base in self.meta["blocks"]:   # 拠点の見出し（レーンの塊の上）
            x0, x1 = self.lane_x(first), self.lane_x(last) + self.lane_w(last)
            page.vertex(f"site-{sid}", esc(name), X["site_header_base"] if base else X["site_header"], x0, L["header_y"] - 30, x1 - x0, 26)
        page.vertex("title", esc(self.meta["title"]), BD.add_style(S["text_style"], self.ext["title"]), L["x0"], L["header_y"] - 92, max(sum(self.vw), 600), 30)
        page.vertex("to-overview", "◀ 一覧へ", BD.add_style(S["text_style"], self.ext["link_text"]), L["x0"], L["header_y"] - 58, 120, 18,
                    attrs={"link": "data:page/id,overview"})
        right = self.lane_x(len(self.lanes) - 1) + self.lane_w(len(self.lanes) - 1)
        prev = None
        for r, (no, _text) in sorted(self.meta["band_rows"].items()):  # 帯の境（タスクの境の、行のすき間）
            if prev is not None:
                y = self.row_y[r] - self.gap / 2
                page.vertex(f"band-{r}", "", X["band_line"], L["x0"], y - 5, right - L["x0"], 10)
            prev = r

    def draw_frames(self):
        """内訳の枠: ほかの拠点が 1 つの箱で書いたタスクの、詳しい拠点の箱を囲む"""
        X = self.meta["X"]
        for k, (title, ids) in enumerate(self.meta["frames"]):
            gs = [self.geo[i] for i in ids if i in self.geo]
            if not gs:
                continue
            x0, y0 = min(g[0] for g in gs) - 10, min(g[1] for g in gs) - 16
            x1, y1 = max(g[0] + g[2] for g in gs) + 10, max(g[1] + g[3] for g in gs) + 6
            self.page.vertex(f"frame-{k}", esc(title), X["frame"], x0, y0, x1 - x0, y1 - y0)

    def draw_nodes(self):
        for n in self.main:
            m = self.meta["node"][n["id"]]
            g = self.geo[n["id"]]
            attrs = {"tooltip": m.get("tip", "")} if m.get("tip") else None
            self.page.vertex(n["id"], self.label[n["id"]], self.node_style(n), g[0], g[1], g[2], g[3], attrs=attrs)
            self.page.nodes[n["id"]] = [round(v, 1) for v in g]

    def build(self):
        self.measure()
        self.place()
        self.draw_lanes()
        self.draw_frames()
        self.draw_nodes()
        self.draw_edges()
        X = self.meta["X"]
        px = self.lx(len(self.vw)) + self.S["legend_gap"]
        rows = [("共通", X["class"]["共通"]), ("差異あり", X["class"]["差異あり"]), ("拠点固有", X["class"]["拠点固有"]), ("記載なし（資料に無い）", X["placeholder"])]
        y = self.L["header_y"]
        self.page.vertex("legend-title", "<b>凡例</b>", self.S["text_style"], px, y, 220, 20)
        for k, (name, st) in enumerate(rows):
            yy = y + 26 + 34 * k
            self.page.vertex(f"legend-{k}", "", BD.add_style("rounded=1;whiteSpace=wrap;html=1;", st), px, yy, 40, 24)
            self.page.vertex(f"legend-{k}-t", esc(name), self.S["text_style"], px + 50, yy, 200, 24)
        notes = ["T-001: タスクの番号（拠点タスク一覧と同じ）", "△ 担当 など: 基準（無ければ記載のある最初の拠点）と違う所",
                 "粗: ほかの拠点が詳しく書いたタスクを、1 つの箱で書いている", "オレンジの点線の枠: 粗く書いた拠点に合わせた、詳しい拠点の内訳",
                 "赤い破線: 原典に無く、推定で置いた箱・線", "出典・担当・差の内容は、箱にマウスを乗せると出る"]
        BD.info_box(self.page, self.S, px, y + 26 + 34 * len(rows) + 10, 330, notes)
        BD.anchors(self.page, self.ext)
        return self.page


# ---------- 1 ページ分のデータ ----------

def lane_order(mdl, rid, sids):
    """役割（用語の組）の並び: 基準の拠点のレーンの順 → ほかの拠点だけの役割。システムは右"""
    order = []
    for sid in sids:
        for k in mdl.row_nodes(rid, sid)[0]:
            n = mdl.node(sid, k) or {}
            if n.get("type") in ("start", "end"):
                continue
            g = mdl.role(sid, k)
            if g not in order:
                order.append(g)
    return sorted(order, key=lambda g: (g.startswith("Y") or g.startswith("?") and False, order.index(g)))


def page_data(mdl, rid, res, sids, tasks, X):
    """CompareDrawing に渡す、見せかけのフロー（全拠点のレーン・箱・線）と、帯でそろえた行・箱ごとの印"""
    names = mdl.names
    fl = {"flow_id": rid, "name": res["name"], "lanes": [], "nodes": [], "edges": []}
    meta = {"X": X, "node": {}, "blocks": [], "frames": [], "band_rows": {}, "task_lane": 0,
            "title": f"{res.get('label') or res['name']}（拠点比較）　" + "・".join(names[s] for s in sids)}
    fl["lanes"].append({"id": "TASK", "name": "タスク", "kind": "role"})
    roles = lane_order(mdl, rid, sids)
    lane_of = {}
    for sid in sids:
        first = len(fl["lanes"])
        used = []
        for k in mdl.row_nodes(rid, sid)[0]:
            n = mdl.node(sid, k) or {}
            if n.get("type") in ("start", "end"):
                continue
            g = mdl.role(sid, k)
            if g not in used:
                used.append(g)
        used.sort(key=lambda g: roles.index(g) if g in roles else 99)
        if not used:
            used = ["?none"]
        for g in used:
            lid = f"{sid}.{g}"
            lane_of[(sid, g)] = lid
            kind = "system" if g.startswith("Y") else "role"
            fl["lanes"].append({"id": lid, "name": mdl.group_name(g) if not g.startswith("?none") else "（記載なし）", "kind": kind})
        meta["blocks"].append((sid, names[sid], first, len(fl["lanes"]) - 1, sid == mdl.baseline))
    by = {t["id"]: t for t in tasks}
    task_res = {x["id"]: x for x in res["tasks"]}
    bands = ["START"] + [x["id"] for x in res["tasks"] if x["id"] in by] + ["END"]
    local = {b: {} for b in bands}                                   # 帯 → 拠点 → [箱の id の並び]
    node_id = lambda sid, k: f"{sid}.{k.replace(':', '.')}"
    for sid in sids:
        nodes, edges = mdl.row_nodes(rid, sid)
        tof = {}
        for x in res["tasks"]:
            for k in (by.get(x["id"], {}).get("members") or {}).get(sid) or []:
                tof[k] = x["id"]
        starts = [k for k in nodes if (mdl.node(sid, k) or {}).get("type") == "start"]
        ends = [k for k in nodes if (mdl.node(sid, k) or {}).get("type") == "end" and split_key(k)[1] == "end"]
        cur = "START"
        for k in nodes:
            n = mdl.node(sid, k) or {}
            t = n.get("type")
            if t == "start":                                         # 行に 2 本以上のフローを当てた拠点は、2 本目からの開始を、続く箱の帯に
                nxt = next((b2 for a, b2, *_ in edges if a == k), None)  # （開始を全部いちばん上に並べると、開始の線が重なって走る）
                b = "START" if k == starts[0] else tof.get(nxt, cur)
            elif k in tof:
                b = tof[k]
            else:
                b = cur                                              # 帯の無い箱（別のフローへ進む箱・途中の終了）は、前の箱の帯に
            if t == "end" and split_key(k)[1] == "end":
                b = "END" if k == ends[-1] else cur                  # いちばん下の帯には、最後のフローの終了だけ
            cur = b if b != "END" else cur
            local[b].setdefault(sid, []).append(k)
            g = mdl.role(sid, k) if t not in ("start", "end") else None
            lane = lane_of.get((sid, g)) if g else None
            if lane is None:                                         # 開始・終了は、つながる箱のレーンに
                near = next((mdl.role(sid, b2) for a, b2, *_ in edges if a == k and (mdl.node(sid, b2) or {}).get("type") not in ("start", "end")), None) or \
                    next((mdl.role(sid, a) for a, b2, *_ in edges if b2 == k and (mdl.node(sid, a) or {}).get("type") not in ("start", "end")), None)
                lane = lane_of.get((sid, near)) or next(l["id"] for l in fl["lanes"] if l["id"].startswith(f"{sid}."))
            nid = node_id(sid, k)
            x = task_res.get(tof.get(k))
            mark = ""
            if x and x["marks"].get(sid) == "△":
                kinds = [d["kind"] for d in x["diffs"] if sid in d["sites"] and d["kind"] != "記載なし"]
                mark = "△ " + "・".join(dict.fromkeys(kinds))
            coarse = "粗: ほかの拠点の " + str(max(x["grain"]["counts"].values())) + " 作業に当たる" if x and sid in x["grain"]["coarse"] else ""
            tip = [f"{names[sid]}｜{mdl.role_name(sid, k)}｜{n.get('label', '')}"]           # 担当・システム・帳票は、比べる用語の名前（日本語）で
            if n.get("detail"):
                tip.append(f"補足: {n['detail']}")
            ys = [mdl.group_name(y) for y in sorted(mdl.systems(sid, k)) if n.get("lane_kind") != "system"]
            if ys:
                tip.append("システム: " + "・".join(ys))
            ds = sorted(mdl.group_name(d) for d in mdl.documents(sid, k))
            if ds:
                tip.append("帳票: " + "・".join(ds))
            if n.get("src"):
                tip.append("出典: " + "・".join(n["src"]))
            if x:
                tip.append(f"タスク {x['no']} {x['name']}（{x['class']}）")
                tip += [f"差（{d['kind']}）: {d['text']}" for d in x["diffs"] if sid in d["sites"] or d["kind"] == "記載なし"]
            fl["nodes"].append({"id": nid, "type": t, "lane": lane, "label": n.get("label", ""), "detail": n.get("detail", ""),
                                "systems": [mdl.group_name(y) for y in mdl.systems(sid, k) if y.startswith("Y") and n.get("lane_kind") != "system"],
                                "inferred": n.get("inferred", False)})
            meta["node"][nid] = {"kind": "node", "no": x["no"] if x else "", "cls": x["class"] if x else None, "mark": mark, "coarse": coarse,
                                 "tip": "\n".join(tip), "band": b}
        for a, b2, lbl, inf in edges:
            fl["edges"].append({"from": node_id(sid, a), "to": node_id(sid, b2), "label": lbl, "inferred": inf})
    rows, r = {}, 0
    for b in bands:
        h = 1
        for sid in sids:
            ks = local[b].get(sid, [])
            ids = [node_id(sid, k) for k in ks]
            if ids:
                sub = [n for n in fl["nodes"] if n["id"] in ids]
                es = [e for e in fl["edges"] if e["from"] in ids and e["to"] in ids]
                got = BD.assign_rows(sub, es) if len(sub) > 1 else {ids[0]: 0}
                base = min(got.values())
                for i in ids:
                    rows[i] = r + got[i] - base
                h = max(h, max(got.values()) - base + 1)
        if b not in ("START", "END"):
            x = task_res[b]
            meta["band_rows"][r] = (x["no"], x["name"])
            lid = f"TASK.{b}"
            fl["nodes"].append({"id": lid, "type": "process", "lane": "TASK", "label": x["name"]})
            meta["node"][lid] = {"kind": "task_label", "no": f"{x['no']}（{x['class']}）", "text": x["name"]}
            rows[lid] = r
            for sid in sids:                                         # 記載なし・対象外の空き枠
                if local[b].get(sid):
                    continue
                mk = x["marks"].get(sid)
                if mk not in ("—", "対象外"):
                    continue
                others = [mdl.role(s2, k) for s2, ks in (by[b].get("members") or {}).items() for k in ks]
                want = max(dict.fromkeys(others), key=others.count) if others else None   # 同じ数なら、拠点の順で先のもの（set の順は実行ごとに変わる）
                lane = lane_of.get((sid, want)) or next(l["id"] for l in fl["lanes"] if l["id"].startswith(f"{sid}."))
                pid = f"{sid}.none.{b}"
                fl["nodes"].append({"id": pid, "type": "process", "lane": lane, "label": "（記載なし）"})
                meta["node"][pid] = {"kind": "placeholder" if mk == "—" else "out", "no": x["no"], "text": "（記載なし）" if mk == "—" else "（対象外）",
                                     "tip": f"{names[sid]}: タスク {x['no']} {x['name']} は、資料に記載なし" if mk == "—" else f"{names[sid]}: 対象外"}
                rows[pid] = r
            for sid in sids:                                         # 内訳の枠
                if sid in x["grain"]["fine"] and x["grain"]["coarse"]:
                    ids = [node_id(sid, k) for k in (by[b].get("members") or {}).get(sid) or []]
                    meta["frames"].append((f"{x['no']} の内訳（{len(ids)} 作業）", ids))
        r += h
    return fl, rows, meta


def row_pages(mdl, rid, res, X, S, ctx):
    """行 1 つのページ（拠点が多ければ、基準と 3 拠点ずつ。タスクが多ければ、タスクの境で）"""
    sids = [s for s in mdl.order]
    head = sids[0]
    groups = [sids] if len(sids) <= MAX_SITES else [[head] + sids[1:][i:i + MAX_SITES - 1] for i in range(0, len(sids) - 1, MAX_SITES - 1)]
    tasks = mdl.tasks(rid)
    order = [x["id"] for x in res["tasks"]]
    chunks = [order[i:i + MAX_TASKS] for i in range(0, max(1, len(order)), MAX_TASKS)] or [[]]
    pages, problems = [], []
    for gi, ss in enumerate(groups):
        for ci, ids in enumerate(chunks):
            sub = dict(res, tasks=[x for x in res["tasks"] if x["id"] in ids])
            fl, rows, meta = page_data(mdl, rid, sub, ss, tasks, X)
            suffix = ("" if len(groups) == 1 else f"-s{gi + 1}") + ("" if len(chunks) == 1 else f"-t{ci + 1}")
            name = (res.get("label") or res["name"]) + (f"（{gi + 1}/{len(groups)}）" if len(groups) > 1 else "") + (f"［{ci + 1}/{len(chunks)}］" if len(chunks) > 1 else "")
            page = None
            for width in (X["lane"]["width"], S["lane"]["width"], S["lane"]["width"] + 80):
                S2 = copy.deepcopy(S)
                S2["lane"]["width"] = width
                c2 = dict(ctx, styles=S2)
                page = CompareDrawing(fl, c2, rows, meta, f"row-{rid}{suffix}", name).build()
                lay = {"nodes": page.nodes, "edges": page.edges, "captions": page.captions}
                bad = [m for kind, m in page_problems(lay) if kind == "fail"]
                if not bad:
                    break
            if bad:
                problems.append((name, bad[:3]))
            pages.append(page)
    return pages, problems


def overview_page(mdl, results, X, S):
    """1 ページ目: 業務種別 × 拠点の表"""
    page = BD.Page("overview", "一覧（複雑性）", S["model"])
    x0, y0 = 40, 110
    sids = mdl.order
    cw = [150, 200] + [190] * len(sids) + [110]
    head = ["業務分類", "業務種別"] + [("【基準】" if s == mdl.baseline else "") + mdl.names[s] for s in sids] + ["差の大きさ"]
    page.vertex("title", esc("拠点比較 — 業務種別ごとの複雑性（大中小）"), BD.add_style(S["text_style"], S["extensions"]["title"]), x0, 40, 900, 30)
    basis = results[0]["basis"] if results else "aligned"
    page.vertex("sub", esc(f"大中小は、{'そろえた粒度' if basis == 'aligned' else '原典の粒度'}の分岐と受け渡しの数から。セルを押すと、その業務種別の拠点比較フローへ。"
                           "決め方と境は、拠点比較.xlsx の「概要」"), S["text_style"], x0, 72, 1100, 22)
    x = x0
    for k, (h, w) in enumerate(zip(head, cw)):
        page.vertex(f"h{k}", esc(h), X["grid_head"], x, y0, w, 34)
        x += w
    y = y0 + 34
    for res in results:
        cells = [res.get("group_name") or "—", res["name"]]
        x = x0
        for k, (t, w) in enumerate(zip(cells, cw)):
            page.vertex(f"r-{res['row']}-{k}", esc(t), X["grid_cell"], x, y, w, 52,
                        attrs={"link": f"data:page/id,row-{res['row']}"} if k == 1 else None)
            x += w
        for s, w in zip(sids, cw[2:]):
            e = res["sites"].get(s) or {}
            c = e.get("complexity")
            if not c:
                txt, lv = ("対象外" if e.get("relation") == "out_of_scope" else "記載なし"), "—"
            else:
                v = c[res["basis"]]
                lv = v["level"]
                txt = f"<b>{lv}</b>　分岐 {v['br']}・受け渡し {v['hand']}（点 {v['score']}）" + ("<br>粒度差（粗く書いた側）" if e.get("grain") else "")
            page.vertex(f"c-{res['row']}-{s}", txt if c else esc(txt), X["heat"][lv], x, y, w, 52, attrs={"link": f"data:page/id,row-{res['row']}"})
            x += w
        page.vertex(f"sp-{res['row']}", esc(res.get("spread") or "—"), X["grid_cell"], x, y, cw[-1], 52)
        y += 52
    return page


def page_labels(results):
    """行ごとのページの名前: 業務種別の名前。同じ名前の行があれば「業務分類 › 業務種別」、それでも重なれば番号を足す"""
    names = [r["name"] for r in results]
    out = [f"{r.get('group_name')} › {r['name']}" if names.count(r["name"]) > 1 and r.get("group_name") else r["name"] for r in results]
    seen = {}
    for i, x in enumerate(out):
        seen[x] = seen.get(x, 0) + 1
        if out.count(x) > 1:
            out[i] = f"{x}（{seen[x]}）"
    return out


def build(cmp, write=True):
    """書き出す（write=False は、点検が作り直して比べるとき。ファイルを書かない）→ (道, ページの数, 配置の点検に当たるページ, 図の文字)"""
    mdl = Model(cmp, sites_of(cmp))
    S, X = styles()
    A._CTX["labels"] = _Plain()                                      # 出典は図に出さない（node_box が出典の部品を呼ばないように）
    ctx = {"styles": S, "manifest": {"docs": []}, "taxonomy": {"categories": []}, "hierarchy": {}, "questions": [], "today": "", "page_ids": set()}
    results = []
    for r in mdl.rows():
        res, err = try_load(result_path(cmp, r["id"]))
        if isinstance(res, dict):
            results.append(res)
    if not results:
        sys.exit("計算の結果が無い（python3 scripts/sc_compute.py を先に）")
    pages, problems = [overview_page(mdl, results, X, S)], []
    for res, label in zip(results, page_labels(results)):
        res["label"] = label                                         # ページの名前（納品物に出る。行の ID は出さない）
        ps, bad = row_pages(mdl, res["row"], res, X, S, ctx)
        pages += ps
        problems += bad
    out = cmp / "out" / FIG
    text, layout = BD.render(out.name, pages)
    if write:
        BD.write(out, pages)
    return out, len(pages), problems, text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    a = ap.parse_args()
    cmp = work_dir(a.work)
    out, n, problems, _text = build(cmp)
    print(f"{out}（{n} ページ）")
    for name, bad in problems:
        print(f"  配置の点検に当たるページ: {name}（{bad[0]}）")


if __name__ == "__main__":
    main()
