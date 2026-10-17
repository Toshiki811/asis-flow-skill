#!/usr/bin/env python3
"""build_activity.py — プログラムのアクティビティ図（activities/*.activity.json）を、draw.io に書き出す。

  python3 scripts/build_activity.py -w <作業フォルダ>          # out/アクティビティ図.drawio

  1ページ目「システムと起動点の一覧」: システム一覧と起動点一覧（システム一覧は、業務マトリクス一覧の同じ名前のシートと、同じ中身。起動点一覧は、ここだけ）。番号から各ページへ飛べる
  2ページ目以降: アクティビティ図を1枚ずつ（番号の順）。レーンは 操作者／プログラム（システム名）／外部。箱・ひし形の右下（枠の外）に、もとのソースの場所（モジュールと行）
配置・線・確認事項の番号・出典の出し方は、業務フローの図（build_drawio.py）と同じ部品を使う。形だけ UML のアクティビティ図にする
（開始は黒丸、終了は二重丸、処理は角の丸い箱、判断はひし形、別ページの処理の呼び出しは左右に縦線）。
形・色・字・寸法・凡例の文は、テンプレート（assets/アクティビティ図_テンプレート.drawio）のもの。harvest_template.py が assets/activity_styles.json に抜き出す。
分かれた道を横に並べる配置は、機械が決める。
PNG は check_layout.py が書き出す（out/png/アクティビティ図_pNN_act-A-02-1.png）。
"""
import argparse
import datetime
import pathlib
import re
import sys

import build_drawio as B
from check_layout import page_problems
from code_struct import TIER_LABEL, LANG_LABEL, cite
from common import ASSETS, collect_questions, flow_key, load, load_flows, manifest, try_load, unlinked_note, work_dir
from systems import Ledger, has_ledger, placed

OUT = "アクティビティ図.drawio"
ENTRY_SHOW = {"きっかけ": "起動方法"}                                  # 起動点の表の見出し（様式の印 {きっかけ} は、そのまま）
STYLES = "activity_styles.json"                                      # アクティビティ図のテンプレートから抜いたスタイル
COLS = 4                                                             # 分岐の先を横に並べる列の数（1レーンに、これまで）


DRAFT = "【下書き】"


def load_activities(work):
    """activities/*.activity.json を、番号の順に（台帳で、図にしない手続きとしたものの、前に作った図は除く）"""
    out, off = [], set()
    if has_ledger(work):
        led = Ledger(pathlib.Path(work))
        off = {u["activity"] for u in led.units if u["kind"] == "shared" and led.off(u) is not None}
    for f in sorted((pathlib.Path(work) / "activities").glob("*.activity.json")):
        a = try_load(f)[0]
        if isinstance(a, dict) and a.get("flow_id") and a["flow_id"] not in off:
            out.append(a)
    key = lambda a: [int(x) for x in a["flow_id"].split("-")[1:] if x.isdigit()]
    return sorted(out, key=key)


def text_box(page, cid, x, y, w, lines, style):
    h = 16 + sum(max(1, B.n_lines(B.html.unescape(ln.replace("<b>", "").replace("</b>", "")), w - 14, 10)) for ln in lines) * 14
    page.vertex(cid, "<br>".join(lines), style, x, y, w, h)
    return y + h


class ActivityDrawing(B.FlowDrawing):
    """アクティビティ図1枚。データの形と線の引き方は業務フローと同じ。形・題・凡例・情報欄を差し替え、分岐の先を横に並べる（レーンの中を列に分ける）"""

    def __init__(self, act, ctx, cap=COLS):
        from review import figure_mark
        mark = figure_mark(act)                                      # 描く前の activity.json の中身の印（照合に出す版と、書き出した図をそろえる）
        look = {"system": "role", "external": "system"}             # 色: プログラムのレーンは担当と同じ配色、外部は灰色（破線の箱）
        act = dict(act, lanes=[dict(l, kind=look.get(l.get("kind"), l.get("kind"))) for l in act["lanes"]])
        self.cap = cap                                               # 1レーンの列の数の上限（1 なら、列を分けない＝縦一列）
        super().__init__(act, ctx)
        self.A = self.S                                              # アクティビティ図のテンプレートから抜いたスタイル一式（開始・終了・凡例の文も、ここ）
        self.draft = sum(1 for n in act["nodes"] if n.get("unnamed"))    # 名前の付いていない箱の数（あれば、下書き）
        self.page = B.Page(f"act-{act['flow_id']}", (DRAFT if self.draft else "") + f"{act['flow_id']} {act['name']}"[:60], self.S["model"])
        self.page.mark = mark

    def split_cols(self):
        return B.assign_cols(self.main, self.E, self.cap) if self.cap > 1 else ({}, None)

    def draw_lanes(self):
        page, S, L = self.page, self.S, self.L
        for i, l in enumerate(self.lanes):
            page.vertex(f"lane-{l['id']}-h", B.esc(l["name"]), self.pal[i]["header_style"], self.lane_x(i), L["header_y"], self.lane_w(i), L["header_h"])
            page.vertex(f"lane-{l['id']}-c", "", self.pal[i]["container_style"], self.lane_x(i), L["content_y"], self.lane_w(i), self.bottom - L["content_y"])
        note = f"{DRAFT}名前の付いていない箱が {self.draft} 個ある。納品物ではない　" if self.draft else ""
        page.vertex("title", B.esc(f"{note}{self.fl['flow_id']}　{self.fl['name']}（アクティビティ図）"), B.add_style(S["text_style"], self.ext["title"]),
                    L["x0"], L["header_y"] - 52, max(sum(self.vw), 600), 30)
        page.vertex("to-overview", "◀ システムと起動点の一覧へ", B.add_style(S["text_style"], self.ext["link_text"]), L["x0"], L["header_y"] - 24, 200, 18,
                    attrs={"link": "data:page/id,systems"})

    def node_style(self, n):
        if n["type"] in ("start", "end"):
            return self.A[n["type"]]
        st = super().node_style(n)
        return B.add_style(st, self.A["process_extra"]) if n["type"] == "process" else st

    def measure(self):
        super().measure()
        w, h = self.A["start_end_size"]
        for n in self.main:
            if n["type"] in ("start", "end"):
                self.geo[n["id"]][2:] = [w, h]
                self.label[n["id"]] = ""

    def draw_nodes(self):
        super().draw_nodes()
        for n in self.main:                                          # 終了の種類（途中で終わる）は、丸の右に小さく書く（丸の下は、横の線が走る）
            if n["type"] == "end" and n.get("label") not in (None, "", "終了"):
                g = self.geo[n["id"]]
                text = re.sub(r"^終了（(.+)）$", r"\1", n["label"])
                self.page.vertex(f"cap-{n['id']}", B.small(text, self.S), self.S["text_style"] + "align=left;", g[0] + g[2] + 5, g[1] + (g[3] - 14) / 2, 84, 14)

    def build(self):
        self.measure()
        self.place()
        self.draw_lanes()
        self.draw_nodes()
        self.draw_files()
        self.draw_edges()
        self.draw_badges()
        px = self.lx(len(self.vw)) + self.S["legend_gap"]
        w = self.A["legend_width"]
        y = text_box(self.page, "legend-a", px, self.L["header_y"], w, self.A["legend_lines"], self.A["legend_style"])
        text_box(self.page, "info", px, y + 20, w, info_lines(self.fl, self.ctx), self.S["info_style"])
        B.anchors(self.page, self.ext)
        return self.page


def activity_page(act, ctx, note=print):
    """アクティビティ図1枚。分岐の先を横に並べて描き、配置の点検が FAIL にする所（箱の重なり・線が箱を貫く・線どうしの重なり）が
    1つでもあれば、列を減らして描き直す（最後は、列を分けない縦一列）。枠の外に出した出典の字の重なりも、同じ点検で見る"""
    cap = COLS
    while True:
        d = ActivityDrawing(act, ctx, cap)
        page = d.build()
        bad = [text for kind, text in page_problems({"nodes": page.nodes, "edges": page.edges, "captions": page.captions}) if kind == "fail"]
        if not bad or cap == 1:
            return page
        cap = max(1, min(cap, max(d.ncol)) - 1)                      # 使った列より1つ少なく
        note(f"  {act['flow_id']}: 横に並べると、線・箱が重なる所がある（{len(bad)} か所）。列を {cap} つまでにして描き直す")


def info_lines(act, ctx):
    led, lab = ctx["ledger"], B.labels()
    st = led.structs.get(act["project"]) or {}
    procs = {p["id"]: p for p in st.get("procedures") or []}
    out = [f"<b>{B.esc(act['flow_id'])}　{B.esc(act['name'])}</b>"]
    if act.get("summary"):
        out.append(B.esc(act["summary"]))
    out.append(f"システム: {B.esc(act['system']['name'])}")
    out.append(f"起動方法: {B.esc(act.get('trigger', ''))}")
    used = [f"{F} {n['id']}" for F, n in ctx["boxes"].get(act["flow_id"], [])]
    callers = [a["flow_id"] for a in ctx["activities"] for n in a["nodes"] if n.get("link_flow") == act["flow_id"]]
    out.append("<b>使われる業務</b>")
    out.append(("業務フロー " + "・".join(used)) if used else ("ほかの処理から呼ばれる: " + "・".join(dict.fromkeys(callers))) if callers
               else unlinked_note(ctx["manifest"]) + "。確認事項を参照")
    out.append("<b>もとのコード</b>")
    for pid in act.get("scope") or []:
        p = procs.get(pid)
        if p:
            m = next((x for x in st["modules"] if x["name"] == p["module"]), {})
            doc = next((c["doc_id"] for c in cite(st, p["module"], *p["lines"])), m.get("doc_id", ""))   # 定義とコードが別の資料に分かれたモジュールは、その行のある資料
            out.append(f"{B.esc(lab.join(doc, '「' + p['module'] + '」'))}{p['lines'][0]}〜{p['lines'][1]} 行（{B.esc(p['name'])}）")
    lang = "accessmacro" if (procs.get(act.get("procedure")) or {}).get("kind") == "macro" else st.get("language", "")
    out.append(f"構造の取り出し: {TIER_LABEL.get(act.get('tier'), '')}（{LANG_LABEL.get(lang, lang)}）")
    lines = act.get("lines") or []
    n_all = sum(len(o["lines"]) for o in lines)
    n_off = sum(len(o["lines"]) for o in lines if o["state"] == "excluded")
    out.append("<b>行の網羅</b>")
    out.append(f"実行文 {n_all} 行 ＝ 図に入れた {n_all - n_off} 行 ＋ 対象外 {n_off} 行")
    why = {}
    for o in lines:
        if o["state"] == "excluded":
            why[o.get("reason", "")] = why.get(o.get("reason", ""), 0) + len(o["lines"])
    if why:
        out.append("対象外の内訳: " + "、".join(f"{B.esc(k)} {v} 行" for k, v in why.items()))
    folded = [b for b in act.get("branches") or [] if b["state"] == "folded"]
    if folded:
        out.append(f"畳んだ分岐・繰り返し {len(folded)} か所（" + "、".join(dict.fromkeys(B.esc(b.get("reason", "")) for b in folded)) + "）")
    myq = [q for q in ctx["questions"] if q.get("flow") == act["flow_id"]]
    out.append(f"<b>確認事項 {len(myq)} 件</b>" + ("（確認事項一覧を参照）" if myq else ""))
    out += [f"{q['id']}［{B.SEV.get(q['severity'], q['severity'])}］{B.esc(q['question'][:60])}" for q in myq[:10]]
    out.append(f"作成: {ctx['today']}（asis-flow-skill）")
    del lab
    return out


TH = 'style="border:1px solid #666666;padding:4px;background:#dae8fc;text-align:left"'
TD = 'style="border:1px solid #999999;padding:4px;vertical-align:top"'


def table(head, rows):
    h = "".join(f"<th {TH}>{B.esc(c)}</th>" for c in head)
    body = "".join("<tr>" + "".join(f"<td {TD}>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f'<table style="width:100%;border-collapse:collapse;font-size:11px;background:#ffffff"><tr>{h}</tr>{body}</table>'


def systems_page(ctx):
    """1ページ目: システム一覧（業務マトリクス一覧の同じ名前のシートと、同じ中身）と、起動点一覧（業務マトリクス一覧の既定の様式には無い。起動点の一覧は、ここに置く）"""
    from build_matrix import ENTRY_HEAD, SYSTEM_HEAD, Matrix
    S, ids = ctx["styles"], ctx["page_ids"]
    mx = Matrix(ctx["work"], write=False)
    p = B.Page("systems", "システムと起動点の一覧", S["model"])
    cell = "text;html=1;whiteSpace=wrap;overflow=fill;align=left;verticalAlign=top;strokeColor=none;fillColor=none;"
    br = lambda t: B.esc(t).replace("\n", "<br>")
    p.vertex("title", "システムと起動点の一覧（アクティビティ図）", B.add_style(S["text_style"], ctx["styles"]["extensions"]["title"]), 80, 30, 900, 30)
    srows = [[br(c.text) for c in row] for row in mx.systems()[1:]]
    y = 70
    drafts = [x["flow_id"] for x in ctx["activities"] if any(n.get("unnamed") for n in x["nodes"])]
    if drafts:                                                       # 名前の付いていない箱が残る図（下書き）は、仕上がった図に紛れないように、ここにも出す
        p.vertex("drafts", B.esc(f"{DRAFT}名前の付いていない箱が残る図（納品物ではない）: " + "・".join(drafts[:30]) + (f" ほか {len(drafts) - 30} 枚" if len(drafts) > 30 else "")),
                 B.add_style(S["text_style"], "fontColor=#C00000;fontStyle=1;"), 80, y, 1400, 20)
        y += 30

    def height(rows, ncols):
        """表の高さの見積もり（列の幅はブラウザが決めるので、広めに取る。足りないと、下の行が切れる）"""
        per = max(8, 1400 / 11 / ncols * 1.6)
        return 40 + sum(12 + 15 * max(sum(max(1, -(-B.units(B.html.unescape(x)) // per)) for x in str(c).split("<br>")) for c in r) for r in rows)
    p.vertex("h1", "<b>システム一覧</b>（1行＝1システム。原典に出てくるシステムすべて）", S["text_style"], 80, y, 800, 20)
    h = height(srows, len(SYSTEM_HEAD))
    p.vertex("t1", table(SYSTEM_HEAD, srows), cell, 80, y + 25, 1400, h)
    y += 25 + h + 30
    erows, at = [], ENTRY_HEAD.index("アクティビティ図")
    for row in mx.entries()[1:]:
        r = [br(c.text) for c in row]
        if r[at] in ids:
            r[at] = f'<a href="data:page/id,act-{r[at]}">{r[at]}</a>'
        erows.append(r)
    p.vertex("h2", "<b>起動点一覧</b>（1行＝プログラムの起動点1つ。ボタン・イベント・バッチなど、プログラムの外から始まる所）", S["text_style"], 80, y, 900, 20)
    h = height([[br(c.text) for c in row] for row in mx.entries()[1:]], len(ENTRY_HEAD))
    p.vertex("t2", table([ENTRY_SHOW.get(x, x) for x in ENTRY_HEAD], erows), cell, 80, y + 25, 1400, h)
    B.anchors(p, S["extensions"])
    return p


def context(work, acts, styles=None, write_questions=True):
    """描くのに要るもの一式（様式・台帳・確認事項・業務フローの側の箱）。write_questions=False は点検用（questions.json を書かない）"""
    m = manifest(work)
    boxes = {}
    for fl in load_flows(work, []):
        for n in placed(fl):
            boxes.setdefault(n["activity"], []).append((flow_key(fl), n))
    return {"styles": styles or load(ASSETS / STYLES), "manifest": m, "taxonomy": {},
            "hierarchy": {}, "questions": collect_questions(work, write_questions), "today": datetime.date.today().isoformat(), "page_ids": {x["flow_id"] for x in acts},
            "page_prefix": "act-", "work": work, "ledger": Ledger(work, m), "boxes": boxes, "activities": acts}


def figure(work, acts, styles, write_questions=True):
    """アクティビティ図のページの並び（システムと起動点の一覧 ＋ 図ごとに1ページ）"""
    ctx = context(work, acts, styles, write_questions)
    return [systems_page(ctx)] + [activity_page(x, ctx) for x in acts]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--styles", default=str(ASSETS / STYLES), help="アクティビティ図のテンプレートから抜いたスタイル（harvest_template.py --activity の出力）")
    a = ap.parse_args()
    work = work_dir(a.work)
    if not isinstance(load(a.styles), dict) or "legend_lines" not in load(a.styles):
        sys.exit(f"{a.styles} が無い・アクティビティ図のスタイルでない（python3 scripts/harvest_template.py で、テンプレートから作り直す）")
    if not has_ledger(work):
        sys.exit("structure/systems.json が無い（プログラムのソースが無ければ、アクティビティ図は作らない）")
    acts = load_activities(work)
    if not acts:
        sys.exit("activities/*.activity.json が無い（python3 scripts/activity.py apply -w <作業フォルダ> --all）")
    pages = figure(work, acts, load(a.styles))
    path = work / "out" / OUT
    B.write(path, pages)
    print(f"{path}  （{len(pages)} ページ: システムと起動点の一覧 ＋ アクティビティ図 {len(acts)}）")


if __name__ == "__main__":
    main()
