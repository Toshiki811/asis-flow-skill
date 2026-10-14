#!/usr/bin/env python3
"""build_drawio.py — flows/*.flow.json を、所定テンプレートのスタイルで draw.io に書き出す。

  python3 scripts/build_drawio.py -w <作業フォルダ>                    # out/業務フロー_as-is.drawio
  python3 scripts/build_drawio.py -w <作業フォルダ> -o out/xxx.drawio     # 相対パスは作業フォルダから
  python3 scripts/build_drawio.py -w <作業フォルダ> --split            # 最上位の業務分類ごとにファイルを分ける

XML は手で書かない。フローを直すときは flow.json（直しの一覧で作ったフローは、直しの一覧）を直して、これを実行し直す（draw.io 側を手で直すと JSON と食い違う）。
  1ページ目「全体構成」: 業務分類の階層（各フローのページへのリンク・結び付く資料・業務どうしの「続く」）と、資料の一覧・役割・資料間の関係
  2ページ目以降: 業務分類ごとに1ページ。レーン＝列、時系列＝上から下。下位フローは別ページにし、親のノードからリンクする
  ノード: 枠の外（右下）に、出典「業務マニュアル p.3」（資料の題名とページ）を小さい字で出し（切らない）、属性に出典・引用・確認事項番号を持つ（マウスを乗せると出る）。
          ページ番号は資料に刷られた番号（Excel はシート名とセル）。ファイルの何枚目かは出さない
  推定で置いた箇所は赤い破線。確認事項の番号（Q-012）をノードの右肩に付ける
スタイルと寸法は assets/template_styles.json（harvest_template.py がテンプレートから抜いたもの）だけを使う。
配置を out/<名前>.layout.json にも書く（check_layout.py と make_compare.py が使う）。
"""
import argparse
import collections
import copy
import datetime
import html
import math
import pathlib
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET

from common import (ASSETS, HIERARCHY_LABEL, STEP_TYPES, collect_questions, docs_by_id, labels, load, load_flows, manifest,
                    question_ids, save, src_label, tree_nodes, work_dir)
from validate import blocking
from common import REL_LABEL as REL, ROLE_LABEL as ROLE, SEV_LABEL as SEV

GAP = 46                  # 行と行の間（線が横に走る帯）
EXIT = {"bottom": (0.5, 1), "left": (0, 0.5), "right": (1, 0.5)}     # 線が出る辺 → draw.io の exitX・exitY
SLOT_Y = (0.5, 0.72, 0.28, 0.9, 0.1)                                 # 同じ辺から何本も出すとき、出口を上下にずらす（辺の高さに対する比）
SLOT_DX = 6                                                          # 同じく、ノードの横へ出る長さを1本ごとに伸ばす（縦の線が重ならないように）
LABEL_FONT = 10                                                      # 線のラベルの字の大きさ（テンプレートの flow_label の fontSize）
LABEL_ALIGN = {"left": "align=right;", "right": "align=left;", "bottom": ""}   # 左右から出る線のラベルは、ノードの外へ向けて伸ばす（ノードの文字を隠さない）
ENTRY = {"top": (0.5, 0), "left": (0, 0.5), "right": (1, 0.5)}      # 線が入る辺 → entryX・entryY
TRACK_NEAR, TRACK_PAD = 5, 12                                        # 同じ通り道とみなす近さ／端どうしがこれより近いと、1本につながって見える
TOP_DX = 24                                                          # 上から入る線が、真上のノードから出る別の線と重なるとき、入る所を横にずらす幅
SIDE_MAX = 3                                                         # 列を分けた図: 条件に当たったときだけ行う処理が、箱いくつまでなら右の列に出すか
COL_GAP, COL_PITCH = 56, 9                                           # 列を分けた図: 行と行の間／そこを横に走る線どうしの間
COL_PAD = 25                                                         # 列を分けた図: 帳票の無い列の、箱の右の余白（横へ出る線の分）
CAP_FONT, CAP_LINE = 9, 11                                           # 出典（枠の外）の字の大きさ・行の送り
CAP_MAX = 200                                                        # 出典の1行の幅の上限（px）。超える行は折る（切らない）
CAP_DX, CAP_DY, CAP_PAD = 8, 2, 12                                   # 出典の置き場所: 形の中心から右へ・形の下端から下へ／列の右に残す余白
BASIS = {"folder": "フォルダ", "toc": "目次", "numbering": "見出し番号", "outline": "しおり", "sections": "セクション",
         "sheets": "シート構成", "cross_ref": "資料間の参照", "filename": "ファイル名", "inferred": "内容からの推定"}


CONTROL = dict.fromkeys(set(range(32)) - {9, 10, 13}, " ")           # XML に書けない制御文字（PowerPoint の行内改行 \x0b など）


def clean(t):
    return str(t if t is not None else "").translate(CONTROL)


def esc(t):
    return html.escape(clean(t), quote=False)


def units(s):
    """文字幅（全角＝1、半角＝0.55）"""
    return sum(1.0 if unicodedata.east_asian_width(c) in "WFA" else 0.55 for c in str(s))


def n_lines(text, px, font):
    if not text:
        return 0
    per = max(1.0, px / font)
    return sum(max(1, math.ceil(units(p) / per)) for p in str(text).split("\n"))


def src_parts(sources):
    """出典を [(資料ID, [書き方…])] にする。同じ資料のページはまとめる（p.3,4）。ページは刷られた番号（common.Labels が引く）"""
    lab, by = labels(), {}
    for s in sources or []:
        ts = by.setdefault(s.get("doc_id"), [])
        if s.get("page") is None:
            continue
        cell = lab.code_lines(s["doc_id"], s["page"], s.get("lines")) or (lab.cell(s["doc_id"], s["page"], s.get("quote", "")) if s.get("quote") else "")
        t = lab.short(s["doc_id"], s["page"]) + (f" {cell}" if cell else "")
        if t not in ts:
            ts.append(t)
    out = []
    for d, ts in by.items():
        parts = []
        for unit in ("p.", "スライド "):
            nums = [t[len(unit):] for t in ts if t.startswith(unit)]
            if nums:
                parts.append(unit + ",".join(nums))
        out.append((d, parts + [t for t in ts if not t.startswith(("p.", "スライド "))]))
    return out


# 半角の字の幅（Helvetica・Arial。1000 分の 1 em）。出典の字（枠の外）の幅を測るのに使う
ASCII_W = dict(zip(map(chr, range(32, 127)), (
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,              # 空白 ! " # $ % & ' ( ) * + , - . /
    556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,              # 0〜9 : ; < = > ?
    1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,             # @ A〜O
    667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,              # P〜Z [ \ ] ^ _
    333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,              # ` a〜o
    556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584)))                 # p〜z { | } ~


def text_px(s, font):
    """文字の幅（px）。半角は字ごとの幅、全角は字の大きさ"""
    return sum(ASCII_W[ch] / 1000 if ch in ASCII_W else 1.0 if unicodedata.east_asian_width(ch) in "WFA" else 0.6 for ch in str(s)) * font


def wrap(line, px, font=CAP_FONT):
    """1行を、幅 px に収まるように折る。空白の所で折り、1語が長ければ字で折る（切らない）"""
    out, cur = [], ""
    for word in re.findall(r"\S+\s*", str(line)) or [str(line)]:
        if cur and text_px((cur + word).rstrip(), font) > px:
            out.append(cur.rstrip())
            cur = ""
        while text_px(word.rstrip(), font) > px:                     # 1語で幅を超える: 字で折る
            k = next((i for i in range(1, len(word) + 1) if text_px(word[:i], font) > px), len(word)) - 1
            out.append(word[:max(1, k)])
            word = word[max(1, k):]
        cur += word
    return out + ([cur.rstrip()] if cur.strip() else [])


D_LINE, D_EDGE, D_MAX = 12, 5, 170                                   # ひし形の中: 1行の高さ・行の端から辺までの余白・高さの上限
NO_HEAD, NO_TAIL = "、。，．）」』】？！：；ー〜…・", "（「『【"              # 行の頭に置かない字・行の末に置かない字


def tokens(text):
    """折ってよい所で分ける。半角の語・数字はまとめ（長い語は 14 字ごと）、行の頭（末）に置かない字は、前（後ろ）の字に付ける"""
    out = []
    for t in re.findall(r"[!-~]{1,14}|\s+|.", str(text)):
        if out and (t.isspace() or t[0] in NO_HEAD or out[-1][-1] in NO_TAIL):
            out[-1] += t
        else:
            out.append(t)
    return out


def break_cost(a, b):
    """a と b の間で折るときの悪さ。0＝文節の切れ目らしい（ひらがな・読点・閉じかっこ・空白のあとで、次がひらがなでない）、
    3＝ひらがなが続く所、14＝語の途中らしい（漢字・カタカナの並びの中、送りがなの前）"""
    kana = "ぁ" <= b[0] <= "ん"
    if a[-1] in "、。）」』】 " or "ぁ" <= a[-1] <= "ん":
        return 3 if kana and a[-1] not in "、。 " else 0
    return 14


def fit_lines(text, widths, font, back):
    """文字を、行ごとの幅（widths）に、前から詰める。(行, 折った所の悪さの合計)。行の数がちょうど len(widths) にならなければ None。
    折る所は、手前 back 個までに文節の切れ目らしい所があれば、そこにする"""
    toks, lines, bad, start = tokens(text), [], 0, 0
    while start < len(toks):
        if len(lines) == len(widths):
            return None
        end = start
        while end < len(toks) and text_px("".join(toks[start:end + 1]).rstrip(), font) <= widths[len(lines)]:
            end += 1
        if end == start:                                             # 1語だけで、その行の幅を超える
            return None
        if end < len(toks):
            end = next((j for j in range(end, max(start, end - 1 - back), -1) if not break_cost(toks[j - 1], toks[j])), end)
            bad += break_cost(toks[end - 1], toks[end])
        lines.append("".join(toks[start:end]).strip())
        start = end
    return (lines, bad) if len(lines) == len(widths) else None


def diamond_lines(label, dw, dh, id_px, font):
    """ひし形（幅 dw）に収まるように、条件の文を行に分ける。(高さ, 行)。
    ひし形は上下が細いので、行ごとに使える幅が違う（番号の行が、いちばん上）。低い形・少ない行・文節で折れる分け方を選ぶ。
    番号の行の分、条件は下に寄る。下に空の行を 1 つ足すと、条件がまん中（いちばん広い所）に来るので、それも試す。
    収まらなければ、上限の高さで、折るのを描画に任せる"""
    best = None
    for h in range(int(dh), D_MAX + 1, 4):
        for k, pad in ((k, pad) for k in range(1, 7) for pad in (0, 1)):
            n = 1 + k + pad
            if n * D_LINE > h - 24:                                  # 上下の角には、字を入れない
                continue
            ws = [dw * (1 - 2 * (abs(i - (n - 1) / 2) * D_LINE + D_EDGE) / h) - 2 for i in range(1 + k)]
            for back in (3, 0) if ws[0] >= id_px else ():            # 文節で折る分け方と、幅いっぱいまで詰める分け方の、よいほう
                got = fit_lines(label, ws[1:], font, back)
                if not got:
                    continue
                rows, bad = got
                cost = h + 8 * k + bad + (12 if k > 1 and text_px(rows[-1], font) < 3 * font else 0)   # 最後の行に 1〜2 字だけ残るのは避ける
                if best is None or cost < best[0]:
                    best = (cost, h, rows + ["\u00a0"] * pad)
    return best[1:] if best else (D_MAX, [str(label)])


def src_lines(sources, px=CAP_MAX):
    """枠の外に出す出典: 「研修資料 スライド 3,4」「請求管理「手順」シート C2」（資料は題名で書く）。1か所1行で、切らない（幅 px を超える行は、折る）"""
    lab, out = labels(), []
    for d, parts in src_parts(sources):
        for t in parts or [""]:
            out += wrap(lab.join(d, t) if t else lab.title(d), px)
    return out


def paren(t):
    """補足を、かっこ書きにする（動詞で終わる箱の名前と、見分けられるように）。全体がすでにかっこで包んであれば、重ねない"""
    t = str(t).strip()
    if t.startswith("（") and t.endswith("）"):
        depth = 0
        for i, ch in enumerate(t):
            depth += (ch == "（") - (ch == "）")
            if depth == 0 and i < len(t) - 1:
                break
        else:
            return t
    return f"（{t}）"


def tied_docs(cat, hier):
    """その業務分類に結び付く資料 [(資料ID, 役割, 詳しくしているか)]。業務の木の refs から引く。無ければ割り当てたページの資料"""
    roles = {d.get("doc_id"): d.get("role") for d in hier.get("docs") or []}
    node = next((n for n, _p in tree_nodes(hier.get("tree")) if n.get("id") == (cat or {}).get("hierarchy_id")), {})
    out = []
    for r in node.get("refs") or []:
        if r.get("doc_id") not in [x[0] for x in out]:
            out.append((r["doc_id"], roles.get(r["doc_id"]), r.get("rel") == "details"))
    for ps in (cat or {}).get("pages") or []:
        if ps.get("doc_id") not in [x[0] for x in out]:
            out.append((ps["doc_id"], roles.get(ps["doc_id"]), False))
    return out


def small(text, S):
    css = S["extensions"]["source_font"]
    size = css.split("font-size:")[1].split(";")[0]
    color = css.split("color:")[1].split(";")[0]
    return f'<font style="font-size:{size}" color="{color}">{esc(text)}</font>'


def tag(text, S):
    """札（使うシステム）と、別の図への番号の字。システム連携の線と同じ青"""
    css = S["extensions"].get("system_tag_font", "font-size:10px;color:#0066CC")
    size = css.split("font-size:")[1].split(";")[0]
    color = css.split("color:")[1].split(";")[0]
    return f'<font style="font-size:{size}" color="{color}">{esc(text)}</font>'


class Page:
    def __init__(self, did, name, model_attrs):
        self.id, self.name = did, name
        self.diagram = ET.Element("diagram", {"name": name, "id": did})
        model = ET.SubElement(self.diagram, "mxGraphModel", model_attrs)
        self.root = ET.SubElement(model, "root")
        ET.SubElement(self.root, "mxCell", {"id": "0"})
        ET.SubElement(self.root, "mxCell", {"id": "1", "parent": "0"})
        self.ids = {"0", "1"}
        self.nodes, self.edges = {}, []
        self.captions = {}                                           # ノード → 枠の外に出した出典の場所 [x, y, 幅, 高さ]
        self.b = [1e9, 1e9, -1e9, -1e9]

    def uid(self, base):
        i, cand = 1, base
        while cand in self.ids:
            i += 1
            cand = f"{base}-{i}"
        self.ids.add(cand)
        return cand

    def grow(self, x, y, w=0, h=0):
        self.b = [min(self.b[0], x), min(self.b[1], y), max(self.b[2], x + w), max(self.b[3], y + h)]

    def vertex(self, cid, value, style, x, y, w, h, parent="1", attrs=None):
        cid = self.uid(cid)
        geo = {"x": f"{x:g}", "y": f"{y:g}", "width": f"{w:g}", "height": f"{h:g}", "as": "geometry"}
        if attrs is None:
            c = ET.SubElement(self.root, "mxCell", {"id": cid, "value": value, "style": style, "vertex": "1", "parent": parent})
        else:
            o = ET.SubElement(self.root, "object", dict({"label": value, "id": cid}, **{k: clean(v) for k, v in attrs.items()}))
            c = ET.SubElement(o, "mxCell", {"style": style, "vertex": "1", "parent": parent})
        ET.SubElement(c, "mxGeometry", geo)
        if parent == "1":
            self.grow(x, y, w, h)
        return cid

    def edge(self, cid, value, style, source, target, points=None, label=None, parent="1"):
        cid = self.uid(cid)
        a = {"id": cid, "style": style, "edge": "1", "parent": parent, "source": source, "target": target}
        if value:
            a["value"] = value
        c = ET.SubElement(self.root, "mxCell", a)
        g = ET.SubElement(c, "mxGeometry", {"relative": "1", "as": "geometry"})
        if label:
            g.set("x", f"{label[0]:g}")
            ET.SubElement(g, "mxPoint", {"x": f"{label[1]:g}", "y": f"{label[2]:g}", "as": "offset"})
        if points:
            arr = ET.SubElement(g, "Array", {"as": "points"})
            for x, y in points:
                ET.SubElement(arr, "mxPoint", {"x": f"{x:g}", "y": f"{y:g}"})
                self.grow(x, y)
        return cid


def add_style(base, extra):
    return base if not extra else (base if base.endswith(";") else base + ";") + extra


def legend(page, S, x, y, sample_proc):
    """テンプレートの凡例をそのまま置き、テンプレートに無い表現（推定・確認事項・別ページ・出典）の説明を足す。戻り値は下端の y"""
    items = S["legend"]
    root = next((i for i in items if i["parent"] == "1"), None)
    if root is None:
        return y
    extra_h = 233
    text_style = next((i["style"] for i in items if i["style"].startswith("text")), S["text_style"])
    for it in items:
        g = dict(it["geometry"])
        if it is root:
            g["x"], g["y"] = f"{x:g}", f"{y:g}"
            g["height"] = f"{float(g['height']) + extra_h:g}"
        a = {"id": it["id"], "style": it["style"], "parent": it["parent"]}
        if it["value"]:
            a["value"] = it["value"]
        if it["vertex"]:
            a["vertex"] = "1"
        if it["edge"]:
            a["edge"] = "1"
            a["source"], a["target"] = it["source"], it["target"]
        page.ids.add(it["id"])
        c = ET.SubElement(page.root, "mxCell", a)
        g["as"] = "geometry"
        ET.SubElement(c, "mxGeometry", g)
    h0 = float(root["geometry"]["height"])
    w0 = float(root["geometry"]["width"])
    page.grow(x, y, w0, h0 + extra_h)
    rid, yy = root["id"], h0 - 20
    ext = S["extensions"]
    rows = [
        (add_style(sample_proc, ext["inferred"]), "推定", 90, 36, "推定で置いた箇所（原典に記載なし。確認事項を参照）"),
        (ext["badge"], "Q-001", 46, 18, "確認事項の番号（確認事項一覧と対応）"),
        (add_style(sample_proc, ext["subprocess"]), "別フローへ", 90, 36, "下位フロー・続く業務（別ページ。クリックで移動）"),
        (None, "", 0, 0, "ノードの右下（枠の外）の灰色の字「業務マニュアル p.3」＝出典。根拠にした資料の題名と、資料に刷られたページ番号（Excel はシート名とセル）。資料のファイルは下の一覧"),
        (None, "", 0, 0, "箱の中の青い字「［◯◯システム］」＝その作業で使うシステム（業務マトリクス一覧の「システム一覧」と対応）。"
                         "「▶ A-02-1」＝そのプログラムの中身を描いたアクティビティ図の番号"),
    ]
    for style, val, w, h, text in rows:
        if style:
            page.vertex(f"legend-x-{len(page.ids)}", val, add_style(style, "fontSize=10;"), 30, yy + (40 - h) / 2, w, h, parent=rid)
            page.vertex(f"legend-x-{len(page.ids)}", text, text_style, 140, yy + 5, w0 - 170, 30, parent=rid)
        else:
            page.vertex(f"legend-x-{len(page.ids)}", text, text_style, 30, yy + 5, w0 - 60, 30, parent=rid)
        yy += 48
    return y + h0 + extra_h


def info_box(page, S, x, y, w, lines):
    body = "<br>".join(lines)
    h = 16 + sum(max(1, n_lines(html.unescape(ln.replace("<b>", "").replace("</b>", "")), w - 14, 10)) for ln in lines) * 14
    page.vertex("info", body, S["info_style"], x, y, w, h)
    return y + h


def anchors(page, ext):
    """書き出し範囲を決める見えない点（make_compare.py が、画像上のノードの位置を正確に求めるため）"""
    b = page.b
    page.vertex("anchor-tl", "", ext["anchor"], b[0] - 20, b[1] - 20, 1, 1)
    page.vertex("anchor-br", "", ext["anchor"], b[2] + 20, b[3] + 20, 1, 1)


def lane_palettes(lanes, S):
    """レーンごとの色。システムのレーンは専用の色、ほかはテンプレートの色を順に使う"""
    pal, k = [], 0
    for l in lanes:
        if l.get("kind") == "system":
            pal.append(S["system_palette"])
        else:
            pal.append(S["palettes"][k % len(S["palettes"])])
            k += 1
    return pal


def back_edges(roots, adj):
    """戻りの線（たどっている途中のノードへ戻る線）の番号。roots の順に、深さ優先でたどって見つける"""
    color, back = {}, set()
    for r in roots:
        if color.get(r):
            continue
        color[r] = 1
        stack = [(r, iter(adj[r]))]
        while stack:
            u, it = stack[-1]
            for v, k in it:
                if color.get(v) == 1:
                    back.add(k)
                elif not color.get(v):
                    color[v] = 1
                    stack.append((v, iter(adj[v])))
                    break
            else:
                color[u] = 2
                stack.pop()
    return back


def flow_order(main, edges, after=None):
    """戻りの線を除いた向きに、ノードを並べる（前のノードがすべて並んだものから、flow.json に書かれた順）。
    戻り値は、(並び, 戻りの線の番号, ノード → 前のノード)。after は、線が無くても先に並べる相手（ノード → [相手]）"""
    mids = [n["id"] for n in main]
    idx = {i: k for k, i in enumerate(mids)}
    adj = {i: [] for i in mids}
    for k, e in enumerate(edges):
        adj[e["from"]].append((e["to"], k))
    back = back_edges([n["id"] for n in main if n["type"] == "start"] + mids, adj)
    indeg, preds, more = {i: 0 for i in mids}, {i: [] for i in mids}, {i: [] for i in mids}
    for k, e in enumerate(edges):
        if k not in back:
            indeg[e["to"]] += 1
            preds[e["to"]].append(e["from"])
    for v, ps in (after or {}).items():
        for p in ps:
            indeg[v] += 1
            preds[v].append(p)
            more[p].append(v)
    ready, order = [i for i in mids if indeg[i] == 0], []
    while ready:
        ready.sort(key=idx.get)
        u = ready.pop(0)
        order.append(u)
        for v in [v for v, k in adj[u] if k not in back] + more[u]:
            indeg[v] -= 1
            if indeg[v] == 0:
                ready.append(v)
    order += [i for i in mids if i not in order]
    return order, back, preds


def assign_rows(main, edges):
    """行（時系列）を決める: 戻りの線を除いた向きで、前のノードより下。同じレーンの同じ行には置かない。{ノード: 行}"""
    lane = {n["id"]: n["lane"] for n in main}
    order, _back, preds = flow_order(main, edges)
    row, taken = {}, set()
    for u in order:
        r = max([row[p] + 1 for p in preds[u] if p in row] or [0])
        while (lane[u], r) in taken:
            r += 1
        row[u] = r
        taken.add((lane[u], r))
    return row


def assign_cols(main, edges, cap):
    """分岐の先を横に並べる配置: レーンの中の列と、行を決める。戻り値は、({ノード: 列}, {ノード: 行})。
      ・長い方の道（ふつうの終了に着く道）は、同じ列をまっすぐ下。ほかの出口の先は、右の列へ順に
      ・条件に当たったときだけ行う短い処理（箱 SIDE_MAX 個まで）は右の列に出し、当たらない側を、まっすぐ下に
      ・右の列に出した先は、その先の全部（合流の手前まで）が入る列を取る。先に入った枝とぶつかれば、もう1つ右の列
      ・繰り返しは、中身も、抜けた先も、同じ列（抜けた先は、中身の全部より下）
      ・合流するノードは、入ってくる道のうち、いちばん左の列
    列は、1レーンに cap 個まで（超える分は、いちばん右の列に縦に並ぶ）"""
    first, back, _preds = flow_order(main, edges)
    mids = [n["id"] for n in main]
    lane = {n["id"]: n["lane"] for n in main}
    kind = {n["id"]: n["type"] for n in main}
    succ, pred = {i: [] for i in mids}, {i: [] for i in mids}
    for k, e in enumerate(edges):
        if k not in back and e["to"] not in succ[e["from"]]:
            succ[e["from"]].append(e["to"])
            pred[e["to"]].append(e["from"])
    goal = {i for i in mids if kind[i] == "end" and i == "end"} or {i for i in mids if kind[i] == "end"}
    height, good, below = {}, {}, {}
    for u in reversed(first):
        height[u] = 1 + max([height.get(v, 0) for v in succ[u]] or [0])
        good[u] = u in goal or any(good.get(v) for v in succ[u])
        below[u] = {u}.union(*[below.get(v, {v}) for v in succ[u]])  # そこからたどれるノード
    loops = {}                                                       # 繰り返しの頭 → 中身のノード
    for k in sorted(back):
        tail, head = edges[k]["from"], edges[k]["to"]
        body, todo = loops.setdefault(head, set()), [tail]
        while todo:
            x = todo.pop()
            if x != head and x not in body:
                body.add(x)
                todo += pred[x]
    after = {}
    for head, body in loops.items():                                 # 繰り返しを抜けた先（頭から抜けた先と、中から抜けて、ふつうの終了へ続く先）は、中身の全部より下。
        for x in [head] + sorted(body, key=mids.index):              # 中から抜けて、そのまま終わる枝（途中で終わる）は、抜ける判断の横に置ける
            for v in succ[x]:
                if v != head and v not in body and not (kind[v] == "end" and len(pred[v]) == 1) and not below[v] & body and (x == head or good[v]):
                    after[v] = [b for b in mids if b in body or b in after.get(v, [])]
    order, _back, preds = flow_order(main, edges, after)

    def main_way(vs):
        """同じ列をまっすぐ下へ進む先"""
        if len(vs) == 2:
            for f, v in (vs, vs[::-1]):                              # f へ飛ばす線と、f に合流する短い処理 v
                if f in below[v] and len([x for x in below[v] - below[f] if kind[x] != "end"]) <= SIDE_MAX:
                    return f
        return max(vs, key=lambda v: (good[v], height[v], -vs.index(v)))

    def depth(h, arm):
        """枝（h から先の、合流の手前まで）の、いちばん長い道のノードの数"""
        d = {}
        for x in reversed(order):
            if x in arm:
                d[x] = 1 + max([d.get(v, 0) for v in succ[x] if v in arm] or [0])
        return d.get(h, 1)

    col, row, want, taken, held, arms = {}, {}, {}, set(), {}, {}
    for u in order:
        c, base, arm, right = min(want.get(u) or [(0, {}, None, False)], key=lambda p: p[0])
        c, r = min(c, cap - 1), max([row[p] + 1 for p in preds[u] if p in row] or [0])
        busy = lambda c, r: ((lane[u], c), r) in taken or u not in arms.get(held.get(((lane[u], c), r)), {u})
        if arm is not None:                                          # 分かれた先の先頭: その先の全部（合流の手前まで）が入る所を取る
            arms[u], n = arm, depth(u, arm if right else {x for x in arm if lane[x] == lane[u]})
            while right and c < cap - 1 and any(busy(c, r + i) for i in range(n)):   # 右の列に出した先: ぶつかれば、もう1つ右の列
                c += 1
            while not right and any(busy(c, r + i) for i in range(n)) and r < len(mids) * 2:   # ほかのレーンへ出た先: ぶつかれば、下へ
                r += 1
            if not any(busy(c, r + i) for i in range(n)):
                held.update({((lane[u], c), r + i): u for i in range(n)})
        elif kind[u] == "end" and len(pred[u]) == 1:                 # 抜ける所ごとの終了の丸: 真下が埋まっていれば、右の列
            while c < cap - 1 and busy(c, r):
                c += 1
        while busy(c, r):
            r += 1
        col[u], row[u] = c, r
        taken.add(((lane[u], c), r))
        here, vs = dict(base, **{lane[u]: c}), succ[u]
        if not vs:
            continue
        keep = set(vs) if u in loops else {main_way(vs)}
        used = {lane[v]: 0 for v in vs if v in keep}                 # レーン → そのレーンで、右へいくつ出したか
        straight = set().union(*[below[m] for m in keep])
        for v in vs:
            off, arm, ctx = 0, None, here
            if v in keep or any(v in below[w] for w in vs if w != v):    # まっすぐ進む先・合流する先は、列を変えない
                if v in keep and lane[v] != lane[u] and len(vs) > 1:     # まっすぐ進む先が、ほかのレーン: そのレーンに、その先の分の場所を取る
                    arm = below[v] - set().union(*[below[w] for w in vs if w != v])
            elif lane[v] not in used:                                # そのレーンへ進む、はじめの枝: 列は変えない（u と同じレーンなら、そのまま下へ続く）
                used[lane[v]] = 0
                arm = below[v] - straight if lane[v] != lane[u] else None
            else:
                used[lane[v]] += 1
                off, arm = used[lane[v]], below[v] - straight
            home = [x for x in order if x in below[v] - straight and lane[x] == lane[u]] if lane[v] != lane[u] and v not in keep else []
            if home:                                                 # ほかのレーンへ出て、もとのレーンへ戻ってくる枝: 戻った所から先は、右の列
                used[lane[u]] = used.get(lane[u], 0) + 1
                ctx = dict(here, **{lane[u]: c + used[lane[u]]})
                want.setdefault(home[0], []).append((c + used[lane[u]], ctx, (below[v] - straight) & below[home[0]], True))
            want.setdefault(v, []).append((here.get(lane[v], 0) + off, ctx, arm, off > 0))
    return col, row


def node_box(n, S):
    """ノードの大きさと、中に書く文字、枠の外に出す出典。(幅, 高さ, 文字の HTML, 出典の行)"""
    t = n["type"]
    if t in ("start", "end"):
        w, h = S["size"]["start_end"]
        return w, h, esc(n.get("label") or ("開始" if t == "start" else "終了")), []
    cite = src_lines(n.get("sources"))                               # 出典は、枠の中に入れない（長い出典で、形が縦に伸びる・字が切れる）
    if t == "decision":
        DW, DH = S["size"]["decision"]
        size = re.search(r"fontSize=([0-9.]+)", S["node"]["decision"])
        h, rows = diamond_lines(n["label"], DW, DH, text_px(n["id"], CAP_FONT), float(size.group(1)) if size else 10)
        return DW, h, f"{small(n['id'], S)}<br>" + "<br>".join(esc(r) for r in rows), cite   # 番号は小さい字で上に（一覧から引けるように）。条件は、形に合わせて折る
    PW, PH = S["size"]["process"]
    inner = PW - 12
    if t == "subprocess":                                            # 両脇の帯の分、文字の入る幅が狭い（見積もりが広いと、文字が箱からはみ出す）
        bar = re.search(r"size=([0-9.]+)", S["extensions"]["subprocess"])
        inner = inner * (1 - 2 * float(bar.group(1) if bar else 0.1)) - 16   # 帯の内側の余白と、行頭・行末の禁則で早く折れる分
    head = f"{n['id']}: {n['label']}"
    parts, lines = [esc(head)], n_lines(head, inner, 11)
    if n.get("loop"):                                                # アクティビティ図: 繰り返す箱（ソースの条件は、図に出さない。一覧の「プログラムのノード一覧」に出る）
        parts.append(small("↻ 繰り返し", S))
        lines += 1
    if n.get("detail"):                                              # 補足は、かっこ書き（箱の名前は、動詞で終わる。その後ろに続く補足を、名前と見分ける）
        parts.append(esc(paren(n["detail"])))
        lines += n_lines(paren(n["detail"]), inner, 11)
    if n.get("folded"):                                              # アクティビティ図: 技術的な分岐・繰り返しを畳んだ箱
        parts.append(small(f"（細かい場合分けを、この箱にまとめた: {n['folded']}）", S))
        lines += n_lines(f"（細かい場合分けを、この箱にまとめた: {n['folded']}）", inner, 9)
    if t == "subprocess" and n.get("link_flow"):
        parts.append(small(f"→ 別ページ {n['link_flow']}", S))
        lines += 1
    for c in n.get("calls") or []:                                   # アクティビティ図: 同じ共通の処理を続けて呼ぶ並びを、1つにまとめた箱
        text = f"→ 別ページ {c['to']} {c.get('label', '')}（{c['times']} 回）"
        parts.append(small(text, S))
        lines += n_lines(text, inner, 9)
    for name in n.get("systems") or []:                              # 札: その作業で使うシステム
        parts.append(tag(f"［{name}］", S))
        lines += n_lines(f"［{name}］", inner, 10)
    if n.get("activity"):                                            # プログラムの起動点の箱: 中身は、この番号のアクティビティ図
        parts.append(tag(f"▶ {n['activity']}", S))
        lines += 1
    return PW, max(PH, 14 + 14 * lines), "<br>".join(parts), cite


def node_attrs(n, lane_name, qs, m, link):
    """ノードの属性（出典・引用・確認事項の番号）。マウスを乗せると出る。check_flow.py と make_compare.py も読む"""
    attrs = {"node_id": n["id"], "node_type": n["type"], "lane": lane_name}
    if n.get("sources"):
        attrs["src"] = "; ".join(src_label(m, s) for s in n["sources"])
        attrs["quote"] = " ／ ".join(s.get("quote", "") for s in n["sources"] if s.get("quote"))[:400]
        attrs["src_image"] = "; ".join(f"pages/{s['doc_id']}/p{int(s['page']):04d}.png" for s in n["sources"]
                                       if isinstance(s.get("page"), int))
    if n.get("inferred"):
        attrs["inferred"] = "1"
    if n.get("systems"):
        attrs["systems"] = "、".join(n["systems"])
    if n.get("activity"):
        attrs["activity"] = n["activity"]
    if qs:
        attrs["questions"] = ", ".join(qs)
    tip = []
    if n.get("systems"):
        tip.append("使うシステム: " + attrs["systems"])
    if n.get("activity"):
        tip.append(f"プログラムの中身: アクティビティ図 {n['activity']}（アクティビティ図.drawio）")
    if attrs.get("src"):
        tip.append("出典: " + attrs["src"])
    if attrs.get("quote"):
        tip.append("引用: " + attrs["quote"])
    if n.get("reference"):                                           # 箱に収まらない一覧・表の中身（図には出さない）
        attrs["reference"] = str(n["reference"])
        tip.append("参考: " + attrs["reference"])
    if n.get("inferred"):
        tip.append("※ 推定で置いた箇所（原典に記載なし）")
    if qs:
        tip.append("確認事項: " + ", ".join(qs))
    if tip:
        attrs["tooltip"] = "\n".join(tip)
    if link:
        attrs["link"] = link
    return attrs


def flow_info_lines(fl, ctx, title):
    """フローのページの右に置く情報欄: 根拠・開始と終了・もとの資料・確認事項"""
    used_docs = []
    for n in fl["nodes"]:
        for s in n.get("sources") or []:
            if s.get("doc_id") not in used_docs:
                used_docs.append(s.get("doc_id"))
    dmap = docs_by_id(ctx["manifest"])
    myq = [q for q in ctx["questions"] if q.get("flow") == fl["flow_id"]]
    lines = [f"<b>{esc(fl['flow_id'])}　{esc(title)}</b>"]
    if fl.get("parent"):
        lines.append(f"上位: {esc(fl['parent'])}")
    lines.append("業務分類の根拠: " + ("目次に準拠" if fl.get("basis") == "toc" else "LLM が分類（確認事項を参照）"))
    if fl.get("start_condition"):
        lines.append(f"開始: {esc(fl['start_condition'])}")
    if fl.get("end_condition"):
        lines.append(f"終了: {esc(fl['end_condition'])}")
    lines.append("<b>もとの資料</b>")
    cat = next((c for c in ctx["taxonomy"].get("categories") or [] if c.get("id") == fl["flow_id"]), None)
    tied = {d: (role, det) for d, role, det in tied_docs(cat, ctx["hierarchy"])}
    for d in used_docs + [x for x in tied if x not in used_docs]:
        role, det = tied.get(d, (None, False))
        tag = "・".join(x for x in (ROLE.get(role, ""), "この業務の詳細" if det else "", "" if d in used_docs else "出典には使っていない") if x)
        lines.append(f"{esc(labels().title(d))}{f'（{tag}）' if tag else ''}" + ("" if d in dmap else "（不明な資料）"))   # ファイルの場所は、資料一覧・出典対応表に
    lines.append(f"<b>確認事項 {len(myq)} 件</b>" + ("（確認事項一覧を参照）" if myq else ""))
    lines += [f"{q['id']}［{SEV.get(q['severity'], q['severity'])}］{esc(q['question'][:60])}" for q in myq[:12]]
    if len(myq) > 12:
        lines.append(f"… ほか {len(myq) - 12} 件")
    lines.append(f"作成: {ctx['today']}（asis-flow-skill）")
    return lines


class FlowDrawing:
    """フロー1本を1ページに描く。行を決め、大きさを測って置き場所を決めてから、レーン・ノード・帳票・線・確認事項の番号・凡例の順に置く"""

    def __init__(self, fl, ctx):
        self.fl, self.ctx = fl, ctx
        self.S = S = ctx["styles"]
        self.L, self.ext = S["lane"], S["extensions"]
        self.q_node, self.q_edge = question_ids(ctx["questions"], fl["flow_id"])
        self.title = f"{fl.get('number', '')} {fl['name']}".strip()
        self.page = Page(f"flow-{fl['flow_id']}", f"{fl['flow_id']} {fl['name']}"[:60], S["model"])
        self.lanes = fl.get("lanes") or [{"id": "L1", "name": "（担当不明）", "kind": "dept"}]
        self.li = {l["id"]: i for i, l in enumerate(self.lanes)}     # レーン → 左から何列目か
        self.pal = lane_palettes(self.lanes, S)
        self.nodes = {n["id"]: n for n in fl["nodes"]}
        self.main = [n for n in fl["nodes"] if n["type"] != "file" and n.get("lane") in self.li]   # 帳票以外（線でつながるノード）
        mids = {n["id"] for n in self.main}
        self.E = [e for e in fl.get("edges", []) if e["from"] in mids and e["to"] in mids]
        files = {n["id"] for n in fl["nodes"] if n["type"] == "file"}
        for n in fl["nodes"]:                                        # 図に置けないものは、黙って落とさない（点検では FAIL になるもの）
            if n["type"] != "file" and n["id"] not in mids:
                print(f"  {fl['flow_id']}: ノード {n['id']} のレーン「{n.get('lane')}」が lanes に無いので、図に置けない")
        for e in fl.get("edges", []):
            if e not in self.E and not {e["from"], e["to"]} & files:
                print(f"  {fl['flow_id']}: 線 {e['from']}→{e['to']} の端が図に無いので、描けない")
        self.col, row = self.split_cols()                            # ノード → レーンの中の列（業務フローは、列を分けない＝どのノードも 0）
        self.gap, self.pitch = (COL_GAP, COL_PITCH) if self.col else (GAP, 6)   # 行と行の間・そこを横に走る線どうしの間（列を分けた図は、横の線が増えるので広げる）
        self.row = row or assign_rows(self.main, self.E)
        self.ncol = [1 + max([self.col.get(n["id"], 0) for n in self.main if self.li[n["lane"]] == i] or [0]) for i in range(len(self.lanes))]
        self.v0 = [sum(self.ncol[:i]) for i in range(len(self.lanes))]   # レーン → そのいちばん左の列（図の左からの通し番号）
        self.gap_use, self.chan_use = {}, {}                         # 行の間・レーン左の通り道を、線が何本使ったか（重ならないようにずらす）
        self.tracks = []                                             # 引いた線の線分: (向き "v"・"h", その x か y, 走る範囲の端, 端, 線)

    def build(self):
        self.measure()
        self.place()
        self.draw_lanes()
        self.draw_nodes()
        self.draw_files()
        self.draw_edges()
        self.draw_badges()
        px = self.lx(len(self.vw)) + self.S["legend_gap"]            # 凡例とフローの情報は、レーンの右
        ly = legend(self.page, self.S, px, self.L["header_y"], self.S["palettes"][0]["process_style"])
        info_box(self.page, self.S, px, ly + 20, 450, flow_info_lines(self.fl, self.ctx, self.title))
        anchors(self.page, self.ext)
        return self.page

    def split_cols(self):
        """分岐の先を横に並べるときの、({ノード: レーンの中の列}, {ノード: 行})。業務フローは、列を分けない"""
        return {}, None

    def vcol(self, nid):
        """ノードの列（図の左からの通し番号。列を分けない図では、レーンの番号と同じ）"""
        return self.v0[self.li[self.nodes[nid]["lane"]]] + self.col.get(nid, 0)

    def lx(self, v):
        """列 v の左端の x（列を分けない図では、レーンの左端）"""
        return self.L["x0"] + sum(self.vw[:v])

    def lane_x(self, i):
        return self.lx(self.v0[i]) if i < len(self.lanes) else self.lx(len(self.vw))

    def lane_w(self, i):
        return sum(self.vw[self.v0[i]:self.v0[i] + self.ncol[i]])

    # ---- 大きさと置き場所 ----

    def measure(self):
        """ノードと帳票の大きさ、中に書く文字"""
        FW, FH = self.S["size"]["file"]
        self.geo, self.label, self.cap = {}, {}, {}                  # ノード → [x, y, 幅, 高さ]／中に書く文字／枠の外に出す出典の行
        for n in self.main:
            w, h, self.label[n["id"]], cite = node_box(n, self.S)
            self.geo[n["id"]] = [0, 0, w, h]
            if cite:
                self.cap[n["id"]] = cite
        self.files_of = {}                                           # プロセス → そこで使う帳票
        for f in self.fl["nodes"]:
            if f["type"] != "file":
                continue
            if f.get("attach_to") in self.geo:
                self.files_of.setdefault(f["attach_to"], []).append(f)
            else:
                print(f"  {self.fl['flow_id']}: ファイル {f['id']} の attach_to が無い・不正なので図に置けない")
        self.fgeo = {f["id"]: [0, 0, FW, max(FH, 16 + 11 * (1 + n_lines(f["label"], FW - 6, 9)))]
                     for fs in self.files_of.values() for f in fs}

    def edge_text(self, e):
        """線のラベル（確認事項の番号つき）"""
        qs = self.q_edge.get((e["from"], e["to"]), [])
        return (e.get("label") or "") + (f"〔{', '.join(qs)}〕" if qs else "")

    def place(self):
        """行の高さと y を決め、ノードをレーンの中に置く。出典は、形のすぐ下・真ん中の線の右（枠の外）。その分、行を高く、列を広くする"""
        L, row, geo = self.L, self.row, self.geo
        self.vw = [L["width"]] * sum(self.ncol)                      # 列の幅。列を分けたレーンでは、帳票の無い列を詰める
        with_file = {self.vcol(n["id"]) for n in self.main if self.files_of.get(n["id"])}
        for i, k in enumerate(self.ncol):
            for v in range(self.v0[i], self.v0[i] + k) if k > 1 else []:
                if v not in with_file:
                    self.vw[v] = min(L["width"], L["node_dx"] + self.S["size"]["process"][0] + COL_PAD)
        self.side, self.slot = self.exit_sides()                     # 線の出る辺（出典の左端と、列の幅を決めるのに要る）
        out_right = collections.Counter(self.E[k]["from"] for k, sd in self.side.items() if sd == "right")
        down = {self.E[k]["from"]: self.edge_text(self.E[k]) for k, sd in self.side.items() if sd == "bottom"}
        self.row_h = [0] * (max(row.values()) + 1 if row else 1)
        block = {}                                                   # ノード → (形と帳票の高さ, 出典の分, 出典を帳票の下に置くか, 出典の左端・幅・高さ)
        for n in self.main:
            nid, (w, h) = n["id"], geo[n["id"]][2:]
            fs = self.files_of.get(nid) or []
            stack_h = sum(self.fgeo[f["id"]][3] + 5 for f in fs) - 5 if fs else 0
            body, hang, under, cap = h, 0, False, None
            if self.cap.get(nid):
                dx = max(CAP_DX, (units(down[nid]) * LABEL_FONT + 6) / 2 + 6 if down.get(nid) else 0)   # 下へ出る線のラベル（真ん中に乗る）の右から
                if fs and out_right[nid]:                            # 帳票があり、右へ出る線もある箱: 線を出典の外へ回せないので、出典を、線の手前の幅で折る
                    self.cap[nid] = [y for x in self.cap[nid] for y in wrap(x, max(36, w / 2 - dx + 2))]
                cw = math.ceil(max(text_px(x, CAP_FONT) for x in self.cap[nid])) + 3
                cap, hang = (dx, cw, CAP_LINE * len(self.cap[nid]) + 2), CAP_DY + CAP_LINE * len(self.cap[nid]) + 2
                reach = L["node_dx"] + w / 2 + dx + cw                # 出典の右端（列の左端から）
                under = bool(fs) and reach > L["file_dx"] - 4 and (len(fs) > 1 or 5 + stack_h > h + CAP_DY - 1)   # 帳票が、形より下まである: 出典は、その下
                body = max(h, stack_h + (5 if len(fs) == 1 else 0)) if under else h
                room = reach + CAP_PAD + (6 + SLOT_DX * out_right[nid] if out_right[nid] and not fs else 0)   # 右へ出る線は、出典の右で折れる
                self.vw[self.vcol(nid)] = max(self.vw[self.vcol(nid)], room)
            block[nid] = (body, hang, under, cap)
            self.row_h[row[nid]] = max(self.row_h[row[nid]], body + hang if under else max(h + hang, stack_h))
        self.row_y, y = [], L["first_y"]
        for h in self.row_h:
            self.row_y.append(y)
            y += h + self.gap
        self.bottom = y - self.gap + 30
        self.cap_box, self.file_top = {}, {}                         # ノード → 出典の場所 [x, y, 幅, 高さ]／帳票の重なりの上端（出典を帳票の下に置くとき）
        for n in self.main:
            v, r, g = self.vcol(n["id"]), row[n["id"]], geo[n["id"]]
            body, hang, under, cap = block[n["id"]]
            top = self.row_y[r] + (self.row_h[r] - body - hang) / 2  # 形（と帳票）＋出典を、1かたまりで、行の中央に
            g[0] = self.lx(v) + (L["width"] - g[2]) / 2 if n["type"] in ("start", "end") else self.lx(v) + L["node_dx"]
            g[1] = top + (body - g[3])                                # 帳票が形より下まであるときは、形を下にそろえる（出典が、形のすぐ下に来る）
            if under:
                self.file_top[n["id"]] = top + (5 if len(self.files_of[n["id"]]) == 1 else 0)
            if cap:
                self.cap_box[n["id"]] = [g[0] + g[2] / 2 + cap[0], top + body + CAP_DY, cap[1], cap[2]]

    # ---- 描く ----

    def draw_lanes(self):
        page, S, L = self.page, self.S, self.L
        for i, l in enumerate(self.lanes):
            page.vertex(f"lane-{l['id']}-h", esc(l["name"]), self.pal[i]["header_style"], self.lane_x(i), L["header_y"], self.lane_w(i), L["header_h"])
            page.vertex(f"lane-{l['id']}-c", "", self.pal[i]["container_style"], self.lane_x(i), L["content_y"], self.lane_w(i), self.bottom - L["content_y"])
        page.vertex("title", esc(f"{self.fl['flow_id']}　{self.title}（as-is 業務フロー）"), add_style(S["text_style"], self.ext["title"]),
                    L["x0"], L["header_y"] - 52, max(sum(self.vw), 600), 30)
        page.vertex("to-overview", "◀ 全体構成へ", add_style(S["text_style"], self.ext["link_text"]), L["x0"], L["header_y"] - 24, 120, 18,
                    attrs={"link": "data:page/id,overview"})

    def node_style(self, n):
        t, i = n["type"], self.li[n["lane"]]
        if t in ("start", "end"):
            style = self.S["node"]["start_end"]
        elif t == "decision":
            style = self.S["node"]["decision"]
        elif t == "system" or self.lanes[i].get("kind") == "system":
            style = self.S["node"]["system"]
        else:
            style = self.pal[i]["process_style"]
        if t == "subprocess":
            style = add_style(style, self.ext["subprocess"])
        if n.get("inferred"):
            style = add_style(style, self.ext["inferred"])
        return style

    def draw_nodes(self):
        lane_name = {l["id"]: l["name"] for l in self.lanes}
        for n in self.main:
            linked = n["type"] == "subprocess" and n.get("link_flow") and n["link_flow"] in self.ctx["page_ids"]
            attrs = node_attrs(n, lane_name[n["lane"]], self.q_node.get(n["id"], []), self.ctx["manifest"],
                               f"data:page/id,{self.ctx.get('page_prefix', 'flow-')}{n['link_flow']}" if linked else None)
            g = self.geo[n["id"]]
            self.page.vertex(n["id"], self.label[n["id"]], self.node_style(n), g[0], g[1], g[2], g[3], attrs=attrs)
            self.page.nodes[n["id"]] = [round(v, 1) for v in g]
        style = add_style(self.S["text_style"], "align=left;verticalAlign=top;whiteSpace=nowrap;spacing=0;")
        for nid, b in self.cap_box.items():                          # 出典（枠の外・右下。切らない）
            self.page.vertex(f"src-{nid}", "<br>".join(small(x, self.S) for x in self.cap[nid]), style, *b)
            self.page.captions[nid] = [round(v, 1) for v in b]

    def draw_files(self):
        """帳票（プロセスの右に小さく。参照の点線でつなぐ）"""
        S, page = self.S, self.page
        for nid, fs in self.files_of.items():
            g = self.geo[nid]
            fy = self.file_top.get(nid, g[1] + 5 if len(fs) == 1 else self.row_y[self.row[nid]])
            for f in fs:
                fg = self.fgeo[f["id"]]
                fg[0], fg[1] = self.lx(self.vcol(nid)) + self.L["file_dx"], fy
                fy += fg[3] + 5
                attrs = {"node_id": f["id"], "node_type": "file"}
                if f.get("sources"):
                    attrs["src"] = "; ".join(src_label(self.ctx["manifest"], s) for s in f["sources"])
                    attrs["tooltip"] = "出典: " + attrs["src"]
                style = add_style(S["node"]["file"], self.ext["inferred"]) if f.get("inferred") else S["node"]["file"]
                page.vertex(f["id"], f"{esc(f['id'])}<br>{esc(f['label'])}", style, fg[0], fg[1], fg[2], fg[3], attrs=attrs)
                page.nodes[f["id"]] = [round(v, 1) for v in fg]
                a, b = (nid, f["id"]) if f.get("direction") == "out" else (f["id"], nid)
                page.edge(f"ref-{f['id']}", "", S["edge"]["ref"], a, b)

    def draw_badges(self):
        """確認事項の番号（ノードの右肩）"""
        for n in self.main:
            qs = self.q_node.get(n["id"], [])
            if qs:
                g = self.geo[n["id"]]
                self.page.vertex(f"q-{n['id']}", qs[0] + ("他" if len(qs) > 1 else ""), self.ext["badge"], g[0] + g[2] - 30, g[1] - 8,
                                 50 if len(qs) > 1 else 44, 16, attrs={"tooltip": "確認事項: " + ", ".join(qs)})

    # ---- 線。横は行と行の間、縦はレーン左の通り道を使い、ノードを貫かないようにする ----

    def rect(self, i):
        x, y, w, h = self.geo[i]
        cap = self.cap_box.get(i) if hasattr(self, "cap_box") else None
        out = cap[0] + cap[2] - self.lx(self.vcol(i)) + 6 if cap and cap[0] + cap[2] > x + w + 2 else 0   # 出典が形の右へ出ている: 右へ出る線は、その先で折る
        return {"cx": x + w / 2, "cy": y + h / 2, "top": y, "bottom": y + h, "left": x, "right": x + w,
                "lane": self.vcol(i), "row": self.row[i], "type": self.nodes[i]["type"], "cap_out": out}

    def clear(self, kind, pos, a, b, e):
        """通り道（kind: 縦 "v"・横 "h"、pos: その x か y、a〜b: 走る範囲）に、行き先の違う別の線が乗っていないか。
        同じ行き先の線は、合流するので重なってよい。端どうしが近いだけでも、1本につながって見えるので、乗っているとみなす"""
        a, b = min(a, b), max(a, b)
        return not any(k == kind and abs(p - pos) < TRACK_NEAR and c < b + TRACK_PAD and a < d + TRACK_PAD and o["to"] != e["to"]
                       for k, p, c, d, o in self.tracks)

    def lay(self, pts, e):
        """引いた線の線分を、通り道の記録に足す"""
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            if abs(x0 - x1) < 0.5 and abs(y0 - y1) > 0.5:
                self.tracks.append(("v", x0, min(y0, y1), max(y0, y1), e))
            elif abs(y0 - y1) < 0.5 and abs(x0 - x1) > 0.5:
                self.tracks.append(("h", y0, min(x0, x1), max(x0, x1), e))

    def gap_y(self, r, x0, x1, e):
        """行 r の下のすき間を横に走る y。使うたびに少しずらし、行き先の違う別の線が同じ所を走っていたら、空いている段へ移る"""
        k = self.gap_use.get(r, 0)
        self.gap_use[r] = k + 1
        base = self.row_y[r] + self.row_h[r] + 10
        ys = [base + self.pitch * ((k + i) % 4) for i in range(4)] + [base + 4 * self.pitch]
        return next((y for y in ys if self.clear("h", y, x0, x1, e)), ys[0])

    def chan_x(self, lane, y0, y1, e):
        """レーンの左端の通り道を縦に走る x。同じ行き先の線が通っている所があれば、そこへまとめる。
        無ければ、使うたびに少しずらし、行き先の違う別の線が同じ所を走っていたら、空いている所へ移る"""
        k = self.chan_use.get(lane, 0)
        self.chan_use[lane] = k + 1
        xs = [self.lx(lane) + 18 + 9 * ((k + i) % 6) for i in range(6)]
        same = [p for kind, p, _c, _d, o in self.tracks if kind == "v" and o["to"] == e["to"] and p in xs]
        return next((x for x in same + xs if self.clear("v", x, y0, y1, e)), xs[0])

    def top_x(self, t, y, e, from_x):
        """行き先に上から入る x。真ん中の線を、行き先の違う別の線（真上のノードから下へ出る線）が使っていたら、線の来る側へずらす"""
        half, sign = (t["right"] - t["left"]) / 2, 1 if from_x >= t["cx"] else -1
        for d in (0, sign * TOP_DX, sign * 2 * TOP_DX, -sign * TOP_DX, -sign * 2 * TOP_DX):
            if abs(d) < half - 8 and self.clear("v", t["cx"] + d, y, t["top"], e):
                return t["cx"] + d
        return None

    def entry_point(self, t, entry):
        """入る辺 → (entryX, entryY, 外形に合わせるか)・入る点。上から真ん中をずらして入るときは ("top", x)"""
        if isinstance(entry, tuple):
            w, h, dx = t["right"] - t["left"], t["bottom"] - t["top"], entry[1] - t["cx"]
            dy = abs(dx) * h / w if t["type"] == "decision" else 0   # ひし形は、上の辺が斜め
            return (round(0.5 + dx / w, 4), round(dy / h, 4), 0), (entry[1], t["top"] + dy)
        if entry == "left":                                          # 左から入る高さは、同じノードから左へ出る線と重ならない所（ひし形は、辺が斜め）
            w, f = t["right"] - t["left"], t.get("lfrac", 0.5)
            dx = abs(f - 0.5) * w if t["type"] == "decision" else 0
            return (round(dx / w, 4), f, 0 if dx else 1), (t["left"] + dx, t.get("ly", t["cy"]))
        return (*ENTRY[entry], 1), {"top": (t["cx"], t["top"]), "right": (t["right"], t["cy"])}[entry]

    def exit_sides(self):
        """線ごとに、出る辺（下・左・右）と、その辺の何本目か（0 から）を決める。
        下へ進む線のうち1本（同じレーンでいちばん近い先）を下から、残りを左右から出す。同じ辺の2本目からは、出口をずらして描く"""
        E, row, outs, side, slot = self.E, self.row, {}, {}, {}
        for k, e in enumerate(E):
            outs.setdefault(e["from"], []).append(k)
        for u, ks in outs.items():
            s = self.rect(u)
            fwd = [k for k in ks if row[E[k]["to"]] > s["row"]]
            maink = None
            if fwd:
                same = [k for k in fwd if self.rect(E[k]["to"])["lane"] == s["lane"]]
                maink = min(same, key=lambda k: row[E[k]["to"]]) if same else fwd[0]
                side[maink] = "bottom"
            used = collections.Counter()
            for k in ks:
                if k == maink:
                    continue
                t = self.rect(E[k]["to"])
                want = "right" if (row[E[k]["to"]] > s["row"] and t["lane"] > s["lane"]) else "left"
                across = self.col and row[E[k]["to"]] > s["row"] and t["lane"] != s["lane"]      # 列を分けた図: となりの列へ進む線は、その側から出す
                if used[want] > used["left" if want == "right" else "right"] and not across:      # 左右の少ないほうへ
                    want = "left" if want == "right" else "right"
                side[k], slot[k] = want, used[want]
                used[want] += 1
            if self.col:                                             # 同じ側の列へ何本も出すときは、近い列への線を下から出す（遠い列への線と、交わらない）
                for sd in ("left", "right"):
                    ks2 = [k for k in ks if side.get(k) == sd and row[E[k]["to"]] > s["row"] and self.rect(E[k]["to"])["lane"] != s["lane"]]
                    if len(ks2) > 1 and len(ks2) == used[sd] <= len(SLOT_Y):
                        ys = sorted(SLOT_Y[:len(ks2)], reverse=True)
                        for k, y in zip(sorted(ks2, key=lambda k: abs(self.rect(E[k]["to"])["lane"] - s["lane"])), ys):
                            slot[k] = SLOT_Y.index(y)
        return side, slot

    def route(self, s, t, sd, occupied, slot=0, e=None):
        """線の折れ点と、入る辺（上・左・右。上から真ん中をずらして入るときは ("top", x)）。
        s・t は両端のノードの位置（rect）、sd は出る辺、slot はその辺の何本目か、e はその線（通り道が、ほかの線と重ならないように選ぶ）"""
        stub = self.lx(s["lane"]) + (self.L["node_dx"] - 9 - SLOT_DX * slot if sd == "left"
                                     else max(self.L["node_dx"] + self.S["size"]["process"][0] + 8, s.get("cap_out", 0)) + SLOT_DX * slot)
        start = [] if sd == "bottom" else [(stub, s["cy"])]          # 左右から出るときは、まずノードの横へ出る
        down_x = s["cx"] if sd == "bottom" else stub
        e, ly = e or {"from": None, "to": None}, t.get("ly", t["cy"])
        if t["row"] > s["row"]:                                      # 下へ進む線
            free = not any((t["lane"], r) in occupied for r in range(s["row"] + 1, t["row"]))   # 行き先のレーンの途中にノードが無い
            if sd == "bottom" and s["lane"] == t["lane"] and free and abs(s["cx"] - t["cx"]) < 1:
                return [], "top"                                     # まっすぐ下
            floor = self.row_y[s["row"]] + self.row_h[s["row"]]
            if self.col and s["lane"] != t["lane"]:                  # 列を分けた図
                lo, hi = sorted((s["lane"], t["lane"]))
                if sd != "bottom" and free and not any((v, s["row"]) in occupied for v in range(lo, hi + 1) if v != s["lane"]):
                    edge_x = s["left"] if sd == "left" else s["right"]   # 判断の横から出て、となりの列の箱へ、上から入る
                    tx = self.top_x(t, s["cy"], e, edge_x)
                    if tx is not None and self.clear("h", s["cy"], edge_x, tx, e):
                        return [(tx, s["cy"])], "top" if tx == t["cx"] else ("top", tx)
                if sd == "bottom" and not free and not any((s["lane"], r) in occupied for r in range(s["row"] + 1, t["row"])):
                    gy = self.gap_y(t["row"] - 1, s["cx"], t["cx"], e)   # 自分の列をまっすぐ下りて、合流するノードの上で1本になる
                    tx = self.top_x(t, gy, e, s["cx"])
                    if tx is not None and self.clear("v", s["cx"], s["bottom"], gy, e):
                        return [(s["cx"], gy), (tx, gy)], "top" if tx == t["cx"] else ("top", tx)
            if free:
                tx = self.top_x(t, floor, e, down_x)
                if tx is not None:
                    gy = self.gap_y(s["row"], down_x, tx, e)
                    return start + [(down_x, gy), (tx, gy)], "top" if tx == t["cx"] else ("top", tx)
                side = "right" if down_x >= t["cx"] else "left"       # 真ん中も、ずらした所も空いていない（丸い終了など）: 横から入る
                sx = t["right"] + 14 if side == "right" else t["left"] - 14
                gy = self.gap_y(s["row"], down_x, sx, e)
                return start + [(down_x, gy), (sx, gy), (sx, t["cy"])], side
            if sd == "left" and s["lane"] == t["lane"]:
                cx = self.chan_x(t["lane"], s["cy"], ly, e)
                return [(cx, s["cy"]), (cx, ly)], "left"
            cx = self.chan_x(t["lane"], floor, ly, e)
            gy = self.gap_y(s["row"], down_x, cx, e)
            return start + [(down_x, gy), (cx, gy), (cx, ly)], "left"
        if s["lane"] == t["lane"] and sd == "left":                  # 戻りの線・同じ行の線は、行き先の左から入る
            cx_t = self.chan_x(t["lane"], s["cy"], ly, e)
            return [(cx_t, s["cy"]), (cx_t, ly)], "left"
        if sd == "left":
            floor = self.row_y[t["row"]] + self.row_h[t["row"]]
            cx_s, cx_t = self.chan_x(s["lane"], s["cy"], floor, e), self.chan_x(t["lane"], ly, floor + self.gap, e)
            gy = self.gap_y(t["row"], cx_s, cx_t, e)
            return [(cx_s, s["cy"]), (cx_s, gy), (cx_t, gy), (cx_t, ly)], "left"
        floor = self.row_y[s["row"]] + self.row_h[s["row"]]
        cx_t = self.chan_x(t["lane"], ly, floor + self.gap, e)
        gy = self.gap_y(s["row"], stub, cx_t, e)
        return [(stub, s["cy"]), (stub, gy), (cx_t, gy), (cx_t, ly)], "left"

    def draw_edges(self):
        S, ext, nodes = self.S, self.ext, self.nodes
        occupied = {(self.vcol(n["id"]), self.row[n["id"]]) for n in self.main}
        side, slot = self.side, self.slot
        out_left = collections.defaultdict(set)                      # ノード → 左へ出る線の、出口の高さ（辺の高さに対する比）
        for k, e in enumerate(self.E):
            if side[k] == "left":
                out_left[e["from"]].add(SLOT_Y[slot.get(k, 0) % len(SLOT_Y)])
        in_y = lambda u: next((y for y in SLOT_Y if y not in out_left[u]), 0.5)   # 左から入る線は、左へ出る線と同じ高さを使わない（同じ所を逆向きに走って見える）
        for k, e in enumerate(self.E):                               # ノードの真下へ出る線は、下のすき間の真ん中を先に取る（あとから上へ入る別の線が、そこを避ける）
            if side[k] == "bottom":
                s = self.rect(e["from"])
                self.tracks.append(("v", s["cx"], s["bottom"], self.row_y[s["row"]] + self.row_h[s["row"]] + self.gap, e))
        label_extra = "".join(p + ";" for p in S["edge"]["flow_label"].split(";") if p and p not in S["edge"]["flow"].split(";"))
        for k, e in enumerate(self.E):
            s, t, sd, n = self.rect(e["from"]), self.rect(e["to"]), side[k], slot.get(k, 0)
            t.update(lfrac=in_y(e["to"]), ly=t["top"] + in_y(e["to"]) * (t["bottom"] - t["top"]))
            ex = EXIT[sd] if sd == "bottom" else (EXIT[sd][0], SLOT_Y[n % len(SLOT_Y)])
            if sd != "bottom":                                       # 左右から出る線は、出口の高さで横へ出る
                s = dict(s, cy=s["top"] + ex[1] * (s["bottom"] - s["top"]))
            pts, entry = self.route(s, t, sd, occupied, n, e)
            en, p1 = self.entry_point(t, entry)
            kind = e.get("kind") or ("system" if "system" in (nodes[e["from"]]["type"], nodes[e["to"]]["type"]) else "flow")
            style = S["edge"]["system"] if kind == "system" else S["edge"]["flow"]
            text = self.edge_text(e)
            if text:
                style = add_style(add_style(add_style(style, label_extra), ext["edge_label_bg"]), LABEL_ALIGN[sd])
            if e.get("inferred"):
                style = add_style(style, ext["inferred"])
            style = add_style(style, f"exitX={ex[0]};exitY={ex[1]};exitDx=0;exitDy=0;entryX={en[0]};entryY={en[1]};entryDx=0;entryDy=0;" + ("" if en[2] else "entryPerimeter=0;"))
            lab = None
            if text:
                lab = (-1, 0, 13) if sd == "bottom" else (-1, -5 if sd == "left" else 5, -9)
            p0 = {"bottom": (s["cx"], s["bottom"]), "left": (s["left"], s["cy"]), "right": (s["right"], s["cy"])}[sd]
            self.lay([p0] + pts + [p1], e)
            if text and sd != "bottom":                              # 横へ伸びるラベルが図の端を越えると、書き出した画像が配置より大きくなる（照合シートの赤枠がずれる）
                w = units(text) * LABEL_FONT + 6
                self.page.grow(p0[0] - 5 - w if sd == "left" else p0[0] + 5, p0[1] - 18, w, 18)
            eid = self.page.edge(f"e-{e['from']}-{e['to']}", esc(text), style, e["from"], e["to"], points=pts, label=lab)
            self.page.edges.append({"id": eid, "from": e["from"], "to": e["to"],
                                    "points": [[round(x, 1), round(y, 1)] for x, y in [p0] + pts + [p1]]})


def flow_page(fl, ctx):
    return FlowDrawing(fl, ctx).build()


class OverviewDrawing:
    """全体構成のページ: 業務分類の階層（各フローへのリンク・結び付く資料・業務どうしの「続く」）と、資料の一覧・役割・資料間の関係"""
    BW, BH, DX, DY = 250, 64, 290, 80                                # 分類の箱の幅・高さ、階層1段の字下げ、行の送り
    DOC_DY = 84                                                      # 資料の箱の行の送り

    def __init__(self, ctx, flows, here):
        self.ctx, self.here = ctx, here
        self.S, self.hier = ctx["styles"], ctx["hierarchy"]
        self.L, self.ext = self.S["lane"], self.S["extensions"]
        self.page = Page("overview", "全体構成", self.S["model"])
        self.by_flow = {f["flow_id"]: f for f in flows}
        self.cats = ctx["taxonomy"].get("categories") or [
            {"id": f["flow_id"], "name": f["name"], "number": f.get("number", ""), "parent": f.get("parent"), "basis": f.get("basis", "llm")}
            for f in flows]
        ids = {c["id"] for c in self.cats}
        self.kids = {}                                               # 親の分類 → 子の分類
        for c in self.cats:
            self.kids.setdefault(c.get("parent") if c.get("parent") in ids else None, []).append(c)
        self.qcount = collections.Counter(q["flow"] for q in ctx["questions"] if q.get("flow"))
        self.x0, self.y0 = self.L["x0"], self.L["header_y"]
        self.pos, self.y = {}, self.y0 + 14                          # 分類 → 箱の位置／次の箱を置く y

    def build(self):
        page, S, x0, y0 = self.page, self.S, self.x0, self.y0
        page.vertex("title", "全体構成（業務分類の階層と資料）", add_style(S["text_style"], self.ext["title"]), x0, y0 - 52, 700, 30)
        page.vertex("sub-a", "■ 業務分類とフロー（青い箱をクリックするとフローのページへ。箱の下段は結び付く資料、右の線は業務どうしの「続く」）",
                    add_style(S["text_style"], "fontSize=12;fontStyle=1;"), x0, y0 - 18, 900, 20)
        for c in self.kids.get(None, []):
            self.category(c, 0)
        self.sequence()
        right = max([x for x, _ in self.pos.values()] + [x0]) + self.BW
        info_box(page, S, right + 60, y0 + 14, 450, self.summary_lines())
        y1 = max(self.y, y0 + 14) + 50
        page.vertex("sub-b", "■ 資料の一覧・役割と資料間の関係（列＝フォルダ）", add_style(S["text_style"], "fontSize=12;fontStyle=1;"), x0, y1 - 26, 600, 20)
        self.relations(self.documents(y1), y1)
        anchors(page, self.ext)
        return page

    def category(self, c, depth):
        """分類の箱を置き、その下に子の分類を字下げして置く"""
        S, page, BW, BH = self.S, self.page, self.BW, self.BH
        x, y = self.x0 + depth * self.DX, self.y
        self.y += self.DY
        fl = self.by_flow.get(c["id"])
        head = f"{c.get('number', '')} {c['name']}".strip()
        basis = "目次準拠" if c.get("basis") == "toc" else "LLM 分類"
        attrs = {"category": c["id"]}
        if fl:
            steps = sum(1 for n in fl["nodes"] if n["type"] in STEP_TYPES)
            sub = f"{c['id']}｜{steps} ステップ｜確認事項 {self.qcount.get(c['id'], 0)} 件｜{basis}"
            style = S["palettes"][0]["process_style"]
            if c["id"] in self.here:
                attrs["link"] = f"data:page/id,flow-{c['id']}"
            else:
                sub += "｜別ファイル"
        else:
            sub = f"{c['id']}｜分類（フローなし）｜{basis}"
            style = add_style(S["palettes"][1 % len(S["palettes"])]["header_style"], "fontSize=11;")
        if c.get("toc_ref"):
            attrs["tooltip"] = f"目次: {src_label(self.ctx['manifest'], c['toc_ref'])}「{c['toc_ref'].get('text', '')}」"
        tied = "・".join(labels().title(d) + ("（詳細）" if det else "") for d, _role, det in tied_docs(c, self.hier))
        page.vertex(f"cat-{c['id']}", f"{esc(head)}<br>{small(sub, S)}" + (f"<br>{small('資料: ' + tied, S)}" if tied else ""),
                    style, x, y, BW, BH, attrs=attrs)
        self.pos[c["id"]] = (x, y)
        for ch in self.kids.get(c["id"], []):
            self.category(ch, depth + 1)
            page.edge(f"tree-{c['id']}-{ch['id']}", "",
                      add_style(S["edge"]["flow"], "exitX=0.08;exitY=1;exitDx=0;exitDy=0;entryX=0;entryY=0.5;entryDx=0;entryDy=0;endArrow=none;strokeWidth=1;"),
                      f"cat-{c['id']}", f"cat-{ch['id']}", points=[(x + BW * 0.08, self.pos[ch["id"]][1] + BH / 2)])

    def sequence(self):
        """業務どうしの「続く」（資料をまたぐ受け渡しも、ここに出る）。箱の右を通す"""
        pos, BW, BH = self.pos, self.BW, self.BH
        of_h = {}                                                    # 業務の木の id → 分類
        for c in self.cats:
            if c.get("hierarchy_id") and c["id"] in pos:
                of_h.setdefault(c["hierarchy_id"], []).append(c["id"])
        style = add_style(self.S["edge"]["flow"], "exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=1;entryY=0.5;entryDx=0;entryDy=0;"
                                                  "fontSize=10;" + self.ext["edge_label_bg"])
        k = 0
        for q in self.hier.get("sequence") or []:
            for a in of_h.get(q.get("from"), []):
                for b in of_h.get(q.get("to"), []):
                    xr = max(pos[a][0], pos[b][0]) + BW + 16 + 8 * (k % 4)
                    self.page.edge(f"seq-{k}", "続く", style, f"cat-{a}", f"cat-{b}", points=[(xr, pos[a][1] + BH / 2), (xr, pos[b][1] + BH / 2)])
                    k += 1

    def summary_lines(self):
        """階層の判定と、分類の根拠"""
        hier, tax, qs = self.hier, self.ctx["taxonomy"], self.ctx["questions"]
        lines = ["<b>階層構造の判定</b>",
                 HIERARCHY_LABEL.get(hier.get("is_hierarchical"), "未判定")
                 + (f"（根拠: {'・'.join(BASIS.get(b, b) for b in hier.get('basis') or [])}）" if hier.get("basis") else "")]
        if hier.get("rationale"):
            lines.append(esc(hier["rationale"]))
        lines.append("<b>業務分類の決め方</b>")
        lines.append({"toc": "目次の業務フロー一覧に準拠", "llm": "目次に一覧が無いため LLM が分類（確認事項を参照）",
                      "mixed": "目次に準拠し、目次に無い部分は LLM が分類"}.get(tax.get("basis"), "未設定"))
        if tax.get("rationale"):
            lines.append(esc(tax["rationale"]))
        lines.append(f"<b>確認事項 {len(qs)} 件</b>（高 {sum(q['severity'] == 'high' for q in qs)}・"
                     f"中 {sum(q['severity'] == 'mid' for q in qs)}・低 {sum(q['severity'] == 'low' for q in qs)}）")
        lines.append(f"作成: {self.ctx['today']}（asis-flow-skill）")
        return lines

    def doc_top(self, y1, row):
        return y1 + self.L["header_h"] + 34 + row * self.DOC_DY

    def documents(self, y1):
        """資料の一覧（フォルダ＝列）。{資料: (列, 行)} を返す"""
        S, L, page, LW = self.S, self.L, self.page, self.L["width"]
        roles = {d.get("doc_id"): d.get("role") for d in self.hier.get("docs") or []}
        folders = {}
        for d in self.ctx["manifest"]["docs"]:
            folders.setdefault("/".join(d["folder"]) or "（直下）", []).append(d)
        cell_of = {}
        tall = max(len(v) for v in folders.values()) if folders else 1
        for i, (fname, ds) in enumerate(folders.items()):
            p = S["palettes"][i % len(S["palettes"])]
            fx = self.x0 + i * LW
            page.vertex(f"folder-{i}-h", esc(fname), p["header_style"], fx, y1, LW, L["header_h"])
            page.vertex(f"folder-{i}-c", "", p["container_style"], fx, y1 + L["header_h"], LW, 30 + tall * self.DOC_DY)
            for j, d in enumerate(ds):
                note = "同一: " + labels().title(d["duplicate_of"]) if d.get("duplicate_of") else (f"{d.get('pages') or 0} ページ" if d.get("page_list") else "読めていない")
                if d.get("has_macro"):
                    note += "・マクロあり"
                if roles.get(d["doc_id"]):
                    note = ROLE.get(roles[d["doc_id"]], roles[d["doc_id"]]) + "・" + note
                page.vertex(f"doc-{d['doc_id']}", f"{esc(labels().title(d['doc_id']))}<br>{small(d['name'] + '｜' + note, S)}",
                            add_style(S["node"]["file"], "align=left;spacingLeft=6;"), fx + 15, self.doc_top(y1, j), LW - 70, 50,
                            attrs={"doc_id": d["doc_id"], "tooltip": d["name"] + (("\n" + d["render_note"]) if d.get("render_note") else "")})
                cell_of[d["doc_id"]] = (i, j)
        return cell_of

    def relations(self, cell_of, y1):
        """資料間の関係の線。隣どうしはまっすぐ、離れた資料は列の右の通り道と行の上のすき間を通す（ほかの資料の箱を横切らない）"""
        S, LW, x0 = self.S, self.L["width"], self.x0
        for k, r in enumerate(self.hier.get("doc_relations") or []):
            if r.get("from") not in cell_of or r.get("to") not in cell_of:
                continue
            style = S["edge"]["flow"] if r.get("type") in ("parent_of", "appendix_of", "basis_of") else S["edge"]["ref"]
            name = esc(REL.get(r.get("type"), r.get("type") or ""))
            (ca, ra), (cb, rb) = cell_of[r["from"]], cell_of[r["to"]]
            if (ca == cb and abs(ra - rb) == 1) or (ra == rb and abs(ca - cb) == 1):
                self.page.edge(f"rel-{k}", name, add_style(style, "fontSize=10;" + self.ext["edge_label_bg"]), f"doc-{r['from']}", f"doc-{r['to']}")
                continue
            xa = x0 + ca * LW + LW - 34 - 7 * (k % 3)
            yg = self.doc_top(y1, rb) - 10 - 6 * (k % 3)
            self.page.edge(f"rel-{k}", name,
                           add_style(style, "fontSize=10;exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=0.5;entryY=0;entryDx=0;entryDy=0;" + self.ext["edge_label_bg"]),
                           f"doc-{r['from']}", f"doc-{r['to']}",
                           points=[(xa, self.doc_top(y1, ra) + 25), (xa, yg), (x0 + cb * LW + 15 + (LW - 70) / 2, yg)])


def overview_page(ctx, flows, here):
    return OverviewDrawing(ctx, flows, here).build()


def render(name, pages):
    """ページの並び → (draw.io の文字, 配置)。書き出し（write）と、点検の作り直し（check_flow.py）が使う"""
    mx = ET.Element("mxfile", {"host": "app.diagrams.net", "agent": "asis-flow-skill", "version": "24.0.0"})
    for p in pages:
        mx.append(copy.deepcopy(p.diagram))
    ET.indent(mx, space=" ")
    return ET.tostring(mx, encoding="unicode") + "\n", {"file": name, "pages": [
        {"index": i + 1, "id": p.id, "name": p.name, "bounds": [round(v, 1) for v in p.b], "nodes": p.nodes, "edges": p.edges, "captions": p.captions}
        for i, p in enumerate(pages)]}


def write(path, pages):
    text, layout = render(path.name, pages)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    ET.parse(path)                                                   # 書いたものが XML として読めることを確かめる
    save(path.with_suffix(".layout.json"), layout)


def figures(work, styles, split=False, write_questions=True):
    """書き出す図: [(ファイル名の後ろ（"" か "_<最上位の分類>"）, ページの並び)]。write_questions=False は点検用（questions.json を書かない）"""
    flows = load_flows(work)
    ctx = {"styles": styles, "manifest": manifest(work), "taxonomy": load(work / "taxonomy.json", {}) or {},
           "hierarchy": load(work / "structure" / "hierarchy.json", {}) or {}, "questions": collect_questions(work, write_questions),
           "today": datetime.date.today().isoformat(), "page_ids": {f["flow_id"] for f in flows}}
    groups = [("", flows)]
    if split:
        cats = {c["id"]: c for c in ctx["taxonomy"].get("categories", [])}

        def top(i):
            seen = set()
            while cats.get(i, {}).get("parent") in cats and i not in seen:
                seen.add(i)
                i = cats[i]["parent"]
            return i
        g = {}
        for f in flows:
            g.setdefault(top(f["flow_id"]), []).append(f)
        groups = [(f"_{k}", v) for k, v in g.items()]
    out = []
    for suffix, fs in groups:
        here = {f["flow_id"] for f in fs}
        ctx["page_ids"] = here
        out.append((suffix, [overview_page(ctx, flows, here)] + [flow_page(f, ctx) for f in fs], len(fs)))
    return out, ctx


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("-o", "--out")
    ap.add_argument("--split", action="store_true", help="最上位の業務分類ごとにファイルを分ける")
    ap.add_argument("--styles", default=str(ASSETS / "template_styles.json"))
    a = ap.parse_args()
    work = work_dir(a.work)
    styles = load(a.styles)
    if styles is None:
        sys.exit("template_styles.json が無い（先に harvest_template.py を実行する）")
    blocking(work)
    if not load_flows(work):
        sys.exit("flows/*.flow.json が無い")
    out = (work / pathlib.Path(a.out).expanduser()).resolve() if a.out else work / "out" / "業務フロー_as-is.drawio"   # 相対パスは作業フォルダから
    mine = out.parent == (work / "out").resolve()                    # 片付けるのは、作業フォルダの out/ の中だけ
    for old in (list(out.parent.glob(f"{out.stem}_*.drawio")) + ([] if not a.split else [out])) if mine else []:
        if old.exists():                                             # 分け方を変えたときの、前の回の図（残すと、古い図まで点検・照合される）
            old.unlink()
            old.with_suffix(".layout.json").unlink(missing_ok=True)

    groups, ctx = figures(work, styles, a.split)
    for suffix, pages, n in groups:
        path = out.with_name(out.stem + suffix + out.suffix)
        write(path, pages)
        print(f"{path}  （{len(pages)} ページ: 全体構成 ＋ フロー {n}）")
    print(f"確認事項 {len(ctx['questions'])} 件 → {work / 'questions.json'}")


if __name__ == "__main__":
    main()
