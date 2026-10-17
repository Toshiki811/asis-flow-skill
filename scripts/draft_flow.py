#!/usr/bin/env python3
"""draft_flow.py — フローの下書きを、カードから機械で作って見せる。LLM は、下書きを見て「直しの一覧」だけを書く。

  python3 scripts/draft_flow.py -w <作業フォルダ> --flow F03       # 下書きを表示する（ファイルには書かない）
  python3 scripts/apply_edits.py -w <作業フォルダ> --flow F03      # 下書き＋直しの一覧（flows/F03.edits.json）→ flows/F03.flow.json

同じカードからは、いつも同じ下書きになる。決め方:
  項目      分類に割り当てたページのカードの手順・分岐を、1項目＝1ノードにする。並びは、ページの順・カードの順
            （分岐は、同じ行か、あとの行で始まる最初の手順の前。行の番号が無ければ、行き先の文に当たる手順の前）
  キー      項目を指す名前は「資料.ページ.項目」（M01.35.s3）。直しの一覧は、このキーで指す
  ラベル    カードの action（分岐は question）そのまま。30 字を超え、かっこ書きがあるときは、かっこの前までをラベルに、かっこの中を補足にする
  レーン    カードの actor を、用語の正本の名前に寄せたもの。actor が空なら、まとめた概要の手順の actor。それも無ければ空（直しの一覧の default_lane に置く）
  帳票      inputs は参照する（in）、outputs は作る（out）
  動作の案  手順の名前の終わりの動詞から、動作の種類の案（決まり 3.22 の読み方と、12 語の意味。ACTION_VERBS）。使うシステムは、点検の札と同じ決め方。
            案を見せるだけで、選ぶのは LLM（点検は、案と比べない）
  まとめる  概要のページ（flow_diagram・overview）の手順・分岐は、あとのページの、同じ番号の見出しの下の最初の手順・分岐にまとめる
            （見出しの題が両方にあれば、似ていること。番号が無ければ、文がほぼ同じ1件）。候補として見せ、直しの一覧の unmerge で外せる
  線        並びの順につなぐ。推定にするのは、番号の無い箇条書きの項から次の項へ進むところ・資料をまたぐところ・補足の中の作業の前後
  分岐      カードの分岐の行き先（to）に当たる手順へ線を引く。書かれていない側は、その手順を飛ばして次へ進む線を、推定で置く
"""
import argparse
import json
import pathlib
import re
import unicodedata

from common import (CardItems, label_core, labels, load, load_card, load_page, manifest, node_systems, norm, page_json, page_units, same_text, system_names, systems_in,
                    unit_marks, when_done, work_dir)

OVERVIEW = ("flow_diagram", "overview")                              # 概要のページ（あとのページが、同じ作業を詳しく書く）
APPROVAL = re.compile(r"^(?:事前に|あらかじめ|必ず)?([^、。をにがは]{1,20}?)の(?:承認|決裁|許可)を(?:得|受け)る$")   # 「X の承認を得る」だけの手順: 担当は X（決まり 3.21）
AUTOMATIC = re.compile(r"自動(?:的)?に|自動で")                                  # 「自動で」行われる手順: そのシステムのレーン（決まり 3.8）
SMALL_HINT = re.compile(r"ボタンを押す|クリックする|実行する|起動する")   # 量「小」の案（決まり 3.22。付けるかどうかと理由は LLM）
HUGE_HINT = re.compile(r"年\s*[1１一]\s*回|年次|決算|全員|全件|全職員|棚卸|予算の編成|予算を編成|半期|四半期|期間をかけ|数日|数週間")   # 量「特大」の手がかりの語
ACTION_VERBS = (                                                     # 動作の種類 → 箱の名前の終わりの動詞（決まり 3.22 と、スキーマの 12 語の意味。下書きに案として出すだけ）
    ("入力する", ("入力する", "登録する", "記入する", "記帳する")),
    ("作成する", ("作成する", "作る")),
    ("確認する", ("確認する", "照合する", "点検する", "チェックする", "確かめる")),
    ("承認する", ("承認する", "決裁する", "押印する")),
    ("提出・回付する", ("提出する", "回付する", "送付する", "差し戻す", "回す", "送る")),
    ("受け取る", ("受け取る", "受領する", "回収する")),
    ("出力する", ("出力する", "印刷する")),
    ("連絡する", ("連絡する", "依頼する", "報告する", "通知する", "知らせる")),
    ("計算する", ("計算する", "集計する")),
    ("調べる", ("調べる", "調査する", "検討する", "選定する")),
    ("保管する", ("保管する", "整理する", "保存する", "廃棄する")),
)
JOINS = {"確かめる": ("確かめ、", "確かめて"), "差し戻す": ("差し戻し、", "差し戻して"), "回す": ("回し、", "回して"), "送る": ("送り、", "送って"),   # 「する」で終わらない動詞の、次の動きへつなぐ形
         "受け取る": ("受け取り、", "受け取って"), "知らせる": ("知らせ、", "知らせて"), "調べる": ("調べ、", "調べて"), "作る": ("作り、", "作って")}
ASK_OTHERS = re.compile(r"(電話|メール|ＦＡＸ|FAX|ファクス)で、?([^たるく、]*)$|問い?合わ?せ(?:て|、)$")   # 相手に確かめる手段のあとの「確かめる」: 連絡する（決まり 3.22。
                                                                     # 間に、別の動き〔〜た・〜る・〜く・「、」〕をはさむ名前は除く: 「メールで届いた請求書を確認する」「電話で依頼を受け、内容を確認する」）
SHORT_VERBS = ("回す", "送る", "作る")                                # 2 字の動詞は、助詞のすぐあとだけ（「見送る」「見回す」「形作る」は別の動き）
PARTICLES = "をにへでと、"
TIMING = re.compile(r"タイミング|時期|頻度|毎月|毎週|毎日|毎年|月末|月初|年度末|期末|随時|都度|受けたら|届いたら")   # 開始条件の手がかりになる補足の語
HEAD = re.compile(r"^(?:[（(]\s*([0-9０-９]{1,2})\s*[)）]|([①-⑳])|([0-9０-９]{1,2})(?=[.．、\s]))[.．、]?\s*(\S.*)$")
PAGE_TYPE_JA = {"reference": "参考", "table": "表", "form": "様式", "toc": "目次", "cover": "表紙", "blank": "白紙", "code": "マクロ",
                "procedure": "手順", "flow_diagram": "フロー図", "overview": "概要"}


def after_particle(before):
    """2 字の動詞の前が、助詞か（「から」「より」も。前に何も無いときも）"""
    return not before or before[-1] in PARTICLES or before.endswith(("から", "より"))


def verb_at(t, v):
    """t が動詞 v で終わるか（2 字の動詞は、助詞のすぐあとか、名前がその動詞だけのとき）"""
    return t.endswith(v) and (v not in SHORT_VERBS or after_particle(t[:-len(v)]))


def read_verb(v, a, before, system):
    """動詞 v（種類 a）を、使うシステム・前の文で読み替える（決まり 3.22）: (種類, 理由)"""
    if a == "作成する" and system:
        return "入力する", f"「{v}」で、使うシステム「{system}」がある（画面で作るなら入力する。紙・ファイルを一から作るなら作成する）"
    if v == "実行する":
        return "入力する", f"「{v}」で、使うシステム「{system}」がある（システムで実行する）"
    m = ASK_OTHERS.search(before) if a == "確認する" else None
    if m:
        return "連絡する", f"「{m.group(1) or '問い合わせ'}」で「{v}」（相手に確かめる）"
    return a, f"「{v}」"


def action_hint(label, system=""):
    """箱の名前の動詞から、動作の種類の案: (語の並び, 理由)。名前（終わりのかっこ書きと句点を除く）の終わりが ACTION_VERBS の動詞でなければ ([], "")。
    使うシステム（点検の札と同じ決め方。下書きが渡す）があれば、作成する → 入力する・実行する → 入力する。電話・メールで確かめる → 連絡する。
    前に、別の種類の動きが「〜し、」「〜して」でつながっていれば、2 つの動きの案も添える（前の動きも、同じ決まりで読み替える）"""
    t = re.sub(r"[。．.\s]+$", "", label_core(label))
    verbs = [(v, a) for a, vs in ACTION_VERBS for v in vs] + ([("実行する", "入力する")] if system else [])
    last = max(((v, a) for v, a in verbs if verb_at(t, v)), key=lambda x: len(x[0]), default=None)
    if last is None:
        return [], ""
    head = t[:-len(last[0])]
    a, why = read_verb(*last, head, system)
    hits = []
    for x, b in verbs:
        for j in JOINS.get(x) or (x[:-2] + "し、", x[:-2] + "して"):
            at = head.find(j)
            if at >= 0 and (x not in SHORT_VERBS or after_particle(head[:at])):
                b2 = read_verb(x, b, head[:at], system)[0]
                if b2 != a:
                    hits.append((at, j, b2))
    hits.sort()
    return ([hits[0][2], a], why + f"。前に「{hits[0][1].rstrip('、')}」もある（2 つの動きなら、2 語で書く）") if hits else ([a], why)


def key_of(doc, page, item):
    return f"{doc}.{page}.{item}"


def head_of(text):
    """見出しの頭の番号と、残りの題: 「（2） 書類の確認」→ ("2", "書類の確認")。番号で始まらなければ (None, 全体)"""
    text = str(text or "").strip()
    m = HEAD.match(text)
    if not m or re.match(r"[0-9０-９/／:：]", m.group(4)):            # 「2024/04/01」は見出しでない
        return None, text
    return unicodedata.normalize("NFKC", m.group(1) or m.group(2) or m.group(3)), m.group(4).strip()


class Terms:
    """用語の正本: 原典の呼び方 → (種類, 正本の名前)"""

    def __init__(self, work):
        self.by = {}
        for t in (load(pathlib.Path(work) / "glossary.json", {}) or {}).get("terms") or []:
            for k in [t.get("name")] + list(t.get("aliases") or []):
                if k:
                    self.by[norm(k)] = (t.get("kind"), t.get("name"))

    def name(self, text, kinds):
        x = self.by.get(norm(text))
        return x[1] if x and x[0] in kinds else ""

    def kind(self, name):
        return (self.by.get(norm(name)) or ("", ""))[0]

    def find(self, text, kinds):
        """文の中に出てくる、その種類の用語の正本の名前（長い呼び方から）。無ければ空"""
        t = norm(text)
        return next((self.by[k][1] for k in sorted(self.by, key=len, reverse=True) if k and k in t and self.by[k][0] in kinds), "")


class Draft:
    """分類1つの下書き: nodes（キー → 項目）、seq（並び）、merged（まとめた項目 → (まとめ先, 理由)）"""

    def __init__(self, work, cid):
        self.work, self.cid = pathlib.Path(work), cid
        self.m = manifest(self.work)
        tax = load(self.work / "taxonomy.json", {}) or {}
        self.cats = {c["id"]: c for c in tax.get("categories") or [] if isinstance(c, dict) and c.get("id")}
        if cid not in self.cats:
            raise SystemExit(f"taxonomy.json に分類 {cid} が無い")
        self.cat = self.cats[cid]
        self.pages = [(dp["doc_id"], p) for dp in self.cat.get("pages") or [] for p in dp.get("pages") or []]
        self.terms, self.lab = Terms(self.work), labels(self.m, self.work)
        self.items_of, self.sys_names = CardItems(self.work), system_names(self.work)   # 動作の案の「使うシステム」は、点検の札と同じ決め方（node_systems）
        self.nodes, self.seq, self.merged, self.notes, self.cards, self.actors = {}, [], {}, [], {}, []
        for doc, page in self.pages:
            self.collect(doc, page)
        self.candidates()
        self.boxes = {}                                              # 手順のキー → その直後に機械が置く、プログラムの起動点の箱
        from systems import Ledger, has_ledger
        if has_ledger(self.work):
            led = Ledger(self.work, self.m)
            for x in led.good_links():
                e = led.entry(x["entry"])
                self.boxes.setdefault(x["step"], []).append(f"{e['system']['name']}「{e['label']}」▶ {e['activity']}")

    # ---- カードから項目を集める ----

    def collect(self, doc, page):
        c = load_card(self.work, doc, page)
        self.cards[(doc, page)] = c
        if not c:
            self.notes.append(f"{doc} p{page:04d}: カードが無い")
            return
        pj = load_page(page_json(self.work, doc, page))
        units, marks = page_units(pj), unit_marks(pj)
        heads = [(k,) + head_of(u) for k, u in enumerate(units, 1) if head_of(u)[0] is not None]   # 見出しの行と、見出しが頭につながった文
        boxes = dict(head_of(str(g.get("label", "")).split("\n")[0]) for g in (c.get("diagram") or {}).get("nodes") or [])
        for a in c.get("actors") or []:
            nm = self.terms.name(a, ("actor", "system"))
            if nm and nm not in self.actors:
                self.actors.append(nm)

        def node(it, kind):
            ln = it.get("lines")
            line = ln[0] if isinstance(ln, list) and len(ln) == 2 and isinstance(ln[0], int) and 1 <= ln[0] <= len(units) else None
            num, title = head_of(it.get("section"))
            if c.get("page_type") in OVERVIEW:                       # 概要の手順の番号: 原典の番号（no）か、小見出しの番号。題は、小見出しか、同じ番号の図の箱
                num = unicodedata.normalize("NFKC", str(it["no"])) if str(it.get("no") or "").strip() else num
                title = title if head_of(it.get("section"))[0] else boxes.get(num, "")
            elif num is None and line is not None:                   # 小見出しが書かれていなければ、自分の行の頭の見出しか、紙面でいちばん近い前の見出し
                near = [h for h in heads if h[0] <= line]
                num, title = near[-1][1:] if near else (None, "")
            actor = str(it.get("actor") or "").strip()
            label = str(it.get("action" if kind == "step" else "question") or "").strip()
            lane, why = self.terms.name(actor, ("actor", "system")) if actor else "", ""
            hit = APPROVAL.match(label) if kind == "step" else None
            if hit and self.terms.name(hit.group(1), ("actor",)):     # 承認・決裁を得る手順は、承認する人のレーン
                lane, why = self.terms.name(hit.group(1), ("actor",)), "承認の相手のレーン"
            elif kind == "step" and AUTOMATIC.search(label):          # 自動で行われる手順は、そのシステムのレーン（システム処理の箱になる）
                auto = self.terms.name(str(it.get("system") or ""), ("system",)) or self.terms.find(label, ("system",))
                lane, why = (auto, "自動の処理: システムのレーン") if auto else (lane, why)
            group = max([k for k in range(1, (line or 0) + 1) if marks[k - 1]], default=0)      # この項目が入っている箇条書き・番号つきの項（その頭の行）
            return {"key": key_of(doc, page, it["id"]), "doc": doc, "page": page, "item": it["id"], "kind": kind, "overview": c.get("page_type") in OVERVIEW,
                    "label": label, "actor": actor, "lane": lane, "lane_why": why, "system": str(it.get("system") or "").strip(),
                    "condition": str(it.get("condition") or "").strip(), "inputs": list(it.get("inputs") or []), "outputs": list(it.get("outputs") or []),
                    "origin": it.get("origin", "body"), "applies_to": it.get("applies_to"), "num": num, "title": title or "",
                    "line": line, "mark": marks[group - 1] if group else "", "group": group,
                    "branches": [b for b in it.get("branches") or [] if isinstance(b, dict)],
                    "sources": [(doc, page, it["id"])], "files": []}
        rows = [node(s, "step") for s in c.get("steps") or [] if isinstance(s, dict) and s.get("id")]
        if not rows and not (c.get("decisions") or []):              # 図だけのページ: 図の箱が項目（flow-rules.md 1.7。線の順にたどる）
            rows = self.diagram_rows(c, node)
        for r in [x for x in rows if x["origin"] == "note" and x["applies_to"]]:       # 補足の中の作業は、掛かる手順のすぐあと
            at = next((i for i, x in enumerate(rows) if x["item"] == r["applies_to"]), None)
            if at is not None:
                rows.remove(r)
                rows.insert(next(i for i, x in enumerate(rows) if x["item"] == r["applies_to"]) + 1, r)
        for d in [node(d, "decision") for d in c.get("decisions") or [] if isinstance(d, dict) and d.get("id")]:
            rows.insert(self.place(rows, d), d)
        for i, r in enumerate(rows):                                 # 概要の分岐の番号は、すぐあとの手順の番号
            nxt = next((x for x in rows[i + 1:] if x["kind"] == "step"), None)
            if r["overview"] and r["kind"] == "decision" and r["num"] is None and nxt is not None:
                r["num"], r["title"] = nxt["num"], nxt["title"]
        for r in rows:
            if r["actor"] and not r["lane"]:
                self.notes.append(f"{r['key']}: 担当「{r['actor']}」が用語の正本に無い（lane で置くレーンを決める。新しい担当なら親に知らせる）")
            self.nodes[r["key"]] = r
            self.seq.append(r["key"])

    @staticmethod
    def diagram_rows(c, node):
        """手順・分岐が 1 つも無い図だけのページ: 図の箱（処理・判断・システム）を、開始から線の順にたどって項目にする。
        箱の字が作業（action）・問い（question）、レーンが担当。分岐の行き先は、線の上の字と、先の箱の字"""
        dg = c.get("diagram") or {}
        boxes = [g for g in dg.get("nodes") or [] if isinstance(g, dict) and g.get("id")]
        by_id = {g["id"]: g for g in boxes}
        outs = {}
        for e in dg.get("edges") or []:
            if isinstance(e, dict) and e.get("from") in by_id and e.get("to") in by_id:
                outs.setdefault(e["from"], []).append(e)
        order, seen = [], set()
        heads = [g["id"] for g in boxes if g.get("shape") == "start"] or [g["id"] for g in boxes if not any(e["to"] == g["id"] for es in outs.values() for e in es)]
        stack = list(reversed(heads))
        while stack:                                                 # 開始から深さ優先（分岐は、線の書かれた順）。線でつながらない箱は、書かれた順に後ろへ
            g = stack.pop()
            if g in seen:
                continue
            seen.add(g)
            order.append(g)
            stack += [e["to"] for e in reversed(outs.get(g, []))]
        order += [g["id"] for g in boxes if g["id"] not in seen]
        rows = []
        for gid in order:
            g = by_id[gid]
            shape, label = g.get("shape"), str(g.get("label") or "").split("\n")[0].strip()
            if shape not in ("process", "decision", "system") or not label:
                continue
            it = {"id": gid, "actor": g.get("lane") or "", "action": label, "question": label, "origin": "body",
                  "branches": [{"label": str(e.get("label") or ""), "to": str(by_id[e["to"]].get("label") or "").split("\n")[0]} for e in outs.get(gid, [])]
                  if shape == "decision" else []}
            rows.append(node(it, "decision" if shape == "decision" else "step"))
        return rows

    @staticmethod
    def place(rows, d):
        """分岐を置く場所（rows の何番目の前か）"""
        steps = [(i, r) for i, r in enumerate(rows) if r["kind"] == "step" and r["origin"] == "body"]
        if d["line"] is not None:
            at = next((i for i, r in steps if r["line"] is not None and r["line"] >= d["line"]), None)
            if at is not None:
                return at
        for b in d["branches"]:
            at = next((i for i, r in steps if b.get("to") and same_text(b["to"], r["label"])), None)
            if at is not None:
                return at
        return len(rows)

    # ---- 概要の手順を、詳しい手順にまとめる候補 ----

    def candidates(self):
        for a in [self.nodes[k] for k in self.seq if self.nodes[k]["overview"]]:
            later = [b for b in (self.nodes[k] for k in self.seq) if b["doc"] == a["doc"] and b["page"] > a["page"] and not b["overview"] and b["kind"] == a["kind"]]
            hit, why = None, ""
            same = [b for b in later if a["num"] and b["num"] == a["num"]]
            if same and (not a["title"] or not same[0]["title"] or same_text(a["title"], same[0]["title"], 0.5)):
                hit, why = same[0], f"番号と見出しが同じ（{a['num']} {a['title'] or same[0]['title']}）"
            elif not same:
                close = [b for b in later if same_text(a["label"], b["label"])]
                if len(close) == 1:
                    hit, why = close[0], "文がほぼ同じ"
            if hit is not None:
                self.merged[a["key"]] = (hit["key"], why)

    # ---- 見せる ----

    def ref(self, n):
        return self.lab.ref(n["doc"], n["page"]) if hasattr(self.lab, "ref") else f"{n['doc']} p{n['page']:04d}"

    def quote(self, n):
        c = self.cards[(n["doc"], n["page"])]
        it = next((x for place in ("steps", "decisions") for x in c.get(place) or [] if x.get("id") == n["item"]), None)
        if it is None:                                               # 図の箱: 引用は箱の字
            g = next((x for x in (c.get("diagram") or {}).get("nodes") or [] if x.get("id") == n["item"]), {})
            return str(g.get("label", "")), False, None
        return str(it.get("quote", "")), bool(it.get("from_image")), it.get("lines")

    def show(self):
        c, out = self.cat, []
        out.append(f"== 下書き {self.cid}「{c.get('name')}」（ページ: " + "・".join(f"{d} p{p:04d}" for d, p in self.pages) + "）")
        out.append("担当の候補（カードの actors を、用語の正本の名前にしたもの）: " + ("・".join(self.actors) or "なし"))
        seq = [self.nodes[k] for k in self.seq if k not in self.merged]
        cue = ([f"最初の作業の条件「{seq[0]['condition']}」"] if seq and seq[0]["condition"] else []) + [
            f"{d} p{p:04d} の補足「{x['text'][:50]}」" for (d, p), c in self.cards.items() for x in c.get("notes") or [] if isinstance(x, dict) and TIMING.search(str(x.get("text", "")))][:3]
        out.append("開始条件（start_condition）の手がかり: " + ("／".join(cue) or "なし") + "。原典に書かれた開始条件だけを書く（無ければ空。機械が確認事項にする）")
        if seq:
            out.append(f"終了条件（end_condition）: 原典に成果が書かれていなければ空にする。機械が、終了に入る最後の作業から入れる（いまの並びなら「{when_done(seq[-1]['label'])}」）")
        if self.merged:
            out.append("\nまとめた候補（概要の項目 → 詳しい項目。違っていれば unmerge で外す）:")
            for a, (b, why) in self.merged.items():
                out.append(f"  {a}「{self.nodes[a]['label'][:30]}」→ {b}「{self.nodes[b]['label'][:30]}」: {why}")
        out.append("\n並び（キー・種類・担当・ラベル）:")
        for i, k in enumerate([k for k in self.seq if k not in self.merged], 1):
            n = self.nodes[k]
            q, img, lines = self.quote(n)
            tags = [x for x in ("補足の中の作業" if n["origin"] == "note" else "", "箇条書き" if n["mark"] == "bullet" else "", "概要" if n["overview"] else "",
                                n.get("lane_why", "")) if x]
            lane = n["lane"] or next((self.nodes[a]["lane"] for a, (b, _w) in self.merged.items() if b == k and self.nodes[a]["lane"]), "")   # まとめた概要の手順の担当
            out.append(f"{i:3d}  {k}  {'分岐' if n['kind'] == 'decision' else '手順'}  [{lane or ('担当なし' if not n['actor'] else '？' + n['actor'])}]  {n['label']}"
                       + (f"  〈{'・'.join(tags)}〉" if tags else ""))
            said = n["label"] + " " + q
            huge = HUGE_HINT.search(said) if n["kind"] == "step" else None
            src = [{"doc_id": d, "page": p, "item": it} for kk in [k] + [a for a, (to, _w) in self.merged.items() if to == k] for d, p, it in self.nodes[kk]["sources"]]
            tagged = node_systems({"type": "process", "sources": src}, lane, self.items_of, self.sys_names)
            acts, why = action_hint(n["label"], "・".join(tagged)) if n["kind"] == "step" and self.terms.kind(lane) != "system" else ([], "")
            shown = "・".join(tagged) or n["system"]                  # 使うシステムは、点検の札と同じ決め方（カードの欄と違えば、欄も見せる）
            extra = [f"条件: {n['condition']}" if n["condition"] else "",
                     f"システム: {shown}" + (f"（カードの欄: {n['system']}）" if n["system"] and tagged and systems_in(n["system"], self.sys_names) != tagged else "") if shown else "",
                     f"動作の案: {acts[0] if len(acts) == 1 else json.dumps(acts, ensure_ascii=False)}（{why}）" if acts else "",
                     f"量の候補: 特大（「{huge.group(0)}」。時間・日の単位でかかると読めるときだけ、理由つきで）" if huge else
                     "量の案: 小（ボタンを押す・実行するだけなら、time で volume と理由）" if n["kind"] == "step" and SMALL_HINT.search(n["label"]) else "",
                     f"参照: {'・'.join(n['inputs'])}" if n["inputs"] else "", f"作る: {'・'.join(n['outputs'])}" if n["outputs"] else "",
                     "行き先: " + " ／ ".join(f"「{b.get('label', '')}」→ {b.get('to') or '（書かれていない）'}" for b in n["branches"]) if n["kind"] == "decision" else ""]
            if any(extra):
                out.append("       " + "  ".join(x for x in extra if x))
            out.append(f"       原文（{'画像' if img else '行 ' + '〜'.join(str(x) for x in sorted(set(lines))) if lines else '引用'}）: 「{q[:90]}」")
            for b in [b for kk in [k] + [a for a, (to, _w) in self.merged.items() if to == k] for b in self.boxes.get(kk, [])]:
                out.append(f"       ⇒ この作業のすぐあとに、機械がシステム処理の箱を置く: {b}（同じ処理の箱を add で足さない）")
        idle = [(d, p) for d, p in self.pages if not any(n["doc"] == d and n["page"] == p for n in self.nodes.values())]
        if idle:
            out.append("\n手順・分岐の無いページ（使わなかったページとして、理由つきで載る）: "
                       + "・".join(f"{d} p{p:04d}（{PAGE_TYPE_JA.get(self.cards[(d, p)].get('page_type'), '？')}）" for d, p in idle))
        if self.notes:
            out.append("\n注意:")
            out += ["  " + x for x in self.notes]
        return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--flow", required=True, help="業務分類の id（F03）")
    a = ap.parse_args()
    work = pathlib.Path(work_dir(a.work))
    d = Draft(work, a.flow)
    print(d.show())


if __name__ == "__main__":
    main()
