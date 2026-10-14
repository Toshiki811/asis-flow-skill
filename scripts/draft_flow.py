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
  まとめる  概要のページ（flow_diagram・overview）の手順・分岐は、あとのページの、同じ番号の見出しの下の最初の手順・分岐にまとめる
            （見出しの題が両方にあれば、似ていること。番号が無ければ、文がほぼ同じ1件）。候補として見せ、直しの一覧の unmerge で外せる
  線        並びの順につなぐ。推定にするのは、番号の無い箇条書きの項から次の項へ進むところ・資料をまたぐところ・補足の中の作業の前後
  分岐      カードの分岐の行き先（to）に当たる手順へ線を引く。書かれていない側は、その手順を飛ばして次へ進む線を、推定で置く
"""
import argparse
import pathlib
import re
import unicodedata

from common import labels, load, load_card, load_page, manifest, norm, page_json, page_units, same_text, unit_marks, when_done, work_dir

OVERVIEW = ("flow_diagram", "overview")                              # 概要のページ（あとのページが、同じ作業を詳しく書く）
APPROVAL = re.compile(r"^(?:事前に|あらかじめ|必ず)?([^、。をにがは]{1,20}?)の(?:承認|決裁|許可)を(?:得|受け)る$")   # 「X の承認を得る」だけの手順: 担当は X（決まり 3.21）
AUTOMATIC = re.compile(r"自動(?:的)?に|自動で")                                  # 「自動で」行われる手順: そのシステムのレーン（決まり 3.8）
SMALL_HINT = re.compile(r"ボタンを押す|クリックする|実行する|起動する")   # 量「小」の案（決まり 3.22。付けるかどうかと理由は LLM）
HUGE_HINT = re.compile(r"年\s*[1１一]\s*回|年次|決算|全員|全件|全職員|棚卸|予算の編成|予算を編成|半期|四半期|期間をかけ|数日|数週間")   # 量「特大」の手がかりの語
TIMING = re.compile(r"タイミング|時期|頻度|毎月|毎週|毎日|毎年|月末|月初|年度末|期末|随時|都度|受けたら|届いたら")   # 開始条件の手がかりになる補足の語
HEAD = re.compile(r"^(?:[（(]\s*([0-9０-９]{1,2})\s*[)）]|([①-⑳])|([0-9０-９]{1,2})(?=[.．、\s]))[.．、]?\s*(\S.*)$")
PAGE_TYPE_JA = {"reference": "参考", "table": "表", "form": "様式", "toc": "目次", "cover": "表紙", "blank": "白紙", "code": "マクロ",
                "procedure": "手順", "flow_diagram": "フロー図", "overview": "概要"}


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
        it = next((x for place in ("steps", "decisions") for x in self.cards[(n["doc"], n["page"])].get(place) or [] if x.get("id") == n["item"]), {})
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
            extra = [f"条件: {n['condition']}" if n["condition"] else "", f"システム: {n['system']}" if n["system"] else "",
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
