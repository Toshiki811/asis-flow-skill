#!/usr/bin/env python3
"""check_extract.py — ページカード（LLM がページ画像を読んで書いたもの）の機械チェック。フロー化の前に FAIL 0 にする。

  python3 scripts/check_extract.py -w <作業フォルダ>
  python3 scripts/check_extract.py -w <作業フォルダ> --only M01,M02 --report <作業フォルダ>/out/チェック_読取.txt
  python3 scripts/check_extract.py -w <作業フォルダ> --only M01 --pages 13-24      # 自分の束のページだけを見る

機械抽出（pages/*.json）と突き合わせて、形の崩れ・作り話・読み落とし・ページの取り違えを拾う。正典: references/flow-rules.md §1
[形]
  FAIL  カードが無いページ／JSON として壊れている／スキーマ（assets/schemas/card.schema.json）と違う（必須・型・決まった語・余分な項目）
  FAIL  項目の id（s1・d1・g1・n1）の重複／applies_to が無い手順を指す／図の線が、定義していない箱を指す
[作り話・読み落とし]
  FAIL  手順・分岐・補足の引用が機械抽出の文字に見当たらないのに「画像から読んだ」の印（from_image）が無い
  FAIL  手順のページ・フロー図のページなのに、手順も図も分岐も空
  FAIL  補足（注記・ただし書き・囲み・欄外）の中に作業らしい文があるのに、手順・分岐にも、「作業でない」の記録（notes の why_not_action）にも無い
  FAIL  行の番号（lines）が範囲の外／引用が、指した行の文字と違う（手で書いた・quotes.py fill をしていない）／画像の中の文字に行の番号
  WARN  指した行が、書いた内容と合っていない疑い／手順の並びが紙面の行の順と逆／同じ1文を、同じ担当の手順2つに分けている
  WARN  引用が機械抽出の文字と完全には一致しない／機械抽出の文字のうち、カードに反映された行が少ない
  WARN  画像・図形・線で描いた図があるのに図の記述が空／矢印の本数に対して図の線が少ない／ひし形があるのに分岐が無い
  WARN  字が小さいページを拡大（crop_page.py）せずに読んでいる／目次のページなのに目次の項目が空
[ページの取り違え]
  FAIL  引用がそのページに無く、前後のページにある（どのページにあるかを名指しする。同じ向きに何件もあれば「系統的なずれ」）
  FAIL  カードの見出しがそのページに無く、隣のページにある
  FAIL  カードの printed_page が、ページの端に刷られた数字と違う／資料の中で番号の並びが1枚だけ飛ぶ
  FAIL  ページ画像と機械抽出が別のスライド・シートのもの（misaligned）／シートとページ画像の対応が取れていない
  WARN  番号が刷られているのに printed_page が null
  INFO  種別ごとの枚数・補足由来の手順の数・画像由来の引用の数・読めなかった箇所の数・曖昧な点の数
終了コード: FAIL があれば 1。
"""
import argparse
import collections
import json

from common import (CJK, NON_CONTENT_TYPES, PAGE_TYPES, QUOTE_OK, SHIFT_MIN, SHIFT_SHARE, Report, _split, all_pages, bare_number, card_json, covers,
                    action_like, guarded, labels, load_markers, load_page, manifest, norm, note_actions, page_json, page_text, page_units,
                    parse_pages, quote_match, span_text, work_dir)
from validate import validate

FLOWISH = {"procedure", "flow_diagram"}
DECISION_SHAPES = ("decision", "diamond")
NEAR = 2                                                             # 引用を探す前後のページ数
SHAPE_ERRORS = 8                                                     # カード1枚につき、形の違反を何件まで出すか
BLANK_MAX_LINES = 4                                                  # 白紙としたページの機械抽出に、文字の行がこれ以上あれば白紙でない
QUOTED = (("手順", "steps", "action"), ("分岐", "decisions", "question"), ("補足", "notes", "text"))   # 引用を持つ項目: 呼び名・置き場・名前の項目
NO_NOTE_ACTIONS = ("code", "toc", "cover", "blank")                  # 補足の中の作業を探さない種別
STEP_PAGES = ("procedure", "flow_diagram", "overview")               # 作業らしい文が、どれかの項目に指されているかを見る種別
OMISSION_TYPES = ("procedure", "flow_diagram", "table", "form", "code")   # 機械抽出の行がカードに出てくるかを見る種別
POINT_MIN = 0.2                                                      # 行を番号で指した項目: 作業の文の2字の並びが、指した行にこの割合より少なければ、指し違いの疑い


def grams(s, n=4):
    return {s[i:i + n] for i in range(max(0, len(s) - n + 1))}


def strings(o):
    if isinstance(o, str):
        yield o
    elif isinstance(o, dict):
        for v in o.values():
            yield from strings(v)
    elif isinstance(o, list):
        for v in o:
            yield from strings(v)


class Page:
    """点検中の1ページ: カード・機械抽出・前後のページ"""

    def __init__(self, checks, doc, entry, card):
        self.doc, self.no, self.entry, self.c = doc, entry["page"], entry, card
        self.where = checks.lab.ref(doc, self.no)
        self.pj = load_page(checks.work / entry["json"])
        self.units = page_units(self.pj)                             # 行の番号（lines）が指す並び
        self.mtext = checks.mtext(doc, self.no)
        seq = checks.order.get(doc, [])
        at = seq.index(self.no) if self.no in seq else -1
        near = [seq[j] for j in range(max(0, at - NEAR), min(len(seq), at + NEAR + 1)) if j != at] if at >= 0 else []
        self.near = sorted(near, key=lambda x: abs(x - self.no))     # 前後のページ（近い順）
        self.png = f"{doc}/p{self.no:04d}.png"


class CardChecks:
    """全カードの点検。資料をまたいで見るもの（系統的なずれ・番号の並び）は、ここに貯めて最後に見る"""

    def __init__(self, work, rp):
        self.work, self.rp = work, rp
        self.m = manifest(work)
        self.lab = labels(self.m)
        self.markers = load_markers()
        self.order = {}                                              # 資料 → ページの並び（前後のページを引く）
        for d, p in all_pages(self.m):
            self.order.setdefault(d["doc_id"], []).append(p["page"])
        self._text = {}
        self.n = collections.Counter()                               # cards・note_steps・cand・cand_step・cand_note・img_quotes・unread・uncertain
        self.by_type = collections.Counter()
        self.by_lines = None
        self.missing = []
        self.shifted, self.checked = {}, collections.Counter()       # 資料 → {ずれ: ページ}／照合した引用の数
        self.printed = {}                                            # 資料 → [(page, カードの printed_page)]

    def mtext(self, doc, page):
        if (doc, page) not in self._text:
            self._text[(doc, page)] = page_text(load_page(page_json(self.work, doc, page)))
        return self._text[(doc, page)]

    def run(self, only=None, pages=None):
        seen = 0
        for d, p in all_pages(self.m):
            if (only and d["doc_id"] not in only) or (pages is not None and p["page"] not in pages):
                continue
            seen += 1
            card = self.read_card(d["doc_id"], p["page"])
            if card is not None:
                with guarded(self.rp, self.lab.ref(d["doc_id"], p["page"])):
                    self.check_page(Page(self, d["doc_id"], p, card))
        if not seen and (only or pages is not None):                 # 絞り込んだのに何にも当たらない。点検したつもりで、何も見ていない
            self.rp.fail(f"指定（--only {'・'.join(sorted(only or [])) or '—'}・--pages）に当たるページが無い。何も点検していない"
                         f"（ある資料: {'・'.join(sorted(self.order)) or 'なし'}）")
        with guarded(self.rp, "資料をまたぐ点検（ずれ・番号の並び）"):
            self.systematic_shift()
            self.number_jumps()
        self.summary()

    def read_card(self, doc, no):
        """カードを読み、形を確かめる。無い・読めない・形が崩れていて中身を見られないときは None"""
        rp, where, cp = self.rp, self.lab.ref(doc, no), card_json(self.work, doc, no)
        if not cp.exists():
            self.missing.append(where)
            return None
        try:
            c = json.loads(cp.read_text(encoding="utf-8"))
        except Exception as e:
            rp.fail(f"{where}: カードが JSON として読めない（{e}）")
            return None
        self.n["cards"] += 1
        errs = validate(c, "card")
        for e in errs[:SHAPE_ERRORS]:
            rp.fail(f"{where}: 形が違う — {e}")
        if len(errs) > SHAPE_ERRORS:
            rp.fail(f"{where}: 形が違う — ほか {len(errs) - SHAPE_ERRORS} 件（python3 scripts/validate.py --file card {cp} で全部出る）")
        if not isinstance(c, dict) or any(not isinstance(c.get(k), list) for k in ("steps", "decisions", "figures", "notes")) \
                or not isinstance(c.get("diagram"), dict) or c.get("page_type") not in PAGE_TYPES:
            return None
        return c

    # ---- カード1枚 ----

    def check_page(self, pg):
        c = pg.c
        if c.get("doc_id") != pg.doc or c.get("page") != pg.no:
            self.rp.fail(f"{pg.where}: カードの doc_id / page がファイルの場所と合わない（{c.get('doc_id')} / {c.get('page')}）")
        self.by_type[c["page_type"]] += 1
        self.item_ids(pg)
        self.quotes(pg)
        self.page_order(pg)
        self.unpointed(pg)
        self.title(pg)
        self.printed_page(pg)
        self.rendering(pg)
        self.contents(pg)
        self.note_actions(pg)
        self.omissions(pg)
        self.n["unread"] += len(c.get("unreadable") or [])
        self.n["uncertain"] += len(c.get("uncertain") or [])

    def item_ids(self, pg):
        c = pg.c
        ids = [x.get("id") for x in c["steps"] + c["decisions"] + c["notes"] + (c["diagram"].get("nodes") or []) if isinstance(x, dict)]
        for i in sorted({x for x in ids if x and ids.count(x) > 1}):
            self.rp.fail(f"{pg.where}: 項目の id が重複: {i}")
        step_ids = {s.get("id") for s in c["steps"]}
        for s in c["steps"]:
            if s.get("applies_to") and s["applies_to"] not in step_ids:
                self.rp.fail(f"{pg.where} {s.get('id')}: applies_to「{s['applies_to']}」という手順がこのカードに無い")
            self.n["note_steps"] += s.get("origin") == "note"

    def quotes(self, pg):
        """引用がそのページの機械抽出にあるか。無くて前後のページにあれば、どのページかを名指しする"""
        rp, doc, no = self.rp, pg.doc, pg.no
        for kind, place, key in QUOTED:
            for it in pg.c[place]:
                name = f"{pg.where} {kind} {it.get('id', '')}「{str(it.get(key, ''))[:20]}」"
                q = str(it.get("quote", "")).strip()
                if not q:
                    rp.fail(f"{name}: 原文の引用（quote）が無い")
                    continue
                if it.get("from_image"):
                    self.n["img_quotes"] += kind != "補足"
                    if "lines" in it:
                        rp.fail(f"{name}: 画像の中の文字（from_image）なのに、行の番号（lines）がある。機械抽出に無い文字は行で指せない（lines を消して quote を写すか、from_image を外す）")
                    continue
                self.checked[doc] += 1
                if "lines" in it:                                    # 行を番号で指した項目: 引用は、その行の文字そのままのはず（機械が入れる）
                    self.pointed(pg, it, name, q, str(it.get(key, "")))
                    continue
                r = quote_match(q, pg.mtext)
                if r >= 1.0:
                    continue
                other = next((x for x in pg.near if quote_match(q, self.mtext(doc, x)) >= 0.9), None) if r < QUOTE_OK else None
                if other is not None:
                    self.shifted.setdefault(doc, {}).setdefault(other - no, []).append(no)
                    rp.fail(f"{name}: 引用がこのページに無く、{self.lab.ref(doc, other)} にある（ページの取り違え。画像 {pg.png} を開き直して書き直す）")
                elif r < QUOTE_OK:
                    rp.fail(f"{name}: 引用が機械抽出の文字に見当たらない（一致 {r:.0%}）。画像から読んだなら from_image を true に、"
                            f"そうでなければ原文どおりに直す: 「{q[:40]}」")
                else:
                    rp.warn(f"{name}: 引用が機械抽出の文字と完全には一致しない（一致 {r:.0%}）。原文の連続した一部をそのまま写す: 「{q[:40]}」")

    def pointed(self, pg, it, name, q, said):
        """行の番号（lines）で原文を指した項目: 番号が範囲の中か、引用がその行の文字か、指した行が作業の文と合っているか"""
        rp, text = self.rp, span_text(pg.units, it["lines"])
        self.n["pointed"] += 1
        if text is None:
            rp.fail(f"{name}: 行の番号 {it['lines']} が範囲の外（このページの行は 1〜{len(pg.units)}。python3 scripts/pages.py lines で番号を見る）")
        elif norm(q) != norm(text):
            rp.fail(f"{name}: 引用（quote）が、指した行 {it['lines']} の文字と違う。引用は手で書かず、python3 scripts/quotes.py fill で入れる: 「{q[:30]}」／行「{text[:30]}」")
        elif CJK.search(said) and not CJK.search(text):              # 日本語以外の資料: 作業の文は日本語に訳して書くので、字の一致では見られない（照合レビューが見る）
            pass
        else:
            a, b = grams(norm(said), 2), grams(norm(text), 2)
            if a and len(a & b) / len(a) < POINT_MIN:
                rp.warn(f"{name}: 指した行 {it['lines']} が、書いた内容と合っていない疑い（別の行を指していないか）: 行「{text[:40]}」")

    def page_order(self, pg):
        """行を番号で指した手順: 並びが紙面の順と逆になっていないか／1つの文を、同じ担当の手順2つに分けていないか"""
        steps = [x for x in pg.c["steps"] if isinstance(x.get("lines"), list) and span_text(pg.units, x["lines"]) is not None and x.get("origin") != "note"]
        for x, y in zip(steps, steps[1:]):
            if y["lines"][0] < x["lines"][0] and x.get("section", "") == y.get("section", ""):
                self.rp.warn(f"{pg.where} 手順 {x.get('id')}→{y.get('id')}: 手順の並びが、紙面の行の順（{x['lines'][0]} 行目→{y['lines'][0]} 行目）と逆。"
                             f"原典の順のまま並べる（一般的な流れに合わせて入れ替えない。図や番号で順が決まっているなら、そのままでよい）")
            if x["lines"] == y["lines"] and x["lines"][0] == x["lines"][1] and norm(str(x.get("actor", ""))) == norm(str(y.get("actor", ""))):
                self.rp.warn(f"{pg.where} 手順 {x.get('id')}・{y.get('id')}: 同じ1文（{x['lines'][0]} 行目）を、同じ担当の手順2つに分けている。"
                             f"原文の1文は1つの手順にする（分けるのは、文の途中で担当が変わるときだけ。flow-rules.md 1.24）")

    def unpointed(self, pg):
        """行を番号で指したカード: 作業らしい文が、手順・分岐・補足のどれにも指されていない（図の箱の中の文を diagram にだけ書いた、など）"""
        items = [x for _kind, place, _key in QUOTED for x in pg.c[place]]
        if self.by_lines is None:                                    # この作業フォルダのカードが、行の番号を使っているか（1回だけ数える）
            self.by_lines = any('"lines"' in f.read_text(encoding="utf-8", errors="ignore") for f in sorted((self.work / "cards").glob("*/*.json")))
        if pg.c["page_type"] not in STEP_PAGES or not self.by_lines or not all("lines" in x or x.get("from_image") for x in items):
            return                                                   # 引用を手で写したカードは、文の番号と突き合わせられない
        held = set()
        for x in items:
            if "lines" in x and span_text(pg.units, x["lines"]) is not None:
                held |= set(range(x["lines"][0], x["lines"][1] + 1))
        said = [pg.units[k - 1] for k in sorted(held)]               # 同じ文が、文字層と図形の両方から出てくることがある。片方が指されていればよい
        for k, u in enumerate(pg.units, 1):
            if k not in held and action_like(u, self.markers) and not any(covers(x, u, 0.85) for x in said):
                self.n["unpointed"] += 1
                self.rp.warn(f"{pg.where}: 作業らしい文が、手順・分岐・補足のどれにも指されていない（{k} 行目）: 「{u[:50]}」"
                             f"（作業なら steps に書く。図の箱の中の文も steps に書く（flow-rules.md 1.7）。作業でなければ notes に why_not_action つきで書く）")

    def title(self, pg):
        """カードの見出しが、このページに無くて隣のページにある（ページの取り違え）"""
        title, here = norm(pg.c.get("title", "")), norm(pg.mtext)
        if len(title) < 4 or not here or title in here or pg.c["page_type"] == "blank":
            return
        hit = next((x for x in pg.near if abs(x - pg.no) == 1 and title in norm(self.mtext(pg.doc, x))), None)
        if hit is not None:
            self.rp.fail(f"{pg.where}: カードの見出し「{pg.c.get('title', '')[:24]}」がこのページに無く、{self.lab.ref(pg.doc, hit)} にある（ページの取り違え）")

    def printed_page(self, pg):
        """カードの printed_page が、機械がページの端から拾った番号と合っているか"""
        if pg.entry.get("kind") not in ("page", "slide"):
            return
        pp = pg.c.get("printed_page")
        self.printed.setdefault(pg.doc, []).append((pg.no, pp))
        e = self.lab.page(pg.doc, pg.no)
        if e["how"] != "machine":
            return
        if pp is None:
            self.rp.warn(f"{pg.where}: ページの端に番号「{e['printed']}」が刷られているのに printed_page が null")
        elif pg.pj.get("printed_candidates") and bare_number(pp) != bare_number(e["printed"]):   # 「- 13 -」「P.13」の飾りは外して比べる
            self.rp.fail(f"{pg.where}: カードの printed_page「{pp}」が、ページの端に刷られた数字「{e['printed']}」と違う"
                         f"（別のページを読んでいる疑い。画像 {pg.png} を開き直す。番号は合っていて書き方だけが違うなら、番号だけを書く）")

    def rendering(self, pg):
        """ページ化の段階で、画像と機械抽出の対応が取れていないもの"""
        if pg.pj.get("misaligned"):
            self.rp.fail(f"{pg.where}: ページ画像の文字と機械抽出の文字が合わない（一致 {pg.pj.get('align', 0):.0%}）。画像と機械抽出が別の"
                         f"スライド・シートのもの。render_pages.py --only {pg.doc} --force でやり直し、直らなければ画像だけを根拠に読む")
        if pg.entry.get("kind") == "sheet" and str(pg.entry.get("label", "")).startswith("シート不明"):
            self.rp.fail(f"{pg.where}: この画像がどのシートのものか対応が取れていない（出典にシート名を書けない）。"
                         f"render_pages.py --only {pg.doc} --force --backend office でやり直すか、画像の見出しからシートを確かめて利用者に知らせる")

    def contents(self, pg):
        """種別に見合う中身があるか"""
        c, dg = pg.c, pg.c["diagram"]
        gids = {n.get("id") for n in dg.get("nodes") or []}
        for e in dg.get("edges") or []:
            if e.get("from") not in gids or e.get("to") not in gids:
                self.rp.fail(f"{pg.where}: 図の線 {e.get('from')}→{e.get('to')} が、定義していない箱を指している")
        if c["page_type"] in FLOWISH and not (c["steps"] or c["decisions"] or dg.get("nodes")):
            self.rp.fail(f"{pg.where}: 種別が {c['page_type']} なのに、手順・分岐・図がどれも空")
        if c["page_type"] == "toc" and not c.get("toc_entries"):
            self.rp.warn(f"{pg.where}: 目次のページなのに toc_entries が空")
        if c["page_type"] in NON_CONTENT_TYPES:                      # 表紙・目次・白紙は、割当と網羅の点検から外れる。中身があるのに外れてはいけない
            n_items = len(c["steps"]) + len(c["decisions"]) + len(dg.get("nodes") or [])
            if n_items:
                self.rp.fail(f"{pg.where}: 種別が {c['page_type']} なのに、手順・分岐・図の箱が {n_items} 件ある"
                             f"（表紙・目次・白紙は、フローの網羅の点検から外れる。種別を見直す。決め方は flow-rules.md 1.14）")
            lines = [x for x in pg.mtext.split("\n") if len(norm(x)) >= 6]
            if c["page_type"] == "blank" and len(lines) >= BLANK_MAX_LINES:
                self.rp.fail(f"{pg.where}: 種別が blank なのに、機械抽出に文字が {len(lines)} 行ある（白紙ではない。画像を開いて読む）")

    def note_actions(self, pg):
        """補足の中の作業らしい文（機械が挙げた候補）が、手順・分岐に入っているか、「作業でない」と理由つきで記録されているか"""
        c = pg.c
        if c["page_type"] in NO_NOTE_ACTIONS:
            return
        item_q = [str(x.get("quote", "")) for x in c["steps"] + c["decisions"]]

        def held(q, sent):                                           # 見出しが文の頭につながって切れることがある。文末の 25 字が引用にあれば、その文と見る
            return covers(q, sent) or (len(norm(sent)) > 25 and norm(sent)[-25:] in norm(q))

        def lined(items, sent):
            """その文を、行の番号で指した項目があるか（文の番号で厳密に見る。引用の似かよりに頼らない）"""
            at = [k for k, u in enumerate(pg.units, 1) if u == sent]
            return any(isinstance(x.get("lines"), list) and len(x["lines"]) == 2 and any(x["lines"][0] <= k <= x["lines"][1] for k in at) for x in items)
        for sent in note_actions(pg.pj, self.markers):
            self.n["cand"] += 1
            if lined(c["steps"] + c["decisions"], sent) or any(held(q, sent) for q in item_q):
                self.n["cand_step"] += 1
            elif any((lined([x], sent) or held(str(x.get("quote", "")), sent)) and str(x.get("why_not_action", "")).strip() for x in c["notes"]):
                self.n["cand_note"] += 1
            else:
                self.rp.fail(f"{pg.where}: 補足の中に作業らしい文があるのに、手順にも「作業でない」の記録にも無い: 「{sent[:60]}」"
                             f"（作業なら steps に origin: note で書く。作業でなければ notes に why_not_action つきで書く）")

    def omissions(self, pg):
        """読み落としの疑い（WARN）: 機械抽出にあるもの（文字の行・画像・矢印・ひし形・小さい字）が、カードに出てこない"""
        rp, c, pj, where = self.rp, pg.c, pg.pj, pg.where
        dg = c["diagram"]
        if c["page_type"] in OMISSION_TYPES:
            lines = [x for x in (norm(x) for x in pg.mtext.split("\n")) if len(x) >= 6][:80]
            if len(lines) >= 4:
                ctext = norm(" ".join(strings(c)))
                cg = grams(ctext)
                hit = sum(1 for x in lines if x in ctext or len(grams(x) & cg) >= 0.6 * max(1, len(grams(x))))
                if hit / len(lines) < 0.4:
                    rp.warn(f"{where}: 機械抽出の文字 {len(lines)} 行のうち、カードに反映されたのは {hit} 行（読み落としの疑い。画像を見直す）")
        n_fig = len(c["figures"]) + (1 if dg.get("nodes") else 0)
        if ((pj.get("images") or 0) > 0 or pj.get("no_text_layer")) and n_fig == 0 and c["page_type"] != "blank":
            rp.warn(f"{where}: 画像があるのに figures / diagram が空（画像の中身を読んで書く）")
        elif pj.get("vector_figure") and n_fig == 0 and c["page_type"] not in ("blank", "toc", "cover"):
            rp.warn(f"{where}: 線で描いた図（箱・矢印）がありそうなのに figures / diagram が空（画像で図を読んで書く）")
        conns = [x for x in pj.get("connectors") or [] if x.get("from") and x.get("to")]
        if len(conns) >= 2 and len(dg.get("edges") or []) < 0.6 * len(conns):
            rp.warn(f"{where}: つながった矢印が {len(conns)} 本あるのに、カードの図の線は {len(dg.get('edges') or [])} 本（矢印の読み落としの疑い）")
        if any(any(k in str(s.get("type", "")).lower() for k in DECISION_SHAPES) for s in pj.get("shapes") or []) \
                and not c["decisions"] and not any(n.get("shape") == "decision" for n in dg.get("nodes") or []):
            rp.warn(f"{where}: ひし形（判断）の図形があるのに、分岐が書かれていない")
        if pj.get("small_text") and not list((self.work / "pages" / pg.doc / "crops").glob(f"p{pg.no:04d}_*.png")):
            rp.warn(f"{where}: 字が小さいページ。crop_page.py で拡大して読み直す")

    # ---- 全カードを見終えてから ----

    def systematic_shift(self):
        """同じ資料の引用が、同じ向きに何件もずれている（ページ画像とカードの対応が1つずつずれている）。先頭に出す"""
        for doc, by in self.shifted.items():
            for delta, at in by.items():
                if len(at) >= SHIFT_MIN and len(at) >= SHIFT_SHARE * max(1, self.checked[doc]):
                    self.rp.fail(f"{doc}: 系統的なずれ — カードの引用 {len(at)} 件が、{abs(delta)} ページ{'後ろ' if delta > 0 else '前'}のページにある"
                                 f"（{len(set(at))} ページ分）。ページ画像とカードの対応が1つずつずれている。その束を読み直す", first=True)

    def number_jumps(self):
        """刷られた番号の並びが1枚だけ飛ぶ（スキャン資料は、カードの読みしか手がかりが無い）"""
        for doc, rows in self.printed.items():
            nums = [(no, _split(pp)) for no, pp in rows if pp is not None and _split(pp)]
            for (a0, x0), (a1, x1), (a2, x2) in zip(nums, nums[1:], nums[2:]):
                if x0[0] == x1[0] == x2[0] and a0 - x0[1] == a2 - x2[1] != a1 - x1[1] and a2 - a0 == 2:
                    self.rp.fail(f"{self.lab.ref(doc, a1)}: printed_page「{x1[0]}{x1[1]}」が、前後のページの並び（{x0[1]}・{x2[1]}）から飛んでいる"
                                 f"（別のページを読んでいる疑い。画像 {doc}/p{a1:04d}.png を開き直す）")

    def summary(self):
        rp, n = self.rp, self.n
        if self.missing:
            rp.fail(f"カードが無いページ {len(self.missing)} 件: {', '.join(self.missing[:30])}" + (" …" if len(self.missing) > 30 else ""))
        rp.info(f"カード {n['cards']} 枚（" + "・".join(f"{k} {v}" for k, v in sorted(self.by_type.items())) + "）")
        rp.info(f"補足の中の作業: 手順にしたもの {n['note_steps']} 件（origin: note）。機械が挙げた候補 {n['cand']} 件のうち、"
                f"手順・分岐に入っている {n['cand_step']}・作業でないと記録 {n['cand_note']}・どちらでもない {n['cand'] - n['cand_step'] - n['cand_note']}")
        rp.info(f"画像から読んだ引用 {n['img_quotes']} 件（機械では照合できない。原典照合レビューで確かめる）")
        if n["pointed"]:
            rp.info(f"行の番号で指した引用 {n['pointed']} 件（引用の文字は機械が入れる）")
        if n["unread"]:
            rp.info(f"読めなかった箇所 {n['unread']} 件（確認事項にする）")
        if n["uncertain"]:
            rp.info(f"読めたが決められなかった点 {n['uncertain']} 件（フロー化のときに確認事項にする）")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="資料 ID（カンマ区切り）")
    ap.add_argument("--pages", help="ページの範囲（ファイル名の番号。例: 13-24 や 3,5,8-10）。--only と併せて、読んだ束だけを見るとき")
    ap.add_argument("--report")
    ap.add_argument("--all-lines", action="store_true", help="同じ種類の行を畳まずに全部出す")
    a = ap.parse_args()
    rp = Report("ページカードの点検（check_extract）", fold=None if a.all_lines else 10)
    CardChecks(work_dir(a.work), rp).run({x.strip() for x in a.only.split(",") if x.strip()} if a.only else None, parse_pages(a.pages))
    rp.finish(a.report)


if __name__ == "__main__":
    main()
