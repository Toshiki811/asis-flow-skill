#!/usr/bin/env python3
"""quotes.py — 引用（quote）を、行の番号（lines）から機械で入れる。LLM に原文を写させない。

  python3 scripts/quotes.py fill -w <作業フォルダ>                              # カードとフローの全部
  python3 scripts/quotes.py fill -w <作業フォルダ> --only M01 --pages 13-22     # 自分の束のカードだけ
  python3 scripts/quotes.py fill -w <作業フォルダ> --flow F03                   # 自分のフローだけ
  python3 scripts/quotes.py locate -w <作業フォルダ>                            # いまの引用から、行の番号を引く（前の版のカードを移すとき）
  python3 scripts/quotes.py locate -w <作業フォルダ> --rewrite                  # 引いた行の文字で、引用を書き換える

行の番号は、ページの機械抽出を文の単位に切った並びの、何番目か（1 から）。python3 scripts/pages.py lines が番号つきで見せる。
  カードの手順・分岐・補足    lines: [開始, 終了] を書く。quote は、ここが入れる（その行の文字そのまま）
  フローの出典               item があれば、カードのその項目の quote と lines を写す。item の無い出典は、lines から入れる
  画像の中の文字             機械抽出に無いので、行では指せない。いままでどおり from_image: true と、手で写した quote を書く（lines は書かない）
写し間違い・言い換え・表記ゆれが、形の上で起きなくなる。代わりに起きうるのは「別の行を指す」誤りで、点検（check_extract.py）が、
作業の文と行の文字の重なりが少ないものを WARN にする。
locate は、番号を書いていない引用（前の版のカード）について、その引用を含むいちばん短い行の範囲を引く。読み直しはしない。
"""
import argparse
import pathlib

from common import all_pages, card_json, load, load_page, manifest, norm, page_json, page_units, parse_pages, save, span_text, work_dir

QUOTED = ("steps", "decisions", "notes")


class Units:
    """(資料, page) → 行の並び。1回切ったら覚えておく"""

    def __init__(self, work):
        self.work, self.cache = work, {}

    def __call__(self, doc, page):
        if (doc, page) not in self.cache:
            self.cache[(doc, page)] = page_units(load_page(page_json(self.work, doc, page)))
        return self.cache[(doc, page)]


def card_files(work, m, only, pages):
    for d, p in all_pages(m):
        if (only and d["doc_id"] not in only) or (pages is not None and p["page"] not in pages):
            continue
        f = card_json(work, d["doc_id"], p["page"])
        if f.exists():
            yield d["doc_id"], p["page"], f


def fill_card(c, units, bad, where):
    """カード1枚: lines のある項目の quote を入れる。入れた数を返す"""
    n = 0
    for place in QUOTED:
        for it in c.get(place) or []:
            if not isinstance(it, dict) or "lines" not in it or it.get("from_image"):
                continue
            text = span_text(units, it["lines"])
            if text is None:
                bad.append(f"{where} {it.get('id')}: 行の番号 {it['lines']} が範囲の外（このページの行は 1〜{len(units)}）")
            elif it.get("quote") != text:
                it["quote"] = text
                n += 1
    return n


def fill_flow(fl, work, units, bad):
    """フロー1本: 出典の quote を入れる。item があればカードの項目から写し、無ければ lines から入れる"""
    n, cards = 0, {}

    def item(doc, page, iid):
        if (doc, page) not in cards:
            cards[(doc, page)] = load(card_json(work, doc, page), {}) or {}
        c = cards[(doc, page)]
        return next((x for place in ("steps", "decisions") for x in c.get(place) or [] if x.get("id") == iid), None)
    for node in fl.get("nodes") or []:
        for s in node.get("sources") or []:
            doc, page = s.get("doc_id"), s.get("page")
            where = f"{fl.get('flow_id')} {node.get('id')}"
            new = None
            if "lines" in s:
                new = span_text(units(doc, page), s["lines"])
                if new is None:
                    bad.append(f"{where}: 行の番号 {s['lines']} が範囲の外（{doc}/p{int(page or 0):04d} の行は 1〜{len(units(doc, page))}）")
            elif s.get("item"):
                it = item(doc, page, s["item"])
                if it is not None and "lines" in it and not it.get("from_image"):   # カードが番号で指している項目は、その引用をそのまま使う
                    new = it.get("quote")
                elif it is not None and not s.get("quote") and it.get("quote"):    # 手で写した・画像の中の文字の項目: 出典に引用が無ければ、項目のものを写す
                    s["quote"], s["from_image"] = it["quote"], bool(it.get("from_image"))
                    n += 1
            if new and (s.get("quote") != new or s.get("from_image") is not False):
                s["quote"], s["from_image"] = new, False
                n += 1
    return n


def cmd_fill(a, work, m):
    units, bad, n_cards, n_items = Units(work), [], 0, 0
    only = set(a.only.split(",")) if a.only else None
    if not a.flow:
        for doc, page, f in card_files(work, m, only, parse_pages(a.pages)):
            c = load(f)
            k = fill_card(c, units(doc, page), bad, f"{doc}/p{page:04d}") if isinstance(c, dict) else 0
            if k:
                save(f, c)
                n_cards, n_items = n_cards + 1, n_items + k
        print(f"カード {n_cards} 枚の引用 {n_items} 件を、行の文字から入れた")
    if not a.only and not a.pages:
        want, n_flows, n_src = (set(a.flow.split(",")) if a.flow else None), 0, 0
        for f in sorted((work / "flows").glob("*.flow.json")):
            fl = load(f)
            if not isinstance(fl, dict) or (want and fl.get("flow_id") not in want):
                continue
            k = fill_flow(fl, work, units, bad)
            if k:
                save(f, fl)
                n_flows, n_src = n_flows + 1, n_src + k
        print(f"フロー {n_flows} 本の出典 {n_src} 件の引用を入れた")
    for x in bad:
        print("  ★ " + x)
    raise SystemExit(1 if bad else 0)


def find_span(units, quote, reach=12):
    """引用を含む、いちばん短い行の範囲 [開始, 終了]（1 から）。見つからなければ None"""
    q, nu, best = norm(quote), [norm(u) for u in units], None
    if not q:
        return None
    for a in range(len(nu)):
        text = ""
        for b in range(a, min(len(nu), a + reach)):
            text += nu[b]
            if q in text:
                if best is None or b - a < best[1] - best[0]:
                    best = [a + 1, b + 1]
                break
    return best


def cmd_locate(a, work, m):
    units, found, lost, same = Units(work), 0, [], 0
    only = set(a.only.split(",")) if a.only else None
    for doc, page, f in card_files(work, m, only, parse_pages(a.pages)):
        c, changed = load(f), False
        for place in QUOTED:
            for it in (c.get(place) if isinstance(c, dict) else None) or []:
                if not isinstance(it, dict) or it.get("from_image") or "lines" in it or not str(it.get("quote", "")).strip():
                    continue
                span = find_span(units(doc, page), it["quote"])
                if span is None:
                    lost.append(f"{doc}/p{page:04d} {it.get('id')}「{str(it['quote'])[:30]}」")
                    continue
                text = span_text(units(doc, page), span)
                found, same = found + 1, same + (norm(text) == norm(it["quote"]))
                if a.rewrite:
                    it["lines"], it["quote"], changed = span, text, True
        if changed:
            save(f, c)
    print(f"行の番号が引けた引用 {found} 件（うち、行の文字と同じもの {same} 件。残りは、行の一部だけを引用している）"
          + ("。引用を行の文字で書き換えた" if a.rewrite else "。書き換えていない（--rewrite で書き換える）"))
    if lost:
        print(f"行が見つからない引用 {len(lost)} 件（機械抽出に無い。画像の中の文字なら from_image を true にする）:")
        for x in lost[:20]:
            print("  " + x)
    if a.rewrite:
        print("このあと: python3 scripts/quotes.py fill -w <作業フォルダ>（フローの出典に写す）→ check_extract.py・check_flow.py")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["fill", "locate"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="資料 ID（カンマ区切り）。カードだけを対象にする")
    ap.add_argument("--pages", help="ページの範囲（ファイル名の番号。例: 13-24）")
    ap.add_argument("--flow", help="fill: フロー ID（カンマ区切り）。フローだけを対象にする")
    ap.add_argument("--rewrite", action="store_true", help="locate: 引いた行の番号と文字を、カードに書く")
    a = ap.parse_args()
    work = work_dir(a.work)
    m = manifest(work)
    (cmd_fill if a.cmd == "fill" else cmd_locate)(a, pathlib.Path(work), m)


if __name__ == "__main__":
    main()
