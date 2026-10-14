#!/usr/bin/env python3
"""pages.py — ページの束・要約・大きさの見積り。大量の入力を、1つのエージェントの文脈に収まる大きさに割ってさばくための小道具。

  python3 scripts/pages.py batches -w <WORK>                 # まだカードの無いページを、読取の束に分ける（1束の見込みが予算内）
  python3 scripts/pages.py batches -w <WORK> --budget 80000 --all
  python3 scripts/pages.py batches -w <WORK> --only M01 --pages 13-22    # 資料とページの範囲を決めて束にする
  python3 scripts/pages.py outline -w <WORK>                 # 資料・章ごとの集計（全体を見渡す。行数は章の数）
  python3 scripts/pages.py digest -w <WORK> --only M03 --pages 10-40   # ページ1行の要約（範囲を決めて見る）
  python3 scripts/pages.py pagemap -w <WORK>                 # 刷られたページ番号とファイルの番号の対応。目次の番号 → 実際のページ
  python3 scripts/pages.py terms -w <WORK>                   # 担当・システム・帳票の表記の一覧と、同じものかもしれない組（用語の正本を書く材料）
  python3 scripts/pages.py terms -w <WORK> --draft           # 用語の正本（glossary.json）に、まだ無い表記を別々の用語として足す（寄せない）
  python3 scripts/pages.py links -w <WORK>                   # 資料どうし・業務どうしの紐付けの候補（参照の文・共有している帳票とシステム・別の業務の名指し）
  python3 scripts/pages.py lines -w <WORK> --only M01 --pages 13-18    # 機械抽出を、行の番号つきで見る（引用を番号で指すとき。刷られた番号・図形の数も出す）
  python3 scripts/pages.py plan-flows -w <WORK>              # 業務分類ごとの材料の大きさ（1エージェントで読み切れるか）
  python3 scripts/pages.py rows -w <WORK>                    # 縦軸（業務マトリクスの行）の案と、目次の項目の扱い。taxonomy.json が無ければ、階層から組んだ下書きで出す
  python3 scripts/pages.py taxonomy -w <WORK> --draft        # 目次に準拠した taxonomy.json の下書き（分類・番号・目次の場所・ページの割当）を機械で書く
  python3 scripts/pages.py orphans -w <WORK>                 # どの分類にも入っていない内容ページのまとまりと、結び付きそうな目次の項目の候補

batches     束は同じ資料の連続したページ。見込みトークン（固定分＋ページごと。字が小さく拡大して読むページは倍）が --budget に収まるように切る。
            既定の予算は 10 万（20 万の文脈の半分。残りは指示文・規約・やり直しの余地）。前の束が読み残したページは、半分の予算で束ね直す。
            途中で止まっても、もう一度実行すれば残りだけが出る。束の一覧は batches.json にも保存する。
outline     親（進行役）は、カードを1枚ずつ開かずにこれで全体を見る。資料が多いと 1行＝1資料になるので、--only で資料を選んで章ごとに見る。
            ページ1行の digest は、全体が 150 ページを超えると範囲の指定が要る。どの一覧も、画面には --show 行までしか出さない（全件はファイル）。
plan-flows  材料（カードの文字数）が上限を超える分類は、フロー化の前に下位の分類へ分ける（check_flow.py --stage structure が FAIL にする）。
taxonomy    目次に業務の一覧があるとき（hierarchy.json の toc_flow_list）、名称・順序・番号・ページの割当は機械が組む。LLM が決めるのは、
            フローを描くか・対象外か（out_of_scope）と、目次に無い部分だけ。taxonomy.json がすでにあれば、分類には触らず、
            対象外にした分類のページを unassigned_pages に足すだけ。
orphans     目次に無い業務を足す前に見る。結び付けるのは、根拠の文があるとき・見出しが目次の中で一意に一致するときだけ（flow-rules.md 2.20）。
            --accept で、見出しが一意に一致するものを機械が結び付ける。根拠の文のある候補は、番号（C1）で採る
links       候補に番号（R1・S1）が付く。採るものの番号を --accept R1,S2 で渡すと、根拠の文・ページごと hierarchy.json に写す（LLM は書き写さない）
"""
import argparse
import difflib
import re

from common import (MATERIAL_CAP, NON_CONTENT_TYPES, TYPE_LABEL, all_pages, card_json, file_stem, labels, load, load_card, load_page, manifest, norm,
                    page_json, page_text, page_units, parse_pages, save, tree_nodes, work_dir)

FIXED_TOKENS = 15000      # 1束の固定分（指示文・規約・スキーマ・点検のやり直し）
PAGE_TOKENS = 12000       # 1ページ（画像を見る・機械抽出を読む・カードを書く）。59 ページの実測（1束 12 ページで 16〜23 万）から
TEXT_PAGE_TOKENS = 3000   # 画像の無いページ（マクロのソース・文字のファイル）
BUDGET = 100000           # 1束の予算（20 万の文脈の半分）
MAX_BATCH = 12            # 予算に収まっても、1束はこの枚数まで
DIGEST_LIMIT = 150        # これを超える入力で、範囲を指定しない digest は outline に切り替える


def page_cost(work, p):
    """1ページを読んでカードにするのに使う見込みトークン"""
    pj = load(work / p["json"], {}) or {}
    chars = len(page_text(pj))
    if not p.get("png"):
        return TEXT_PAGE_TOKENS + 3 * chars
    return PAGE_TOKENS * (2 if pj.get("small_text") else 1) + 2 * chars


def rng(lab, doc, ps):
    """ページの集まりの書き方（人に見せる番号＋ファイルの番号）"""
    show = lab.short(doc, ps[0]) if len(ps) == 1 else f"{lab.short(doc, ps[0])}〜{lab.short(doc, ps[-1])}"
    files = f"p{ps[0]:04d}" if len(ps) == 1 else f"p{ps[0]:04d}〜p{ps[-1]:04d}"
    return f"{show}〔{files}〕"


def cmd_batches(a, work, m, lab, only):
    left, bad = set(), []                                            # 前の束が読み残したページ／形の崩れた報告
    from validate import validate
    for f in sorted((work / "reports").glob("read_*.json")) if (work / "reports").exists() else []:
        r = load(f, {}) or {}
        bad += [f"reports/{f.name}: {e}" for e in validate(r, "report_read")[:3]]
        left |= {(r.get("doc_id"), p) for p in r.get("pages_left") or []}
    if bad:
        raise SystemExit("読取の報告の形が崩れている。直してから束を作る（読み残したページを束に戻すため）:\n  " + "\n  ".join(bad))
    batches, cur, total, done = [], None, 0, 0
    for d, p in all_pages(m):
        if (only and d["doc_id"] not in only) or (a.page_set is not None and p["page"] not in a.page_set):
            continue
        total += 1
        if card_json(work, d["doc_id"], p["page"]).exists() and not a.all:
            done += 1
            cur = None
            continue
        retry = (d["doc_id"], p["page"]) in left
        cost = page_cost(work, p)
        cap = (a.budget // 2 if retry else a.budget) - FIXED_TOKENS
        if cur is None or cur["doc_id"] != d["doc_id"] or cur["retry"] != retry or len(cur["pages"]) >= a.max_pages \
                or (cur["pages"] and cur["est_tokens"] - FIXED_TOKENS + cost > cap):
            cur = {"batch": len(batches) + 1, "doc_id": d["doc_id"], "doc": d["rel_path"], "pages": [], "est_tokens": FIXED_TOKENS, "retry": retry}
            batches.append(cur)
        cur["pages"].append(p["page"])
        cur["est_tokens"] += cost
    save(work / "batches.json", batches)
    est = sum(b["est_tokens"] for b in batches)
    print(f"全 {total} ページ・カード済み {done}・残り {total - done} ページ → {len(batches)} 束"
          f"（1束の予算 {a.budget:,} トークン。見込みの合計 {est:,} トークン）")
    over = [b for b in batches if b["est_tokens"] > a.budget]
    for b in batches[:a.show] if a.show else batches:
        print(f"  束{b['batch']:>3}: {b['doc_id']} {rng(lab, b['doc_id'], b['pages'])}（{len(b['pages'])} ページ・見込み {b['est_tokens']:,}）"
              + ("  ※読み残しのやり直し" if b["retry"] else "") + f"  {b['doc']}  --only {b['doc_id']} --pages {b['pages'][0]}-{b['pages'][-1]}")
    if a.show and len(batches) > a.show:
        print(f"  … ほか {len(batches) - a.show} 束（一覧は {work / 'batches.json'}）")
    if over:
        print(f"※ 1ページだけで予算を超える束が {len(over)} 件（{', '.join('束' + str(b['batch']) for b in over[:8])}）。そのページは1ページずつ読む")
    if batches:
        print(f"\n1束を1つのサブエージェントに渡す（同時に 4〜6 束）。行末の --only … --pages … は、その束だけを点検するときにそのまま付ける。"
              f"ページ画像: pages/<資料ID>/pNNNN.png、機械抽出: 同じ名前の .json、"
              f"カード: cards/<資料ID>/pNNNN.json、報告: reports/read_<資料ID>_<先頭ページ>.json")


def card_of(work, d, p):
    return load(card_json(work, d["doc_id"], p["page"]))


def cmd_digest(a, work, m, lab, only):
    pages = a.page_set
    rows = [(d, p) for d, p in all_pages(m) if (not only or d["doc_id"] in only) and (pages is None or p["page"] in pages)]
    if len(rows) > DIGEST_LIMIT and not a.force:
        print(f"{len(rows)} ページある。ページ1行の要約は {DIGEST_LIMIT} ページまで（--only・--pages で範囲を決める。全部出すなら --force）。"
              f"全体は章ごとの集計で見る:\n")
        return cmd_outline(a, work, m, lab, only)
    for d, p in rows:
        c = card_of(work, d, p)
        head = f"{lab.ref(d['doc_id'], p['page'])}"
        if c is None:
            print(f"{head} （カードなし）")
            continue
        flags = []
        if c.get("steps"):
            n = sum(s.get("origin") == "note" for s in c["steps"])
            flags.append(f"手順{len(c['steps'])}" + (f"（うち補足{n}）" if n else ""))
        if c.get("decisions"):
            flags.append(f"分岐{len(c['decisions'])}")
        if (c.get("diagram") or {}).get("nodes"):
            flags.append(f"図{len(c['diagram']['nodes'])}")
        if c.get("unreadable") or c.get("uncertain"):
            flags.append("要確認")
        sec = " > ".join(c.get("section_path") or [])
        who = "・".join((c.get("actors") or [])[:4])
        print(f"{head} [{c.get('page_type', '?'):<12}] {str(c.get('title', ''))[:30]}"
              + (f"  〔{sec[:40]}〕" if sec else "") + (f"  担当: {who}" if who else "") + (f"  {' '.join(flags)}" if flags else ""))


def cmd_outline(a, work, m, lab, only):
    """資料・章ごとの集計。資料が少なければ 1行＝1つの章（section_path の先頭 depth 段）。多ければ 1行＝1資料にして、--only で掘り下げる"""
    rows = []
    for d in m["docs"]:
        if d.get("duplicate_of") or not d.get("page_list") or (only and d["doc_id"] not in only):
            continue
        groups, nocard = [], 0
        for p in d["page_list"]:
            c = load(card_json(work, d["doc_id"], p["page"]))
            if c is None:
                nocard += 1
                continue
            key = " > ".join((c.get("section_path") or [])[:a.depth]) or f"（見出しなし・{c.get('page_type')}）"
            if not groups or groups[-1]["key"] != key:
                groups.append({"key": key, "pages": [], "types": {}, "steps": 0, "notes": 0, "dec": 0, "ask": 0, "chars": 0, "who": {}})
            g = groups[-1]
            g["pages"].append(p["page"])
            g["types"][c.get("page_type")] = g["types"].get(c.get("page_type"), 0) + 1
            g["steps"] += len(c.get("steps") or [])
            g["notes"] += sum(s.get("origin") == "note" for s in c.get("steps") or [])
            g["dec"] += len(c.get("decisions") or [])
            g["ask"] += len(c.get("unreadable") or []) + len(c.get("uncertain") or [])
            g["chars"] += len(card_json(work, d["doc_id"], p["page"]).read_text(encoding="utf-8"))
            for w in c.get("actors") or []:
                g["who"][w] = g["who"].get(w, 0) + 1
        rows.append((d, groups, nocard))
    save(work / "structure" / "outline.json", [
        {"doc_id": d["doc_id"], "doc": d["rel_path"], "pages": len(d["page_list"]), "no_card": nocard,
         "sections": [{"title": g["key"], "pages": g["pages"], "steps": g["steps"], "note_steps": g["notes"], "decisions": g["dec"],
                       "ask": g["ask"], "card_chars": g["chars"]} for g in groups]} for d, groups, nocard in rows])
    limit = a.show or 10 ** 9
    if sum(1 + len(g) for _d, g, _n in rows) > 3 * limit:            # 行が多すぎる。1行＝1資料にする
        print(f"資料 {len(rows)} 件（1行＝1資料。章ごとに見るには --only で資料を選ぶ。全件は structure/outline.json）")
        for d, groups, nocard in rows[:limit]:
            tot = lambda k: sum(g[k] for g in groups)
            print(f"  {d['doc_id']} {d['rel_path'][:44]}｜{len(d['page_list'])} ページ" + (f"（カードなし {nocard}）" if nocard else "")
                  + f"｜章 {len(groups)}｜手順 {tot('steps')}" + (f"（補足 {tot('notes')}）" if tot("notes") else "")
                  + f"・分岐 {tot('dec')}・要確認 {tot('ask')}｜カード {tot('chars'):,} 字")
        if len(rows) > limit:
            print(f"  … ほか {len(rows) - limit} 件")
        return
    for d, groups, nocard in rows:
        print(f"{d['doc_id']} {d['rel_path']}（{len(d['page_list'])} ページ" + (f"・カードなし {nocard}" if nocard else "") + "）")
        for g in groups:
            who = "・".join(k for k, _v in sorted(g["who"].items(), key=lambda x: -x[1])[:4])
            types = "・".join(f"{k}{v}" for k, v in g["types"].items())
            print(f"  {rng(lab, d['doc_id'], g['pages'])} {g['key'][:36]}｜{len(g['pages'])} ページ（{types}）｜手順 {g['steps']}"
                  + (f"（補足 {g['notes']}）" if g["notes"] else "") + f"・分岐 {g['dec']}・要確認 {g['ask']}｜カード {g['chars']:,} 字" + (f"｜{who}" if who else ""))


def cmd_pagemap(a, work, m, lab, only):
    out = {"docs": [], "toc": []}
    for d in m["docs"]:
        if d.get("duplicate_of") or not d.get("page_list") or (only and d["doc_id"] not in only):
            continue
        info = lab.docs.get(d["doc_id"], {})
        rows = [{"page": p["page"], "shown": lab.short(d["doc_id"], p["page"]), "how": lab.page(d["doc_id"], p["page"])["how"]} for p in d["page_list"]]
        out["docs"].append({"doc_id": d["doc_id"], "basis": info.get("basis"), "offsets": info.get("offsets"), "pages": rows})
        basis = {"printed": "刷られた番号で出す", "count": "番号の印字なし。先頭から数えた番号で出す", "sheet": "シート名で出す"}.get(info.get("basis"), "")
        print(f"{d['doc_id']} {d['rel_path']}: {basis}")
        if info.get("basis") == "printed":
            runs = []
            for r in rows:                                           # 同じ差が続く区間にまとめる
                e = lab.page(d["doc_id"], r["page"])
                off = r["page"] - int(re.search(r"(\d+)$", e["printed"]).group(1)) if e["printed"] else None
                if runs and runs[-1]["off"] == off:
                    runs[-1]["pages"].append(r["page"])
                else:
                    runs.append({"off": off, "pages": [r["page"]]})
            for r in runs:
                ps = r["pages"]
                files = f"p{ps[0]:04d}" if len(ps) == 1 else f"p{ps[0]:04d}〜p{ps[-1]:04d}"
                if r["off"] is None:
                    print(f"  {files}: 番号が刷られていない → " + "、".join(lab.short(d["doc_id"], x) for x in ps[:6]) + (" …" if len(ps) > 6 else ""))
                else:
                    print(f"  {files}: {lab.short(d['doc_id'], ps[0])}〜{lab.short(d['doc_id'], ps[-1])}"
                          f"（ファイルの番号 − 刷られた番号 ＝ {r['off']}）")
            bad = [r["page"] for r in rows if r["how"] == "machine" and (load(card_json(work, d["doc_id"], r["page"]), {}) or {}).get("printed_page")
                   not in (None, lab.page(d["doc_id"], r["page"])["printed"])]
            if bad:
                print(f"  ※ カードの printed_page が、ページの端の数字と違うページ: {', '.join(f'p{x:04d}' for x in bad[:12])}（check_extract.py が FAIL にする）")
        # 目次の番号 → 実際のページ
        for p in d["page_list"]:
            c = load(card_json(work, d["doc_id"], p["page"]), {}) or {}
            for t in c.get("toc_entries") or []:
                if t.get("page") in (None, ""):
                    continue
                hit = lab.find(d["doc_id"], t["page"]) if info.get("basis") == "printed" else (
                    int(t["page"]) if str(t["page"]).isdigit() and any(x["page"] == int(t["page"]) for x in d["page_list"]) else None)
                out["toc"].append({"doc_id": d["doc_id"], "number": t.get("number"), "title": t.get("title"), "toc_page": t["page"], "page": hit})
                print(f"  目次「{t.get('number', '')} {t.get('title', '')}」… {t['page']} → "
                      + (f"{d['doc_id']}/p{hit:04d}" if hit else "当たるページが見つからない（番号を確かめる）"))
    save(work / "structure" / "pagemap.json", out)
    print(f"\n対応表 → {work / 'structure' / 'pagemap.json'}。データ（taxonomy.json・flow.json の page）にはファイルの番号を書く。"
          f"目次の番号から割り当てるときは、上の「→」の右を使う（足し引きしない）")


def draft_glossary(work, g, seen):
    """用語の正本の下書き: まだ載っていない表記を、別々の用語として足す。寄せない（同じものかどうかは、原典で確かめた人が決める）。
    すでにある用語・別の表記・確認事項には触らない"""
    g = {"terms": list(g.get("terms") or []), "questions": list(g.get("questions") or [])}
    known = {norm(x) for t in g["terms"] for x in [t.get("name")] + list(t.get("aliases") or [])}
    added = 0
    for kind in ("actor", "system", "document"):
        for name, _where in sorted(seen[kind].items(), key=lambda x: (-len(x[1]), x[0])):
            if norm(name) not in known:
                known.add(norm(name))
                g["terms"].append({"kind": kind, "name": name, "aliases": []})
                added += 1
    save(work / "glossary.json", g)
    print(f"用語の正本 → {work / 'glossary.json'}（{added} 件を足した。計 {len(g['terms'])} 件）。足したものは別々の用語のまま。"
          f"上の「同じものかもしれない組」を原典で確かめ、同じと確かめられたものだけ aliases に寄せて evidence を書く。"
          f"確かめられない組は、寄せずに確認事項（種別「用語」）にする")


def cmd_terms(a, work, m, lab, only):
    seen = {"actor": {}, "system": {}, "document": {}}
    for d, p in all_pages(m):
        c = load(card_json(work, d["doc_id"], p["page"]), {}) or {}
        bag = {"actor": list(c.get("actors") or []) + [s.get("actor") for s in c.get("steps") or []],
               "system": list(c.get("systems") or []) + [s.get("system") for s in c.get("steps") or []],
               "document": list(c.get("documents") or []) + [x for s in c.get("steps") or [] for x in (s.get("inputs") or []) + (s.get("outputs") or [])]}
        for kind, names in bag.items():
            for n in {str(x).strip() for x in names if x and str(x).strip()}:
                seen[kind].setdefault(n, []).append((d["doc_id"], p["page"]))
    g = load(work / "glossary.json", {}) or {}
    known = {norm(x) for t in g.get("terms") or [] for x in [t.get("name")] + list(t.get("aliases") or [])}
    out = {"terms": [], "similar": []}
    for kind, title in (("actor", "担当"), ("system", "システム"), ("document", "帳票・ファイル")):
        rows = sorted(seen[kind].items(), key=lambda x: (-len(x[1]), x[0]))
        print(f"■ {title}（{len(rows)} 種類）")
        for name, where in rows[:a.show or len(rows)]:
            docs = sorted({w[0] for w in where})
            mark = "" if not g or norm(name) in known else "  ※用語の正本に無い"
            print(f"  {name}｜{len(where)} ページ｜{'・'.join(docs)}｜例: {'、'.join(lab.src({'doc_id': w[0], 'page': w[1]}) for w in where[:3])}{mark}")
            out["terms"].append({"kind": kind, "name": name, "pages": len(where), "docs": docs,
                                 "where": [{"doc_id": w[0], "page": w[1]} for w in where[:20]]})
        if a.show and len(rows) > a.show:
            print(f"  … ほか {len(rows) - a.show} 種類（全件は structure/terms.json）")
        names = [n for n, _w in rows]
        for i, x in enumerate(names):
            for y in names[i + 1:]:
                nx, ny = norm(x), norm(y)
                if min(len(nx), len(ny)) >= 2 and (nx in ny or ny in nx or difflib.SequenceMatcher(None, nx, ny).ratio() >= 0.8):
                    out["similar"].append({"kind": kind, "a": x, "b": y})
    if out["similar"]:
        print(f"■ 同じものかもしれない組（{len(out['similar'])} 組）— 原典で同じと確かめられたものだけ aliases に寄せる。確かめられなければ別の用語のまま確認事項にする")
        for s in out["similar"][:a.show or len(out["similar"])]:
            print(f"  「{s['a']}」と「{s['b']}」")
    save(work / "structure" / "terms.json", out)
    print(f"\n一覧 → {work / 'structure' / 'terms.json'}。これを見て glossary.json（用語の正本）を書く"
          + ("" if a.draft else "（--draft で、まだ無い表記を別々の用語として足せる）"))
    if a.draft:
        draft_glossary(work, g, seen)


REF_KIND = [("appendix_of", re.compile(r"別紙|別添|別表|様式|付録|添付")), ("basis_of", re.compile(r"に基づ|による|に定め|に従|規程|規則|要領|基準|細則")),
            ("refers_to", re.compile(r"参照|を見|詳細は|記載|とおり"))]


def flow_targets(c, cats):
    """分類 c を指す参照の行き先。フローを作らない分類（章）なら、その下のフローを作る分類"""
    if c.get("make_flow"):
        return [c]
    return [t for k in cats if k.get("parent") == c.get("id") for t in flow_targets(k, cats)]


def named_refs(work, m, tax, where, seen):
    """カードの参照の文が、別の業務分類を番号と名前で名指ししているもの（「第8章 備品の購入」を参照）。同じ資料の中の章どうしの受け渡しは、
    帳票の共有には出ないことが多い。名前だけの当たりは採らない（「出張手続」は「海外出張手続」にも含まれる）"""
    bare = lambda t: re.sub(r"[「」『』（）()]", "", norm(t))
    cats = [c for c in tax.get("categories") or [] if c.get("number") and c.get("name")]
    out = []
    for d, p in all_pages(m):
        here = [c for c in where.get((d["doc_id"], p["page"]), []) if c.get("make_flow") and c.get("hierarchy_id")]
        c = load(card_json(work, d["doc_id"], p["page"]), {}) or {}
        for x in c.get("cross_refs") or []:
            t = bare(str(x.get("text", "")))
            for hit in [k for k in cats if bare(str(k["number"]) + k["name"]) in t]:
                for frm in here:
                    for to in flow_targets(hit, tax.get("categories") or []):
                        pair = (frm["hierarchy_id"], to.get("hierarchy_id"))
                        if to.get("hierarchy_id") and pair[0] != pair[1] and pair not in seen:
                            seen.add(pair)
                            out.append({"from": pair[0], "to": pair[1], "via": "名指しの参照", "evidence": x.get("text", ""),
                                        "sources": [{"doc_id": d["doc_id"], "page": p["page"]}]})
    return out


def cmd_links(a, work, m, lab, only):
    """紐付けの候補を挙げる。採るかどうかは、根拠の文を読んで1件ずつ決める（hierarchy.json に書く）"""
    docs = [d for d in m["docs"] if not d.get("duplicate_of") and d.get("page_list")]
    stems = {d["doc_id"]: norm(file_stem(d["name"])) for d in docs}
    out = {"relations": [], "shared": [], "sequence": []}
    seen = set()
    for d, p in all_pages(m):
        c = load(card_json(work, d["doc_id"], p["page"]), {}) or {}
        for x in c.get("cross_refs") or []:
            t = norm(str(x.get("target", "")) + str(x.get("text", "")))
            for other, stem in stems.items():
                # 資料名から「マニュアル」などを外した芯でも当てる。ただし、ほかの資料名にも含まれる芯（「経費精算」など）は、どの資料か決まらないので使わない
                core = re.sub(r"(マニュアル|手順書|規程|規則|要領|一覧|\d+|第\d+版|v\d+)", "", stem)
                unique = len(core) >= 3 and not any(core in x for o, x in stems.items() if o != other)
                if other == d["doc_id"] or len(stem) < 3 or not (stem in t or (unique and core in t)):
                    continue
                kind = next((k for k, r in REF_KIND if r.search(str(x.get("text", "")))), "refers_to")
                frm, to = (other, d["doc_id"]) if kind == "basis_of" and re.search(r"に基づ|による|に従", str(x.get("text", ""))) else (d["doc_id"], other)
                if (frm, to, kind) in seen:
                    continue
                seen.add((frm, to, kind))
                out["relations"].append({"from": frm, "to": to, "type": kind, "evidence": x.get("text", ""),
                                         "sources": [{"doc_id": d["doc_id"], "page": p["page"]}]})
    tax = load(work / "taxonomy.json", {}) or {}
    where = {}
    for c in tax.get("categories") or []:
        for ps in c.get("pages") or []:
            for pg in ps.get("pages") or []:
                where.setdefault((ps.get("doc_id"), pg), []).append(c)
    use = {}
    for d, p in all_pages(m):
        c = load(card_json(work, d["doc_id"], p["page"]), {}) or {}
        for s in c.get("steps") or []:
            for kind, names in (("in", s.get("inputs") or []), ("out", s.get("outputs") or [])):
                for n in names:
                    for cat in where.get((d["doc_id"], p["page"]), [{"id": None, "hierarchy_id": None}]):
                        use.setdefault(str(n).strip(), []).append((kind, d["doc_id"], p["page"], cat.get("id"), cat.get("hierarchy_id"), s.get("quote", "")))
        for n in (c.get("documents") or []) + (c.get("systems") or []):
            use.setdefault(str(n).strip(), []).append(("ref", d["doc_id"], p["page"], None, None, ""))
    for name, us in sorted(use.items()):
        ds = sorted({u[1] for u in us})
        if len(ds) >= 2:
            out["shared"].append({"name": name, "docs": ds})
        for o in [u for u in us if u[0] == "out" and u[4]]:
            for i in [u for u in us if u[0] == "in" and u[4] and u[4] != o[4]]:
                if (o[4], i[4]) not in seen:
                    seen.add((o[4], i[4]))
                    out["sequence"].append({"from": o[4], "to": i[4], "via": name, "evidence": o[5],
                                            "sources": [{"doc_id": o[1], "page": o[2]}, {"doc_id": i[1], "page": i[2]}]})
    out["sequence"] += named_refs(work, m, tax, where, seen)
    for i, r in enumerate(out["relations"], 1):                      # 候補の番号（同じ入力からは同じ番号）。採るものを --accept R1,S2 で指す
        r["id"] = f"R{i}"
    for i, s in enumerate(out["sequence"], 1):
        s["id"] = f"S{i}"
    save(work / "structure" / "link_candidates.json", out)
    if a.accept is not None:
        return accept_links(work, out, a.accept)
    print(f"■ 資料どうしの関係の候補（{len(out['relations'])} 件）— 参照の文から")
    for r in out["relations"][:a.show or None]:
        print(f"  {r['id']}  {r['from']} →（{r['type']}）→ {r['to']}｜「{str(r['evidence'])[:50]}」（{lab.src(r['sources'][0])}）")
    print(f"■ 複数の資料に出てくる帳票・システム（{len(out['shared'])} 件）— 同じ業務を別の資料が書いている手がかり")
    for s in out["shared"][:a.show or None]:
        print(f"  {s['name']}｜{'・'.join(s['docs'])}")
    print(f"■ 業務どうしの「続く」の候補（{len(out['sequence'])} 件）— ある業務が作る帳票を別の業務が使っている・別の業務を番号と名前で名指ししている")
    for s in out["sequence"][:a.show or None]:
        print(f"  {s['id']}  {s['from']} → {s['to']}｜{s['via']}｜「{str(s['evidence'])[:40]}」")
    print(f"\n候補 → {work / 'structure' / 'link_candidates.json'}。機械的な当たりなので、根拠の文を読んで採否を決め、採るものの番号だけを "
          f"python3 scripts/pages.py links -w <作業フォルダ> --accept R1,S2 で渡す（根拠の文とページは、機械が hierarchy.json に写す）。根拠の無い紐付けは作らない")


def accept_links(work, out, ids):
    """紐付けの候補のうち、番号で指したものを hierarchy.json（doc_relations・sequence）に写す。根拠の文・ページは候補のまま（LLM は書き写さない）"""
    hp = work / "structure" / "hierarchy.json"
    hier = load(hp)
    if not isinstance(hier, dict):
        raise SystemExit(f"{hp} が無い・読めない（階層を先に書く）")
    by = {x["id"]: (key, x) for key in ("relations", "sequence") for x in out[key]}
    want = [x.strip() for x in str(ids).split(",") if x.strip()]
    miss = [x for x in want if x not in by]
    if miss:
        raise SystemExit(f"候補に無い番号: {'・'.join(miss)}（pages.py links で、いまの候補の番号を見る）")
    added = []
    for x in want:
        key, c = by[x]
        if key == "relations":
            row, rows = {k: c[k] for k in ("from", "to", "type", "evidence", "sources")}, hier.setdefault("doc_relations", [])
            same = any((r.get("from"), r.get("to"), r.get("type")) == (c["from"], c["to"], c["type"]) for r in rows if isinstance(r, dict))
        else:
            row, rows = {k: c[k] for k in ("from", "to", "evidence", "sources")}, hier.setdefault("sequence", [])
            same = any((r.get("from"), r.get("to")) == (c["from"], c["to"]) for r in rows if isinstance(r, dict))
        if not same:
            rows.append(row)
            added.append(f"{x}（{c['from']} → {c['to']}）")
    save(hp, hier)
    print(f"hierarchy.json に写した {len(added)} 件: {'、'.join(added) or 'なし（どれも、すでにある）'}")


def cmd_plan_flows(a, work, m, lab, only):
    tax = load(work / "taxonomy.json")
    if tax is None:
        raise SystemExit("taxonomy.json が無い（先に業務分類を書く）")
    total = big = 0
    for c in tax.get("categories") or []:
        if not c.get("make_flow"):
            continue
        secs, size = [], 0
        for ps in c.get("pages") or []:
            for pg in ps.get("pages") or []:
                f = card_json(work, ps.get("doc_id"), pg)
                n = len(f.read_text(encoding="utf-8")) if f.exists() else 0
                size += n
                card = load(f, {}) or {}
                key = (ps.get("doc_id"), " > ".join((card.get("section_path") or [])[:a.depth + 1]) or "（見出しなし）")
                if not secs or secs[-1]["key"] != key:
                    secs.append({"key": key, "pages": [], "chars": 0})
                secs[-1]["pages"].append(pg)
                secs[-1]["chars"] += n
        total += 1
        over = size > a.cap
        big += over
        est = FIXED_TOKENS + int(size * 1.3)
        print(f"{c['id']} {c.get('name', '')}｜カード {size:,} 字（見込み {est:,} トークン）｜{'★大きすぎる。分ける' if over else 'そのまま1エージェントで'}")
        if over or a.sections:
            for s in secs[:a.show or None]:
                print(f"    {rng(lab, s['key'][0], s['pages'])} {s['key'][1][:40]}｜{s['chars']:,} 字")
            if a.show and len(secs) > a.show:
                print(f"    … ほか {len(secs) - a.show} 節（--show 0 で全部）")
    print(f"\nフローを作る分類 {total}・材料が上限（{a.cap:,} 字）を超える分類 {big}。"
          + ("超える分類は、上の節の切れ目で下位の分類（F03-1、F03-2…）に分け、親の分類には下位フローへのリンクだけを置く。" if big else "")
          + f"分類1つを1つのサブエージェントに渡す")


def cmd_lines(a, work, m, lab, only):
    """機械抽出を、行の番号つきで見せる。カードの手順・分岐・補足は、原文を写さずに、この番号で場所を指す（lines）。
    あわせて、機械で決まることを先に出す: ページの端から拾えた刷られた番号、図形とコネクタの数、字の小ささ"""
    if not only or a.page_set is None:
        raise SystemExit("lines は、資料とページの範囲を決めて使う（--only M01 --pages 13-18）")
    for d, p in all_pages(m, code=True):
        if d["doc_id"] not in only or p["page"] not in a.page_set:
            continue
        pj = load_page(page_json(work, d["doc_id"], p["page"]))
        e = lab.page(d["doc_id"], p["page"])
        notes = [f"刷られた番号 {e['printed']}（ページの端の数字から。画像で確かめる）" if e["how"] == "machine" else "刷られた番号は機械で決まらない（画像で読む）"]
        first = 1
        if pj.get("code"):                                           # プログラムのソース: 番号は、モジュールの行の番号（ページをまたいでも続く。図の出典・ブロックの木と同じ番号）
            c = pj["code"]
            first = c.get("line_from") or 1
            notes = [f"プログラムのソース（{c.get('language')}）「{c.get('module')}」・行の番号は、モジュールの行の番号（このページは {first} 行目から）"]
            if c.get("secrets"):
                notes.append(f"認証情報を伏せた行 {len(c['secrets'])}")
        if pj.get("shapes") or pj.get("connectors"):
            notes.append(f"図形 {len(pj.get('shapes') or [])}・コネクタ {len(pj.get('connectors') or [])}（箱と矢印は、機械抽出の shapes・connectors を写す）")
        if pj.get("small_text"):
            notes.append("字が小さい（crop_page.py で拡大して読む）")
        units = page_units(pj)
        print(f"== {d['doc_id']}/p{p['page']:04d}｜" + "｜".join(notes) + f"｜{len(units)} 行")
        for k, u in enumerate(units, first):
            print(f"{k:>4}| {u}")
        if not units:
            print("     （文字層なし。画像から読み、from_image: true と quote を書く）")


# ---------- 業務分類の下書き・縦軸・目次に無いページ ----------

NUM_HEAD = re.compile(r"^\s*(第\s*[0-9０-９一二三四五六七八九十百]+\s*[章節項編部]|[(（]\s*[0-9０-９]+\s*[)）]|[0-9０-９]+(?:[.．\-][0-9０-９]+)*(?:[.．)）]|(?=\s)))\s*")
LEADER = re.compile(r"[\s.．…・･‥]{2,}[0-9０-９ivxIVX]*\s*$|\s+[0-9０-９]+\s*$")           # 目次の行の、点線とページ番号


def split_number(title):
    """「第3章 受付」→（「第3章」,「受付」）。番号で始まらなければ（""、もとの題名）。「3月分の処理」は番号と見ない"""
    t = str(title or "").strip()
    mt = NUM_HEAD.match(t)
    return (re.sub(r"\s+", "", mt.group(1)), t[mt.end():].strip()) if mt and t[mt.end():].strip() else ("", t)


def same_title(a, b):
    """目次の項目と業務の木の項目が、同じ題名か（片方が、かっこ書きなどを足した長い書き方でもよい）"""
    x, y = norm(split_number(a)[1]), norm(split_number(b)[1])
    return bool(x) and bool(y) and (x == y or x.startswith(y) or y.startswith(x))


def pair_up(nodes, entries):
    """業務の木の項目（上からたどった順）と、目次の項目（書かれた順）を、題名と段で突き合わせる: {木の項目の id: 目次の項目}。
    同じ題名が何度も出る目次（「記録の作成」）でも取り違えないよう、先に合ったものから順に取る。
    段は、上からの相対で比べる（木の根に資料の題名を 1 段目に置き、目次は章が 1 段目、のようにずれていても合う。ずれは、題名の合う組でいちばん多い差）"""
    import collections
    lv = lambda x: x.get("level") if isinstance(x.get("level"), int) and not isinstance(x.get("level"), bool) else None
    diffs = collections.Counter(lv(n) - lv(e) for n in nodes for e in entries if lv(n) is not None and lv(e) is not None and same_title(n.get("title"), e.get("title")))
    shift = diffs.most_common(1)[0][0] if diffs else 0
    out, at = {}, 0
    for n in nodes:
        for k in range(at, len(entries)):
            e = entries[k]
            if same_title(n.get("title"), e.get("title")) and (lv(e) is None or lv(n) is None or lv(e) + shift == lv(n)):
                out[n.get("id")], at = e, k + 1
                break
    return out


def toc_line(work, tdoc, tpages, name, start):
    """目次のページで、その名前が書かれた行: (page, 行の文字, 次に探し始める位置)。見つからなければ (None, "", start)"""
    k = 0
    for pg in tpages:
        for ln in load_page(page_json(work, tdoc, pg)).get("text") or []:
            k += 1
            if k > start and norm(name) and norm(name) in norm(ln):
                return pg, LEADER.sub("", str(ln)).strip() or str(ln).strip(), k
    return None, "", start


def starts_mid_page(work, doc, page, title, body=30):
    """その見出しが、ページの途中から始まっているか（見出しの行より前に、本文らしい長さの行がある）。
    途中から始まるページには、前の項目の続きも載っている"""
    name, before = norm(split_number(title)[1]), 0
    for ln in load_page(page_json(work, doc, page)).get("text") or []:
        if name and name in norm(ln):
            return before >= body
        before += len(norm(ln)) if len(norm(ln)) > 12 else 0          # 柱・ページ番号・日付・章の題などの短い行は数えない
    return False


def draft_taxonomy(work, m, lab, hier):
    """目次に準拠した業務分類の下書き。名称・順序・番号・目次の場所・ページの割当を、階層の判定と目次のカードから機械で組む。
    決めないもの: フローを描くか・対象外か（目次の業務一覧に無い項目は make_flow: false で置く）、目次に無い業務"""
    toc = hier.get("toc_flow_list") or {}
    tdoc = next((x.get("doc_id") for x in toc.get("source") or [] if isinstance(x, dict)), None)
    tpages = [pg for x in toc.get("source") or [] if isinstance(x, dict) and x.get("doc_id") == tdoc for pg in x.get("pages") or []]
    flat = [n for n, _p in tree_nodes(hier.get("tree"))]
    flow_at = pair_up(flat, [e for e in toc.get("entries") or [] if isinstance(e, dict)])
    listed = [t for pg in tpages for t in load_card(work, tdoc, pg).get("toc_entries") or [] if isinstance(t, dict)]
    printed_at = pair_up(flat, listed)                                # 木の項目 → 目次に刷られたページ番号
    seq = [p["page"] for d in m["docs"] if d["doc_id"] == tdoc for p in d.get("page_list") or []]

    def start_of(n):
        e = printed_at.get(n.get("id"))
        if not e or e.get("page") in (None, ""):
            return None
        if lab.docs.get(tdoc, {}).get("basis") == "printed":
            return lab.find(tdoc, e["page"])
        return int(e["page"]) if str(e["page"]).isdigit() and int(e["page"]) in seq else None

    def own_pages(n):
        """その項目のページ。階層の判定に書かれていればそれ。無ければ、目次の番号から（次の項目の前のページまで）"""
        sets = [{"doc_id": r["doc_id"], "pages": list(r["pages"])} for r in n.get("refs") or [] if r.get("doc_id") and r.get("pages")]
        if sets or start_of(n) is None:
            return sets
        inside = {id(x) for x, _p in tree_nodes(n.get("children"))}
        after = [(start_of(x), x) for x in flat[flat.index(n) + 1:] if id(x) not in inside and start_of(x) is not None]
        first, (nxt, nxt_node) = start_of(n), (after[0] if after else (None, None))
        last = max(first, nxt - 1) if nxt is not None else seq[-1]
        if nxt is not None and nxt > first and starts_mid_page(work, tdoc, nxt, nxt_node.get("title")):
            last = nxt                                               # 次の項目がページの途中から始まる。そのページは、両方の分類に入れる
            shared.append((n, nxt))
        return [{"doc_id": tdoc, "pages": [pg for pg in seq if first <= pg <= last and pg not in plain]}]

    plain = {p["page"] for d, p in all_pages(m) if d["doc_id"] == tdoc
             and load_card(work, tdoc, p["page"]).get("page_type") in NON_CONTENT_TYPES}      # 表紙・目次・白紙は、どの分類にも入れない
    cats, pos, shared = [], 0, []
    below = lambda n: any(x.get("id") in flow_at for x, _p in tree_nodes(n.get("children")))
    wide = max(2, len(str(len(hier.get("tree") or []))))

    def walk(children, parent, number):
        nonlocal pos
        for k, n in enumerate(children or [], 1):
            is_flow, has = n.get("id") in flow_at, below(n)
            if parent is not None and not is_flow and not has and parent["make_flow"]:
                continue                                             # フローの中の項目（上位のフローに含む）は、分類にしない
            if parent is not None and not is_flow and not has and parent.get("_out"):
                continue                                             # 対象外の章の下の項目も、分類にしない
            own, name = (str(n["number"]), str(n.get("title", ""))) if n.get("number") else split_number(n.get("title"))
            num = str(flow_at[n["id"]].get("number") or "") if is_flow and flow_at[n["id"]].get("number") else f"{number} {own}".strip()
            cid = f"F{k:0{wide}d}" if parent is None else f"{parent['id']}-{k}"
            pg, text, pos = toc_line(work, tdoc, tpages, name, pos) if tdoc else (None, "", pos)
            c = {"id": cid, "number": num, "name": name, "parent": parent["id"] if parent else None, "hierarchy_id": n.get("id"),
                 "level": (parent["level"] + 1) if parent else 1, "basis": "toc",
                 "toc_ref": {"doc_id": tdoc, "page": pg or (tpages[0] if tpages else 1), "text": text or name},
                 "pages": own_pages(n) if is_flow else [], "make_flow": is_flow, "_out": not is_flow and not has}
            cats.append(c)
            walk(n.get("children"), c, num)
    walk(hier.get("tree"), None, "")
    undecided = [c for c in cats if c.pop("_out")]
    skip = [{"doc_id": d["doc_id"], "page": p["page"], "reason": TYPE_LABEL[t]} for d, p in all_pages(m)
            for t in [load_card(work, d["doc_id"], p["page"]).get("page_type")] if t in NON_CONTENT_TYPES]
    return {"basis": "toc", "rationale": "目次の項目を、名称・順序・番号のまま業務分類にした（pages.py taxonomy --draft による下書き）。",
            "categories": cats, "unassigned_pages": skip, "questions": []}, undecided, [(n.get("title"), pg) for n, pg in shared]


def assigned_pages(tax):
    """(分類に割り当てたページ, 割り当てないと書いたページ)"""
    took = {(ps.get("doc_id"), pg) for c in tax.get("categories") or [] for ps in c.get("pages") or [] for pg in ps.get("pages") or []}
    return took, {(u.get("doc_id"), u.get("page")) for u in tax.get("unassigned_pages") or []}


def cmd_taxonomy(a, work, m, lab, only):
    hier = load(work / "structure" / "hierarchy.json", {}) or {}
    path = work / "taxonomy.json"
    if not a.draft:
        raise SystemExit("taxonomy は --draft を付けて使う（下書きを書く。すでにある分類には触らない）")
    if path.exists():                                                # 分類には触らない。対象外にした分類のページを、割り当てないページに足すだけ
        tax = load(path)
        took, skipped = assigned_pages(tax)
        under = {}
        for n, _p in tree_nodes(hier.get("tree")):
            under[n.get("id")] = [(r.get("doc_id"), pg) for x, _q in tree_nodes([n]) for r in x.get("refs") or [] for pg in r.get("pages") or []]
        added = 0
        for c in tax.get("categories") or []:
            for k in dict.fromkeys(under.get(c.get("hierarchy_id"), []) if str(c.get("out_of_scope") or "").strip() else []):
                if k not in took and k not in skipped:
                    tax.setdefault("unassigned_pages", []).append({"doc_id": k[0], "page": k[1], "reason": c["out_of_scope"]})
                    skipped.add(k)
                    added += 1
        if added:
            save(path, tax)
        print(f"taxonomy.json はすでにある。分類には触っていない。対象外の分類のページを、割り当てないページ（unassigned_pages）に {added} 件足した")
    else:
        if not (hier.get("toc_flow_list") or {}).get("found"):
            raise SystemExit("目次に業務の一覧が無い（hierarchy.json の toc_flow_list.found が false）ので、下書きは作らない。"
                             "LLM が分類して taxonomy.json を書く（references/structure-guide.md「目次に一覧が無いとき」）")
        tax, undecided, shared = draft_taxonomy(work, m, lab, hier)
        save(path, tax)
        flows = [c for c in tax["categories"] if c["make_flow"]]
        print(f"taxonomy.json の下書きを書いた: 分類 {len(tax['categories'])}（フローを描く {len(flows)}・ページ {sum(len(ps['pages']) for c in flows for ps in c['pages'])}）")
        if shared:
            print(f"■ 次の項目がページの途中から始まる境目 {len(shared)} 件 — そのページを前の分類にも入れた（前の項目の続きが載っているかを確かめる）: "
                  + "、".join(f"{t}／{lab.short(tax['categories'][0]['toc_ref']['doc_id'], pg)}" for t, pg in shared[:8]))
        if undecided:
            print(f"■ 目次の業務一覧に無い項目 {len(undecided)} 件 — 業務でなければ out_of_scope に理由を書き、もう一度 --draft（ページを unassigned_pages に足す）。業務なら make_flow を true に")
            for c in undecided[:a.show or None]:
                print(f"  {c['id']} {c['number']} {c['name']}")
    took, skipped = assigned_pages(load(path))
    lost = [(d["doc_id"], p["page"]) for d, p in all_pages(m) if (d["doc_id"], p["page"]) not in took | skipped
            and load_card(work, d["doc_id"], p["page"]).get("page_type") not in NON_CONTENT_TYPES]
    if lost:
        print(f"■ どの分類にも入っていない内容ページ {len(lost)} 件 — python3 scripts/pages.py orphans -w <作業フォルダ> で、まとまりと目次の候補を見る")
    print("このあと: python3 scripts/check_flow.py -w <作業フォルダ> --stage structure（見出しの位置・目次の番号とのずれ・対象外の理由を確かめる）")


def cmd_rows(a, work, m, lab, only):
    """縦軸（業務マトリクスの行）の案。確認点①では階層から組んだ下書きで、確認点②では taxonomy.json で出す"""
    from build_matrix import AXIS_HEAD, Matrix
    hier = load(work / "structure" / "hierarchy.json", {}) or {}
    tax, note = load(work / "taxonomy.json"), "taxonomy.json"
    if tax is None:
        if not (hier.get("toc_flow_list") or {}).get("found"):
            raise SystemExit("taxonomy.json が無く、目次に業務の一覧も無いので、縦軸の案を機械では出せない（LLM が分類してから、もう一度）")
        tax, note = draft_taxonomy(work, m, lab, hier)[0], "階層の判定から組んだ下書き（taxonomy.json はまだ無い）"
    mx = Matrix(work, write=False, tax=tax)
    rows = [c for c in mx.cats if c.get("make_flow")]
    print(f"■ 縦軸（業務マトリクスの行）— {len(rows)} 行｜{note}")
    print("  大分類｜中分類｜分類名の根拠")
    for c in rows[:a.show or None]:
        fl = {"flow_id": c.get("id"), "name": c.get("name"), "basis": c.get("basis")}
        top, mid, _k = mx.names(fl)
        basis = "".join(t for t, _s in mx.name_basis(fl)).replace("\n", "（") + "）"
        print(f"  {top}｜{mid}｜{basis}")
    if len(rows) > (a.show or 10 ** 9):
        print(f"  … ほか {len(rows) - a.show} 行")
    axis = [[c.text for c in r] for r in mx.axis()[1:]]
    count = {}
    for r in axis:
        count[r[3]] = count.get(r[3], 0) + 1
    print(f"■ 業務の木の項目 {len(axis)} 件の扱い: " + "・".join(f"{k} {v}" for k, v in count.items()))
    kind = AXIS_HEAD.index("扱い")
    for r in [r for r in axis if r[kind].startswith(("対象外", "未設定", "目次に無く"))][:a.show or None]:
        print(f"  {r[kind]}｜{r[1].strip()}｜{r[5]}")


def cmd_orphans(a, work, m, lab, only):
    """どの分類にも入っていない内容ページのまとまり（資料 × 見出しの先頭）と、結び付きそうな目次の項目の候補。
    候補は機械的な当たり。結び付けるのは、根拠の文があるとき・見出しが目次の中で一意に一致するときだけ"""
    hier, tax = load(work / "structure" / "hierarchy.json", {}) or {}, load(work / "taxonomy.json", {}) or {}
    took, skipped = assigned_pages(tax)
    if not tax:                                                      # 分類の前は、業務の木に結び付けたページを「入っている」と見る
        took = {(r.get("doc_id"), pg) for n, _p in tree_nodes(hier.get("tree")) for r in n.get("refs") or [] for pg in r.get("pages") or []}
    items = [n for n, _p in tree_nodes(hier.get("tree"))]
    bare = lambda t: norm(split_number(t)[1])
    groups = []
    for d, p in all_pages(m):
        k = (d["doc_id"], p["page"])
        c = load_card(work, *k)
        if k in took or k in skipped or c.get("page_type") in NON_CONTENT_TYPES or (only and d["doc_id"] not in only):
            continue
        head = str((c.get("section_path") or [c.get("title") or ""])[0] or "")
        if not groups or (groups[-1]["doc_id"], groups[-1]["heading"]) != (d["doc_id"], head) or groups[-1]["pages"][-1] != p["page"] - 1:
            groups.append({"doc_id": d["doc_id"], "heading": head, "pages": [], "steps": 0, "candidates": []})
        g = groups[-1]
        g["pages"].append(p["page"])
        g["steps"] += len(c.get("steps") or []) + len(c.get("decisions") or [])
        for x in c.get("cross_refs") or []:                          # このページの参照の文が、目次の項目を名指ししている
            hit = [n for n in items if len(bare(n.get("title"))) >= 3 and bare(n.get("title")) in norm(str(x.get("text", "")) + str(x.get("target", "")))]
            if len({bare(n.get("title")) for n in hit}) == 1 and len(hit) == 1:
                g["candidates"].append({"hierarchy_id": hit[0].get("id"), "title": hit[0].get("title"), "why": "参照の文",
                                        "evidence": str(x.get("text", "")), "sources": [{"doc_id": d["doc_id"], "page": p["page"]}]})
    stems = {d["doc_id"]: norm(file_stem(d["name"])) for d in m["docs"]}
    tied = {}                                                        # ページ → 業務の木の項目（いちばん下の段）
    for n, _p in tree_nodes(hier.get("tree")):
        for r in n.get("refs") or []:
            for pg in r.get("pages") or []:
                tied[(r.get("doc_id"), pg)] = n
    for g in groups:
        same = [n for n in items if bare(n.get("title")) and bare(n.get("title")) == bare(g["heading"])]
        near = [n for n in items if n not in same and len(bare(n.get("title"))) >= 3 and len(bare(g["heading"])) >= 3
                and (bare(n.get("title")) in bare(g["heading"]) or bare(g["heading"]) in bare(n.get("title")))]
        if len(same) == 1:
            g["candidates"].insert(0, {"hierarchy_id": same[0].get("id"), "title": same[0].get("title"), "why": "見出しが目次の中で一意に一致", "evidence": "", "sources": []})
        elif same:
            g["note"] = f"同じ名前の目次の項目が {len(same)} 個ある（{'・'.join(str(n.get('id')) for n in same)}）。見出しだけでは決まらない。根拠の文が要る"
        for n in near[:3]:
            g["candidates"].append({"hierarchy_id": n.get("id"), "title": n.get("title"), "why": "名前が近い（根拠の文が要る）", "evidence": "", "sources": []})
        stem = stems.get(g["doc_id"], "")
        for d, p in all_pages(m):                                    # ほかのページの参照の文が、この資料を名指ししている
            if d["doc_id"] == g["doc_id"] or len(stem) < 3:
                continue
            for x in load_card(work, d["doc_id"], p["page"]).get("cross_refs") or []:
                n = tied.get((d["doc_id"], p["page"]))
                if stem in norm(str(x.get("target", "")) + str(x.get("text", ""))) and n is not None:
                    g["candidates"].append({"hierarchy_id": n.get("id"), "title": n.get("title"), "why": "この資料を名指しした参照の文",
                                            "evidence": str(x.get("text", "")), "sources": [{"doc_id": d["doc_id"], "page": p["page"]}]})
    k = 0
    for g in groups:                                                 # 根拠の文のある候補の番号（--accept C1,C2 で採る）
        for c in g["candidates"]:
            if c["evidence"]:
                k += 1
                c["id"] = f"C{k}"
    save(work / "structure" / "orphans.json", groups)
    if a.accept is not None:
        return accept_orphans(work, hier, tax, groups, a.accept)
    print(f"■ どの分類にも入っていない内容ページのまとまり {len(groups)} 件（{sum(len(g['pages']) for g in groups)} ページ）")
    for g in groups[:a.show or None]:
        print(f"  {g['doc_id']} {rng(lab, g['doc_id'], g['pages'])}「{g['heading'][:30]}」｜手順・分岐 {g['steps']}")
        for c in g["candidates"]:
            print(f"    候補 {c.get('id', '')}{' ' if c.get('id') else ''}{c['hierarchy_id']}「{str(c['title'])[:24]}」— {c['why']}" + (f"「{c['evidence'][:40]}」" if c["evidence"] else ""))
        if g.get("note"):
            print(f"    ※ {g['note']}")
        if not g["candidates"] and not g.get("note"):
            print("    候補なし → 手順・分岐があれば、新しい分類（basis: llm）の候補。無ければ unassigned_pages に理由を書く")
    print("\n見出しが目次の中で一意に一致するまとまりは、python3 scripts/pages.py orphans -w <作業フォルダ> --accept で、機械が業務の木に結び付ける。"
          "根拠の文のある候補は、文を読んで採るものの番号を --accept C1,C2 で渡す（根拠の文とページは、機械が写す）")
    print(f"\n一覧 → {work / 'structure' / 'orphans.json'}。目次の項目に結び付けるのは、根拠の文があるとき・見出しが目次の中で一意に一致するときだけ。"
          f"結び付かない業務は、新しい分類として同じ段の末尾に足し、その分類を指す確認事項（種別「分類」・flow）を書く（structure-guide.md）")


def accept_orphans(work, hier, tax, groups, ids):
    """どの分類にも入っていないまとまりを、業務の木の項目（refs）に結び付ける: 見出しが目次の中で一意に一致する候補はいつも、根拠の文のある候補は番号で指したものだけ"""
    want = {x.strip() for x in str(ids).split(",") if x.strip()}
    known = {c["id"] for g in groups for c in g["candidates"] if c.get("id")}
    if want - known:
        raise SystemExit(f"候補に無い番号: {'・'.join(sorted(want - known))}（pages.py orphans で、いまの候補の番号を見る）")
    node = {n.get("id"): n for n, _p in tree_nodes(hier.get("tree"))}
    done = []
    for g in groups:
        pick = next((c for c in g["candidates"] if c["why"] == "見出しが目次の中で一意に一致"), None) or next((c for c in g["candidates"] if c.get("id") in want), None)
        n = node.get(pick["hierarchy_id"]) if pick else None
        if n is None:
            continue
        refs = n.setdefault("refs", [])
        ref = next((r for r in refs if isinstance(r, dict) and r.get("doc_id") == g["doc_id"]), None)
        if ref is None:
            ref = {"doc_id": g["doc_id"], "pages": [], "rel": "details" if pick["evidence"] else "describes"}
            if pick["evidence"]:
                ref.update(evidence=pick["evidence"], sources=pick["sources"])
            refs.append(ref)
        ref["pages"] = sorted(set(ref.get("pages") or []) | set(g["pages"]))
        cat = next((c for c in tax.get("categories") or [] if isinstance(c, dict) and c.get("hierarchy_id") == n.get("id")), None)
        if cat is not None:                                          # 当たる分類のページにも足す（フローは、目次の項目のフローに入る）
            ps = next((p for p in cat.setdefault("pages", []) if p.get("doc_id") == g["doc_id"]), None)
            if ps is None:
                ps = {"doc_id": g["doc_id"], "pages": []}
                cat["pages"].append(ps)
            ps["pages"] = sorted(set(ps.get("pages") or []) | set(g["pages"]))
        done.append(f"{g['doc_id']} {g['pages'][0]}〜{g['pages'][-1]} → {n.get('id')}「{str(n.get('title'))[:20]}」（{pick['why']}）"
                    + (f"・分類 {cat.get('id')} のページに足した" if cat is not None else ""))
    save(work / "structure" / "hierarchy.json", hier)
    if tax:
        save(work / "taxonomy.json", tax)
    print(f"業務の木に結び付けた {len(done)} 件" + "".join(f"\n  {x}" for x in done))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["batches", "digest", "outline", "pagemap", "terms", "links", "plan-flows", "rows", "taxonomy", "orphans", "lines"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="batches・digest・outline・pagemap・orphans・lines: 資料 ID（カンマ区切り）。ほかのサブコマンドに付けると止める")
    ap.add_argument("--pages", help="batches・digest・lines: ページの範囲（ファイル名の番号。例: 13-24。lines では必須）")
    ap.add_argument("--budget", type=int, default=BUDGET, help=f"batches: 1束の予算（トークン。既定 {BUDGET}）")
    ap.add_argument("--max-pages", type=int, default=MAX_BATCH, help=f"batches: 1束の枚数の上限（既定 {MAX_BATCH}）")
    ap.add_argument("--all", action="store_true", help="batches: カードのあるページも束に入れる（読み直すとき）")
    ap.add_argument("--max-material", "--cap", dest="cap", type=int, default=MATERIAL_CAP,
                    help=f"plan-flows: 材料の上限（カードの文字数。既定 {MATERIAL_CAP}。check_flow.py の --max-material と同じ）")
    ap.add_argument("--draft", action="store_true", help="terms: glossary.json にまだ無い表記を、別々の用語として足す（寄せない。すでにある用語には触らない）／"
                    "taxonomy: 目次に準拠した taxonomy.json の下書きを書く")
    ap.add_argument("--depth", type=int, default=1, help="outline: 見出しを何段まで分けて集計するか／plan-flows: 業務分類の下を、見出しの何段まで分けるか（既定 1 で、見出しの 2 段目まで）")
    ap.add_argument("--sections", action="store_true", help="plan-flows: 上限内の分類も、節ごとの大きさを出す")
    ap.add_argument("--show", type=int, default=40, help="一覧に出す行数の上限（0 で全部）。全件はファイルに書く")
    ap.add_argument("--force", action="store_true", help="digest: 150 ページを超えても全部出す")
    ap.add_argument("--accept", nargs="?", const="", default=None,
                    help="links: 採る候補の番号（R1,S2）を、根拠の文・ページごと hierarchy.json に写す／"
                         "orphans: 見出しが目次の中で一意に一致するまとまりを業務の木に結び付ける（番号 C1,C2 を足すと、根拠の文のある候補も）"),
    a = ap.parse_args()
    work = work_dir(a.work)
    m = manifest(work)
    lab = labels(m)
    only = set(a.only.split(",")) if a.only else None
    if only and a.cmd in ("terms", "links", "plan-flows", "rows", "taxonomy"):
        raise SystemExit(f"{a.cmd} は資料で絞れない（--only を外す。全部の資料をまとめて見る）")
    a.page_set = parse_pages(a.pages)
    {"batches": cmd_batches, "digest": cmd_digest, "outline": cmd_outline, "pagemap": cmd_pagemap, "terms": cmd_terms, "links": cmd_links,
     "plan-flows": cmd_plan_flows, "rows": cmd_rows, "taxonomy": cmd_taxonomy, "orphans": cmd_orphans, "lines": cmd_lines}[a.cmd](a, work, m, lab, only)


if __name__ == "__main__":
    main()
