#!/usr/bin/env python3
"""diff_runs.py — 同じ入力を2回処理した結果を突き合わせ、どこがどれだけ揺れたかを数で出す。

  python3 scripts/diff_runs.py cards <作業フォルダA> <作業フォルダB>                  # 同じページを2回読んだカード
  python3 scripts/diff_runs.py cards <A> <B> --only M01 --pages 13-22
  python3 scripts/diff_runs.py flows <A> <B> --only F03                           # 同じ分類を2回フロー化した結果
  python3 scripts/diff_runs.py activities <A> <B> --only A-01-2                   # 同じプログラムから、アクティビティ図を2回作った結果
  python3 scripts/diff_runs.py structure <A> <B>                                  # 階層・業務分類・用語の正本を2回作った結果
  python3 scripts/diff_runs.py systems <A> <B>                                    # システム台帳・用語の対応・プログラムの木（機械がつなぎ直したあと）を2回作った結果
  python3 scripts/diff_runs.py cards <A> <B> --min 0.9 --report <作業フォルダ>/out/揺れ_読取.md   # 一致率が下回ったら終了コード 1

LLM の読みは、毎回まったく同じにはならない。揺れてよい所（要約の言い回し・ラベルの語尾）と、揺れてはいけない所を分けて測る。
  カード  ページの種別／刷られた番号／見出しの道／手順・分岐・補足の数／手順の対応（引用で突き合わせる）／対応した手順の担当・条件・補足由来の印
  フロー  レーンの名前／ノードの対応（引いたカードの項目で突き合わせる）／対応したノードの種類・担当・推定の印／線と、その推定の印／
          帳票の名前／作業のノードの数／対象外にした項目／ラベル（言い回しは揺れてよい。目安として出す）／確認事項の対象と重要度／
          開始条件・終了条件／対応したノードの動作の種類・量（処理時間の仮の値の鍵）
  アクティビティ図  箱の対応（同じ行を箱にしたか。行の重なりで突き合わせる）／対応した箱の行の範囲・種類・レーン／対象外にした行と理由／
          分岐・繰り返しの扱い（判断にしたか・畳んだか）／ラベル（目安）／確認事項の数
  階層・分類・用語  業務の木の項目／資料どうしの関係／業務どうしの「続く」／業務分類（番号・名前・親・描くか）とページの割当／用語と呼び方
  台帳・用語の対応・木  システムと種類／紐付け／図にしない起動点・手続きと理由／用語の対応の言葉と根拠／木のブロック（手続き・種類・行）
一致率の低い項目は、決まり（references/flow-rules.md）が足りていない所。そこに決め方を1行足して、もう一度測る。
使い方の例: 作業フォルダを写し、片方のカード（またはフロー）を消して同じ指示文で作り直させ、2つを比べる。
"""
import argparse
import pathlib
import sys

from common import load, norm, parse_pages, same_text


def rate(a, b):
    return a / b if b else 1.0


def match(xs, ys, same):
    """xs と ys の対応を、先に合ったものから1対1で取る。[(i, j)]"""
    used, out = set(), []
    for i, x in enumerate(xs):
        for j, y in enumerate(ys):
            if j not in used and same(x, y):
                used.add(j)
                out.append((i, j))
                break
    return out


class Tally:
    def __init__(self):
        self.rows, self.notes = {}, []

    def add(self, name, ok, total=1):
        r = self.rows.setdefault(name, [0, 0])
        r[0] += ok
        r[1] += total

    def note(self, text):
        self.notes.append(text)

    def lines(self, title):
        out = [f"# {title}", "", "| 見るもの | 一致 | 件数 | 一致率 |", "| --- | --- | --- | --- |"]
        out += [f"| {k} | {v[0]} | {v[1]} | {rate(*v):.0%} |" for k, v in self.rows.items()]
        if self.notes:
            out += ["", "## 違っていた所", ""] + [f"- {x}" for x in self.notes]
        return out

    def worst(self):
        return min([rate(*v) for v in self.rows.values()] + [1.0])


def cmp_cards(A, B, only, pages, t):
    files = sorted((A / "cards").glob("*/p*.json"))
    n = 0
    for fa in files:
        doc, no = fa.parent.name, int(fa.stem[1:])
        fb = B / "cards" / doc / fa.name
        if (only and doc not in only) or (pages is not None and no not in pages) or not fb.exists():
            continue
        a, b = load(fa), load(fb)
        n += 1
        where = f"{doc}/{fa.stem}"
        for key, name in (("page_type", "ページの種別"), ("printed_page", "刷られたページ番号"), ("section_path", "見出しの道")):
            same = a.get(key) == b.get(key)
            t.add(name, same)
            if not same:
                t.note(f"{where}: {name}が違う（{a.get(key)} ／ {b.get(key)}）")
        for key, name in (("steps", "手順の数"), ("decisions", "分岐の数"), ("notes", "補足（作業でないもの）の数")):
            na, nb = len(a.get(key) or []), len(b.get(key) or [])
            t.add(name, na == nb)
            if na != nb:
                t.note(f"{where}: {name}が違う（{na} ／ {nb}）")
        sa, sb = a.get("steps") or [], b.get("steps") or []
        def same_place(x, y):
            """同じ箇所を指しているか。両方が行の番号で指していれば、行の重なりで見る（引用の似かよりより確か）"""
            if isinstance(x.get("lines"), list) and isinstance(y.get("lines"), list) and len(x["lines"]) == 2 and len(y["lines"]) == 2:
                return x["lines"][0] <= y["lines"][1] and y["lines"][0] <= x["lines"][1]
            return same_text(x.get("quote", ""), y.get("quote", ""), 0.8)
        pairs = match(sa, sb, lambda x, y: x.get("lines") == y.get("lines") and same_place(x, y)) if all("lines" in x for x in sa + sb) else []
        left_a, left_b = [i for i in range(len(sa)) if i not in {p[0] for p in pairs}], [j for j in range(len(sb)) if j not in {p[1] for p in pairs}]
        pairs += [(left_a[i], left_b[j]) for i, j in match([sa[i] for i in left_a], [sb[j] for j in left_b], same_place)]   # 同じ行のものを先に、残りを重なりで
        t.add("手順の対応（同じ箇所を手順にしたか）", 2 * len(pairs), len(sa) + len(sb))
        for i in sorted(set(range(len(sa))) - {p[0] for p in pairs}):
            t.note(f"{where}: 片方だけにある手順「{str(sa[i].get('action'))[:30]}」（A）")
        for j in sorted(set(range(len(sb))) - {p[1] for p in pairs}):
            t.note(f"{where}: 片方だけにある手順「{str(sb[j].get('action'))[:30]}」（B）")
        for i, j in pairs:
            for key, name in (("actor", "対応した手順の担当"), ("condition", "対応した手順の条件"), ("origin", "対応した手順の補足由来の印")):
                same = norm(str(sa[i].get(key, ""))) == norm(str(sb[j].get(key, "")))
                t.add(name, same)
                if not same:
                    t.note(f"{where}: 「{str(sa[i].get('action'))[:24]}」の{name.replace('対応した手順の', '')}が違う（{sa[i].get(key)} ／ {sb[j].get(key)}）")
        da, db = a.get("decisions") or [], b.get("decisions") or []
        dp = match(da, db, lambda x, y: same_text(x.get("quote", ""), y.get("quote", ""), 0.8))
        t.add("分岐の対応", 2 * len(dp), len(da) + len(db))
    return n


def node_keys(fl):
    """作業のノード → 引いたカードの項目の集まり（無ければラベル）"""
    out = []
    for n in fl.get("nodes") or []:
        if n.get("type") in ("process", "decision", "system"):
            refs = frozenset((s.get("doc_id"), s.get("page"), s.get("item")) for s in n.get("sources") or [] if s.get("item"))
            out.append((n, refs))
    return out


def cmp_flows(A, B, only, t):
    n = 0
    for fa in sorted((A / "flows").glob("*.flow.json")):
        F = fa.name.split(".")[0]
        fb = B / "flows" / fa.name
        if (only and F not in only) or not fb.exists():
            continue
        a, b = load(fa), load(fb)
        n += 1
        la, lb = {norm(l.get("name", "")) for l in a.get("lanes") or []}, {norm(l.get("name", "")) for l in b.get("lanes") or []}
        t.add("レーンの名前", len(la & lb), len(la | lb))
        if la != lb:
            t.note(f"{F}: レーンが違う（A だけ {sorted(la - lb)} ／ B だけ {sorted(lb - la)}）")
        na, nb = node_keys(a), node_keys(b)
        pairs = match(na, nb, lambda x, y: bool(x[1] & y[1]) or (not x[1] and not y[1] and same_text(x[0].get("label", ""), y[0].get("label", ""), 0.8)))
        t.add("ノードの対応（同じ手順をノードにしたか）", 2 * len(pairs), len(na) + len(nb))
        for i in sorted(set(range(len(na))) - {p[0] for p in pairs}):
            t.note(f"{F}: 片方だけにあるノード {na[i][0].get('id')}「{str(na[i][0].get('label'))[:24]}」（A）")
        for j in sorted(set(range(len(nb))) - {p[1] for p in pairs}):
            t.note(f"{F}: 片方だけにあるノード {nb[j][0].get('id')}「{str(nb[j][0].get('label'))[:24]}」（B）")
        lane_a = {l["id"]: norm(l.get("name", "")) for l in a.get("lanes") or []}
        lane_b = {l["id"]: norm(l.get("name", "")) for l in b.get("lanes") or []}
        ida, idb = {}, {}
        for k, (i, j) in enumerate(pairs):
            x, y = na[i][0], nb[j][0]
            ida[x["id"]], idb[y["id"]] = k, k
            for name, same in (("対応したノードの種類", x.get("type") == y.get("type")),
                               ("対応したノードの担当", lane_a.get(x.get("lane")) == lane_b.get(y.get("lane"))),
                               ("対応したノードの推定の印", bool(x.get("inferred")) == bool(y.get("inferred"))),
                               ("対応したノードの動作の種類", x.get("action") == y.get("action")),
                               ("対応したノードの量", x.get("volume") == y.get("volume"))):
                t.add(name, same)
                if not same:
                    t.note(f"{F}: {x.get('id')}「{str(x.get('label'))[:20]}」の{name.replace('対応したノードの', '')}が違う")
            alike = norm(x.get("label", "")) == norm(y.get("label", ""))
            t.add("対応したノードのラベル（一字一句同じ）", alike)
            t.add("対応したノードのラベル（8割以上同じ）", alike or same_text(x.get("label", ""), y.get("label", ""), 0.8))
            if not alike:
                t.note(f"{F}: ラベルが違う「{str(x.get('label'))[:24]}」／「{str(y.get('label'))[:24]}」")
        ea = {(ida[e["from"]], ida[e["to"]]): bool(e.get("inferred")) for e in a.get("edges") or [] if e.get("from") in ida and e.get("to") in ida}
        eb = {(idb[e["from"]], idb[e["to"]]): bool(e.get("inferred")) for e in b.get("edges") or [] if e.get("from") in idb and e.get("to") in idb}
        t.add("線（対応したノードどうし）", len(set(ea) & set(eb)), len(set(ea) | set(eb)))
        t.add("両方にある線の推定の印", sum(ea[k] == eb[k] for k in set(ea) & set(eb)), len(set(ea) & set(eb)))
        for k in sorted(set(ea) ^ set(eb)):
            x, y = (na[pairs[k[0]][0]][0], na[pairs[k[1]][0]][0])
            t.note(f"{F}: 片方だけにある線「{str(x.get('label'))[:16]}」→「{str(y.get('label'))[:16]}」（{'A' if k in ea else 'B'}）")
        fa_, fb_ = ({norm(n.get("label", "")) for n in fl.get("nodes") or [] if n.get("type") == "file"} for fl in (a, b))
        t.add("帳票（ファイルのノード）の名前", len(fa_ & fb_), len(fa_ | fb_))
        sa_, sb_ = (sum(n.get("type") in ("process", "decision", "system", "subprocess") for n in fl.get("nodes") or []) for fl in (a, b))
        t.add("作業のノードの数", sa_ == sb_)
        if sa_ != sb_:
            t.note(f"{F}: 作業のノードの数が違う（{sa_} ／ {sb_}）")
        xa = {(x.get("doc_id"), x.get("page"), x.get("item")) for x in a.get("excluded_items") or []}
        xb = {(x.get("doc_id"), x.get("page"), x.get("item")) for x in b.get("excluded_items") or []}
        t.add("対象外にした項目", len(xa & xb), len(xa | xb))
        def asked(fl, idx, tag):
            """確認事項の対象（対応したノードの番号に直したもの）→ 重要度。開始・終了・帳票・別のフローへの箱は、種類とラベルで見る。対応の無いノードは、その回だけのもの"""
            ends = {n["id"]: (n.get("type"), norm(n.get("label", ""))) for n in fl.get("nodes") or [] if n.get("type") in ("start", "end", "file", "subprocess")}
            key = lambda i: idx.get(i, ends.get(i, f"{tag}:{i}"))
            out = {}
            for q in fl.get("questions") or []:
                ids = ([q["node"]] if q.get("node") else []) + list(q.get("nodes") or [])
                k = ("線",) + tuple(key(i) for i in q["edge"]) if q.get("edge") else ("ノード", frozenset(key(i) for i in ids)) if ids else ("全体", norm(q.get("question", ""))[:12])
                out[k] = q.get("severity")
            return out
        sa_q, sb_q = asked(a, ida, "A"), asked(b, idb, "B")
        ta, tb = set(sa_q), set(sb_q)
        t.add("確認事項の対象（同じノード・線に問いがあるか）", len(ta & tb), len(ta | tb))
        both = ta & tb
        t.add("両方にある確認事項の重要度", sum(sa_q[k] == sb_q[k] for k in both), len(both))
        for key, name in (("start_condition", "開始条件"), ("end_condition", "終了条件")):
            same = norm(str(a.get(key) or "")) == norm(str(b.get(key) or ""))
            t.add(name, same)
            if not same:
                t.note(f"{F}: {name}が違う（{str(a.get(key))[:30]} ／ {str(b.get(key))[:30]}）")
        qa, qb = len(a.get("questions") or []), len(b.get("questions") or [])
        t.add("確認事項の数", qa == qb)
        if qa != qb:
            t.note(f"{F}: 確認事項の数が違う（{qa} ／ {qb}）")
    return n


def cmp_activities(A, B, only, t):
    """アクティビティ図: 箱は、受け持つソースの行（モジュール・行）の重なりで突き合わせる（半分以上重なれば、同じ所を箱にしたと見る）"""
    n = 0
    for fa in sorted((A / "activities").glob("*.activity.json")):
        F = fa.name.split(".")[0]
        fb = B / "activities" / fa.name
        if (only and F not in only) or not fb.exists():
            continue
        a, b = load(fa), load(fb)
        n += 1

        def boxes(act):
            own = {}
            for o in act.get("lines") or []:
                if o.get("state") == "box" and o.get("node"):
                    own.setdefault(o["node"], set()).update((o["module"], ln) for ln in o["lines"])
            by = {x["id"]: x for x in act.get("nodes") or []}
            return [(by[i], lines) for i, lines in own.items() if i in by]
        na, nb = boxes(a), boxes(b)
        pairs = match(na, nb, lambda x, y: len(x[1] & y[1]) * 2 >= len(x[1] | y[1]))
        t.add("箱の対応（同じ行を箱にしたか）", 2 * len(pairs), len(na) + len(nb))
        for i in sorted(set(range(len(na))) - {p[0] for p in pairs}):
            t.note(f"{F}: 片方だけにある箱 {na[i][0]['id']}「{str(na[i][0].get('label'))[:24]}」（A）")
        for j in sorted(set(range(len(nb))) - {p[1] for p in pairs}):
            t.note(f"{F}: 片方だけにある箱 {nb[j][0]['id']}「{str(nb[j][0].get('label'))[:24]}」（B）")
        lane_a, lane_b = ({l["id"]: l.get("kind") for l in act.get("lanes") or []} for act in (a, b))
        for i, j in pairs:
            (x, lx), (y, ly) = na[i], nb[j]
            for name, same in (("対応した箱の行の範囲（まとめ方・分け方）", lx == ly), ("対応した箱の種類", x.get("type") == y.get("type")),
                               ("対応した箱のレーン", lane_a.get(x.get("lane")) == lane_b.get(y.get("lane")))):
                t.add(name, same)
                if not same:
                    t.note(f"{F}: {x['id']}「{str(x.get('label'))[:20]}」の{name.replace('対応した箱の', '')}が違う")
            alike = norm(x.get("label", "")) == norm(y.get("label", ""))
            t.add("対応した箱のラベル（一字一句同じ）", alike)
            t.add("対応した箱のラベル（8割以上同じ）", alike or same_text(x.get("label", ""), y.get("label", ""), 0.8))
            if not alike:
                t.note(f"{F}: ラベルが違う「{str(x.get('label'))[:24]}」／「{str(y.get('label'))[:24]}」")
        xa, xb = ({(o["module"], ln): o.get("reason", "") for o in act.get("lines") or [] if o.get("state") == "excluded" for ln in o["lines"]} for act in (a, b))
        t.add("対象外にした行", len(set(xa) & set(xb)), len(set(xa) | set(xb)))
        t.add("両方で対象外にした行の理由", sum(xa[k] == xb[k] for k in set(xa) & set(xb)), len(set(xa) & set(xb)))
        if set(xa) ^ set(xb):
            t.note(f"{F}: 片方だけが対象外にした行 {len(set(xa) ^ set(xb))}（{', '.join(f'{m} {ln}' for m, ln in sorted(set(xa) ^ set(xb))[:6])}）")
        ba, bb = ({x["key"]: x.get("state") for x in act.get("branches") or []} for act in (a, b))
        t.add("分岐・繰り返しの扱い（判断にしたか・畳んだか）", sum(ba[k] == bb.get(k) for k in ba), len(set(ba) | set(bb)))
        for k in sorted(set(ba) | set(bb)):
            if ba.get(k) != bb.get(k):
                t.note(f"{F}: {k} の扱いが違う（{ba.get(k, 'なし')} ／ {bb.get(k, 'なし')}）")
        qa, qb = len(a.get("questions") or []), len(b.get("questions") or [])
        t.add("確認事項の数", qa == qb)
        if qa != qb:
            t.note(f"{F}: 確認事項の数が違う（{qa} ／ {qb}）")
    return n


def overlap(t, name, xa, xb, where=""):
    """集まりどうしの一致を数え、違いを書き留める"""
    t.add(name, len(xa & xb), len(xa | xb))
    if xa != xb:
        t.note(f"{where}{name}: A だけ {sorted(map(str, xa - xb))[:4]} ／ B だけ {sorted(map(str, xb - xa))[:4]}")


def cmp_structure(A, B, only, t):
    """階層（業務の木・資料どうしの関係・業務どうしの「続く」）・業務分類・用語の正本"""
    n = 0
    ha, hb = (load(x / "structure" / "hierarchy.json", None) for x in (A, B))
    if isinstance(ha, dict) and isinstance(hb, dict):
        n += 1
        from common import tree_nodes
        items = lambda h: {(x.get("id"), norm(x.get("title", ""))) for x, _p in tree_nodes(h.get("tree"))}
        overlap(t, "業務の木の項目（番号・名前）", items(ha), items(hb))
        overlap(t, "資料どうしの関係", {(r.get("from"), r.get("to"), r.get("type")) for r in ha.get("doc_relations") or []},
                {(r.get("from"), r.get("to"), r.get("type")) for r in hb.get("doc_relations") or []})
        overlap(t, "業務どうしの「続く」", {(r.get("from"), r.get("to")) for r in ha.get("sequence") or []}, {(r.get("from"), r.get("to")) for r in hb.get("sequence") or []})
    ta, tb = (load(x / "taxonomy.json", None) for x in (A, B))
    if isinstance(ta, dict) and isinstance(tb, dict):
        n += 1
        cat = lambda tx: {(c.get("id"), norm(c.get("name", "")), c.get("parent"), bool(c.get("make_flow"))) for c in tx.get("categories") or [] if not only or c.get("id") in only}
        overlap(t, "業務分類（番号・名前・親・描くか）", cat(ta), cat(tb))
        pages = lambda tx: {(c.get("id"), p.get("doc_id"), pg) for c in tx.get("categories") or [] if not only or c.get("id") in only
                            for p in c.get("pages") or [] for pg in p.get("pages") or []}
        overlap(t, "ページの割当（分類 × ページ）", pages(ta), pages(tb))
    ga, gb = (load(x / "glossary.json", None) for x in (A, B))
    if isinstance(ga, dict) and isinstance(gb, dict):
        n += 1
        overlap(t, "用語（種類・名前）", {(x.get("kind"), norm(x.get("name", ""))) for x in ga.get("terms") or []}, {(x.get("kind"), norm(x.get("name", ""))) for x in gb.get("terms") or []})
        alias = lambda g: {(norm(x.get("name", "")), norm(y)) for x in g.get("terms") or [] for y in x.get("aliases") or []}
        overlap(t, "呼び方（寄せた先）", alias(ga), alias(gb))
    return n


def cmp_systems(A, B, only, t):
    """システム台帳・用語の対応・プログラムの木（code_struct.py がつなぎ直したあとのブロック）"""
    n = 0
    sa, sb = (load(x / "structure" / "systems.json", None) for x in (A, B))
    if isinstance(sa, dict) and isinstance(sb, dict):
        n += 1
        overlap(t, "システム（名前・種類）", {(norm(s.get("name", "")), s.get("kind")) for s in sa.get("systems") or []}, {(norm(s.get("name", "")), s.get("kind")) for s in sb.get("systems") or []})
        overlap(t, "紐付け（起動点 × 手順）", {(x.get("entry"), x.get("step")) for x in sa.get("links") or []}, {(x.get("entry"), x.get("step")) for x in sb.get("links") or []})
        overlap(t, "図にしない起動点と理由", {(x.get("entry"), x.get("reason")) for x in sa.get("excluded_entries") or []},
                {(x.get("entry"), x.get("reason")) for x in sb.get("excluded_entries") or []})
        overlap(t, "図にしない手続きと理由", {(x.get("procedure"), x.get("reason")) for x in sa.get("folded_procedures") or []},
                {(x.get("procedure"), x.get("reason")) for x in sb.get("folded_procedures") or []})
    ca, cb = (load(x / "structure" / "code_terms.json", None) for x in (A, B))
    if isinstance(ca, dict) and isinstance(cb, dict):
        n += 1
        ra, rb = ({(r.get("project"), r.get("name")): r for r in c.get("terms") or []} for c in (ca, cb))
        both = set(ra) & set(rb)
        t.add("用語の対応（同じ名前を載せたか）", len(both), len(set(ra) | set(rb)))
        t.add("両方にある名前の言葉", sum(norm(ra[k].get("word", "")) == norm(rb[k].get("word", "")) for k in both), len(both))
        t.add("両方にある名前の根拠", sum(ra[k].get("basis") == rb[k].get("basis") for k in both), len(both))
    for fa in sorted((A / "code").glob("*.struct.json")) if (A / "code").exists() else []:
        fb = B / "code" / fa.name
        if (only and fa.name.split(".")[0] not in only) or not fb.exists():
            continue
        n += 1

        def blocks(st):
            out = set()

            def walk(pid, nodes):
                for x in nodes or []:
                    out.add((pid, x.get("kind"), tuple(x.get("lines") or [])))
                    for b in x.get("branches") or []:
                        walk(pid, b.get("body"))
                    walk(pid, x.get("body"))
            for p in st.get("procedures") or []:
                walk(p.get("id"), p.get("body"))
            return out
        overlap(t, "木のブロック（手続き・種類・行）", blocks(load(fa, {}) or {}), blocks(load(fb, {}) or {}), f"{fa.name.split('.')[0]}: ")
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("kind", choices=["cards", "flows", "activities", "structure", "systems"])
    ap.add_argument("a")
    ap.add_argument("b")
    ap.add_argument("--only", help="資料 ID（cards）・フロー ID（flows）・アクティビティ図の番号（activities）。カンマ区切り")
    ap.add_argument("--pages", help="cards: ページの範囲（ファイル名の番号）")
    ap.add_argument("--min", type=float, help="どれかの一致率がこれを下回ったら終了コード 1（例: 0.9）")
    ap.add_argument("--report")
    a = ap.parse_args()
    A, B = pathlib.Path(a.a).expanduser().resolve(), pathlib.Path(a.b).expanduser().resolve()
    only = set(a.only.split(",")) if a.only else None
    t = Tally()
    n = cmp_cards(A, B, only, parse_pages(a.pages), t) if a.kind == "cards" else \
        {"flows": cmp_flows, "activities": cmp_activities, "structure": cmp_structure, "systems": cmp_systems}[a.kind](A, B, only, t)
    if not n:
        sys.exit("比べるものが無い（両方の作業フォルダに同じカード・フロー・アクティビティ図があること）")
    title = f"2回の結果の揺れ（{ {'cards': 'カード', 'flows': 'フロー', 'activities': 'アクティビティ図', 'structure': '階層・分類・用語', 'systems': '台帳・用語の対応・木'}[a.kind] } " \
            f"{n} {'本' if a.kind == 'flows' else 'ファイル' if a.kind in ('structure', 'systems') else '枚'}）"
    lines = t.lines(title)
    shown = lines if len(t.notes) <= 30 else lines[:lines.index("## 違っていた所") + 32] + [f"- … ほか {len(t.notes) - 30} 件（--report のファイルに全件）"]
    print("\n".join(shown))
    if a.report:
        p = pathlib.Path(a.report)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("\n".join(lines) + "\n", encoding="utf-8")
    low = t.worst()
    print(f"\nいちばん低い一致率 {low:.0%}" + (f"（下限 {a.min:.0%}）" if a.min is not None else ""))
    sys.exit(1 if a.min is not None and low < a.min else 0)


if __name__ == "__main__":
    main()
