#!/usr/bin/env python3
"""sc_check.py — 拠点間比較の機械チェック。FAIL があれば終了コード 1。

  python3 scripts/sc_check.py -w <比較の作業フォルダ>                      # できている段を全部
  python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage sites        # 段を 1 つ
  python3 scripts/sc_check.py -w <比較の作業フォルダ> --fast               # 拠点の門（asis の点検を回す）を飛ばす
  python3 scripts/sc_check.py -w <比較の作業フォルダ> --report out/チェック_比較.txt   # 全行をファイルに（画面は同じ種類を 10 件まで）

[sites]     sites.json の形・拠点 ID と別名の重なり・拠点の asis の作業フォルダ・拠点の門（asis の check_flow --stage flow／review が FAIL 0）・
            共通の資料（読取が済んでいる・拠点の棚卸しに入っている・カードが共通のものと同じ・共通の棚卸しが古くない）
[applies]   書き分け（applies.json）: 拾った項目が全部ある・未定が無い・only／except の拠点と根拠・資料が変わっていない。
            拠点ごとに: ほかの拠点だけの手順がフローに入っていない・外す理由が決まった形・その拠点の手順を外していない・拠点で分かれる判断を置いていない
[snapshot]  拠点の写しがある・形・写したあとに拠点の作業フォルダが変わっていない（変わっていれば sc_collect.py）
[terms]     用語の対応（terms.json）: 比べる箱に出る用語が、ちょうど 1 つの組にある・組の ID と種類・same_name は名前が同じ・訳と言い換えは理由・
            同じ拠点の 2 つの用語を 1 つの組にしない・担当と帳票の組の名前は日本語
[axis]      比べる行（axis.json）: 基準の有無と mode・行と業務分類の ID・全拠点が 1 回ずつ・当て方と箱の有無・理由・フローと箱が写しにある・
            名前は日本語（日本語でない拠点は、フローの名前の日本語 title）・作業の箱が、ちょうど 1 つの行か not_compared に入る・origin
[rows]      タスクの対応（rows/<行>.match.json）: 行ごとにある・拠点の欄・箱がその行の作業の箱・箱がちょうど 1 つのタスクに入る・名前は日本語の動詞で終える・
            粒度の決まり（どの拠点も 2 つ以上の箱なら joined に理由）・記載なしは機械の候補を全部確かめる・内容の差の拠点と根拠の箱
[compute]   計算の結果（rows/<行>.result.json）が、いまのデータから計算し直したものと同じ（違えば sc_compute.py）
"""
import argparse
import pathlib
import re
import sys

from sc_common import (OTHER_SITE, STEP_TYPES, SWITCH_WORDS, A, Report, ends_with_verb, is_jp, load, match_path, mentions, norm, result_path,
                       shared_work, site_ids, site_names, site_work, sites_of, snap_path, split_key, try_load, work_dir)
from sc_validate import validate

STAGES = ["sites", "applies", "snapshot", "terms", "axis", "rows", "compute"]


class Ctx:
    def __init__(self, cmp, fast):
        self.cmp, self.fast = cmp, fast
        self.sites = sites_of(cmp)
        self.ids = site_ids(self.sites)
        self.names = site_names(self.sites)
        self._snaps, self._fresh = {}, {}

    def site(self, sid):
        return next(x for x in self.sites["sites"] if x["id"] == sid)

    def work(self, sid):
        return site_work(self.cmp, self.sites, sid)

    def snap(self, sid):
        if sid not in self._snaps:
            s, err = try_load(snap_path(self.cmp, sid))
            self._snaps[sid] = s if isinstance(s, dict) else None
        return self._snaps[sid]

    def uses_shared(self, sid):
        return bool(self.sites.get("shared")) and self.site(sid).get("uses_shared", True)


def shape(rp, data, kind, where):
    errs = validate(data, kind)
    for e in errs[:10]:
        rp.fail(f"{where}: 形が違う — {e}")
    if len(errs) > 10:
        rp.fail(f"{where}: 形の違反が、ほかに {len(errs) - 10} 件")
    return not errs


# ---------- [sites] ----------

def check_sites(cx, rp):
    s = cx.sites
    if not shape(rp, s, "sites", "sites.json"):
        return
    if len(set(cx.ids)) != len(cx.ids):
        rp.fail("sites.json: 拠点 ID が重なっている")
    if s.get("baseline") is not None and s["baseline"] not in cx.ids:
        rp.fail(f"sites.json: 基準の拠点 {s['baseline']} が、拠点の一覧に無い")
    seen = {}
    for x in s["sites"]:
        for n in [x["name"]] + list(x.get("aliases") or []):
            k = norm(n)
            if k in seen and seen[k] != x["id"]:
                rp.fail(f"sites.json: 「{n}」が {cx.names[seen[k]]} と {x['name']} の両方の名前・別名になっている（書き分けを拾えない）")
            seen[k] = x["id"]
    for sid in cx.ids:
        w = cx.work(sid)
        if not (w / "manifest.json").exists():
            rp.fail(f"{cx.names[sid]}: asis の作業フォルダができていない（{w}。references/site-guide.md の手順で、棚卸しから原典照合まで）")
            continue
        if not cx.fast:
            import sc_collect
            for name, n, why in sc_collect.gate(w):
                if n:
                    rp.fail(f"{cx.names[sid]}: 拠点の門を通らない — asis の {name} が {n} FAIL" + (f"〔{why}〕" if why else "")
                            + "（拠点の as-is を原典照合まで済ませる）")
    check_shared(cx, rp)


def same_card(a, b):
    a, b = dict(a), dict(b)
    a.pop("doc_id", None)
    b.pop("doc_id", None)
    return a == b


def check_shared(cx, rp):
    sh = shared_work(cx.cmp, cx.sites)
    if sh is None:
        for sid in cx.ids:
            if cx.site(sid).get("uses_shared") is False:
                rp.warn(f"{cx.names[sid]}: 共通の資料が無いのに uses_shared: false がある（消してよい）")
        return
    m = load(sh / "manifest.json", None)
    if not m:
        rp.fail(f"共通の資料の棚卸しが無い（python3 <asis>/scripts/inventory.py <共通の資料> -w {sh}）")
        return
    docs = {d["sha1"]: d for d in m["docs"] if not d.get("duplicate_of")}
    for d in docs.values():
        if not d.get("rendered"):
            rp.fail(f"共通の資料「{d['name']}」をページ化していない（python3 <asis>/scripts/render_pages.py -w {sh}）")
            continue
        lack = [p["page"] for p in d.get("page_list") or [] if not A.card_json(sh, d["doc_id"], p["page"]).exists()]
        if lack:
            rp.fail(f"共通の資料「{d['name']}」の読取が終わっていない（カードの無いページ {len(lack)}）。共通の作業フォルダで読取を終える")
    roots = [pathlib.Path(x).expanduser().resolve() for x in (cx.sites["shared"].get("inputs") or [])]
    for sid in cx.ids:
        w = cx.work(sid)
        sm = load(w / "manifest.json", None)
        if not sm:
            continue
        mine = {d["sha1"]: d for d in sm["docs"] if not d.get("duplicate_of")}
        if not cx.uses_shared(sid):
            continue
        for sha, d in docs.items():
            if sha not in mine:
                rp.fail(f"{cx.names[sid]}: 共通の資料「{d['name']}」が、拠点の棚卸しに入っていない（inventory.py の入力に、共通の資料のフォルダを入れる）")
                continue
            md = mine[sha]
            for p in d.get("page_list") or []:
                a = load(A.card_json(sh, d["doc_id"], p["page"]), None)
                b = load(A.card_json(w, md["doc_id"], p["page"]), None)
                if a is not None and (b is None or not same_card(a, b)):
                    rp.fail(f"{cx.names[sid]}: 共通の資料「{d['name']}」p.{p['page']} のカードが、共通の作業フォルダのものと違う"
                            f"（共通の資料のカードは、共通の作業フォルダで直す。python3 scripts/sc_share.py -w <比較の作業フォルダ> --site {sid} で写し直す）")
        for d in sm["docs"]:
            path = pathlib.Path(d.get("abs_path", ""))
            if any(r == path or r in path.parents for r in roots) and d.get("sha1") not in docs:
                rp.fail(f"{cx.names[sid]}: 共通の資料「{d['name']}」の中身が、共通の作業フォルダの棚卸しと違う（共通の資料が変わった。"
                        f"共通の作業フォルダの棚卸し・読取をやり直してから、sc_share.py で写す）")


# ---------- [applies] ----------

def check_applies(cx, rp):
    import sc_applies
    if not cx.sites.get("shared"):
        return
    cand = sc_applies.candidates(cx.cmp, cx.sites)
    ap, err = try_load(cx.cmp / "applies.json")
    if err:
        rp.fail(f"applies.json が読めない（{err}）")
        return
    if ap is None:
        if cand:
            rp.fail(f"共通の資料に、拠点の書き分けの候補が {len(cand)} 件あるのに applies.json が無い（python3 scripts/sc_applies.py draft）")
        return
    if not shape(rp, ap, "applies", "applies.json"):
        return
    have = {it["key"]: it for it in ap["items"]}
    for c in cand:
        it = have.get(c["key"])
        if it is None:
            rp.fail(f"applies.json: 拾った項目 {c['key']}「{c['text'][:24]}」が無い（python3 scripts/sc_applies.py draft で足す）")
        elif it["doc"] != c["doc"]:
            rp.fail(f"applies.json {c['key']}: 共通の資料の中身が変わった（sc_applies.py draft をやり直し、決め直す）")
    keys = {c["key"] for c in cand}
    for it in ap["items"]:
        k = it["key"]
        if k not in keys:
            rp.warn(f"applies.json {k}: いまの共通の資料から拾われない項目（sc_applies.py draft で外れる）")
        if it["scope"] == "未定":
            rp.fail(f"applies.json {k}「{it['text'][:24]}」: 当てはまる拠点が未定（all・only・except を決める）")
        elif it["scope"] in ("only", "except") and not it["sites"]:
            rp.fail(f"applies.json {k}: {it['scope']} なのに拠点（sites）が無い")
        elif it["scope"] == "all" and it["sites"]:
            rp.fail(f"applies.json {k}: all なのに拠点（sites）が書いてある（only か except にする）")
        if it["scope"] in ("only", "except") and not it["evidence"].strip():
            rp.fail(f"applies.json {k}: 拠点を絞った根拠（evidence）が無い（原文の言い回しを書く）")
        for s in it["sites"]:
            if s not in cx.ids:
                rp.fail(f"applies.json {k}: 拠点 {s} が sites.json に無い")
    for sid in cx.ids:
        if cx.uses_shared(sid):
            check_site_scope(cx, rp, sid, ap)
        check_switch_decisions(cx, rp, sid)


def check_site_scope(cx, rp, sid, ap):
    import sc_applies
    w = cx.work(sid)
    m = load(w / "manifest.json", None)
    if not m:
        return
    by_sha = {d["sha1"]: d["doc_id"] for d in m["docs"] if not d.get("duplicate_of")}
    cited, excl = {}, {}
    for fl in A.load_flows(w):
        for n in fl["nodes"]:
            for s in n.get("sources") or []:
                if s.get("item"):
                    cited.setdefault(f"{s['doc_id']}.{s['page']}.{s['item']}", []).append(f"{fl['flow_id']} {n['id']}")
        for x in fl.get("excluded_items") or []:
            excl[f"{x['doc_id']}.{x['page']}.{x['item']}"] = x.get("reason", "")
    who = cx.names[sid]
    for it in ap["items"]:
        ok = sc_applies.applies_to(it, sid)
        if ok is None or it["doc"] not in by_sha:
            continue
        _d, page, item = it["key"].split(".")
        key = f"{by_sha[it['doc']]}.{page}.{item}"
        if not ok:
            if key in cited:
                rp.fail(f"{who} {cited[key][0]}: ほかの拠点だけの手順「{it['text'][:24]}」が、この拠点のフローに入っている（applies.json の決め。"
                        f"python3 scripts/sc_site_ops.py --site {sid} の直しを貼る）")
            elif key in excl and not excl[key].startswith(OTHER_SITE):
                rp.fail(f"{who}: ほかの拠点だけの手順 {key}「{it['text'][:24]}」を外す理由が、決まった形でない（「{excl[key][:20]}」。sc_site_ops.py の理由を使う）")
        elif excl.get(key, "").startswith(OTHER_SITE):
            rp.fail(f"{who}: この拠点に当てはまる手順 {key}「{it['text'][:24]}」を、ほかの拠点だけの手順として外している（applies.json の決めと合わない）")


def check_switch_decisions(cx, rp, sid):
    w = cx.work(sid)
    if not (w / "manifest.json").exists():
        return
    for fl in A.load_flows(w):
        for n in fl["nodes"]:
            if n["type"] != "decision":
                continue
            lab = n.get("label", "")
            if mentions(lab, cx.sites) or any(norm(x) in norm(lab) for x in SWITCH_WORDS):
                rp.fail(f"{cx.names[sid]} {fl['flow_id']} {n['id']}「{lab[:24]}」: 拠点で分かれる判断を置いている（拠点のフローには、その拠点の側だけを描く）")


# ---------- [snapshot] ----------

def check_snapshot(cx, rp):
    import sc_collect
    for sid in cx.ids:
        who = cx.names[sid]
        s = cx.snap(sid)
        if s is None:
            rp.fail(f"{who}: 拠点の写しが無い（python3 scripts/sc_collect.py -w <比較の作業フォルダ> --site {sid}）")
            continue
        if s.get("site") != sid or not isinstance(s.get("nodes"), dict) or not isinstance(s.get("families"), dict):
            rp.fail(f"{who}: 拠点の写しの形が崩れている（sc_collect.py で取り直す。手で直さない）")
            continue
        w = cx.work(sid)
        if not (w / "manifest.json").exists():
            continue
        fresh = sc_collect.collect(cx.cmp, cx.sites, sid)
        if fresh["stamp"] != s.get("stamp"):
            rp.fail(f"{who}: 拠点の作業フォルダが、写したあとに変わった（python3 scripts/sc_collect.py -w <比較の作業フォルダ> --site {sid}。"
                    f"照合し直す所は、点検が挙げる）")


# ---------- [terms] ----------

KIND_PREFIX = {"actor": "R", "system": "Y", "document": "D"}


def check_terms(cx, rp):
    import sc_terms
    if not any(cx.snap(s) for s in cx.ids):
        return
    tm, err = try_load(cx.cmp / "terms.json")
    if err or tm is None:
        rp.fail(f"terms.json が{'読めない（' + err + '）' if err else '無い'}（python3 scripts/sc_terms.py draft で下書き）")
        return
    if not shape(rp, tm, "terms", "terms.json"):
        return
    ids, where = set(), {}
    for g in tm["groups"]:
        gid = g["id"]
        if gid in ids:
            rp.fail(f"terms.json: 組の ID {gid} が重なっている")
        ids.add(gid)
        if not gid.startswith(KIND_PREFIX[g["kind"]]):
            rp.fail(f"terms.json {gid}: ID の頭が種類と合わない（{g['kind']} は {KIND_PREFIX[g['kind']]}）")
        if g["kind"] in ("actor", "document") and not is_jp(g["name"]):
            rp.fail(f"terms.json {gid}「{g['name']}」: 担当・帳票の組の名前は日本語にする（納品物に出る）")
        per = {}
        for m in g["members"]:
            if m["site"] not in cx.ids:
                rp.fail(f"terms.json {gid}: 拠点 {m['site']} が sites.json に無い")
            per.setdefault(m["site"], []).append(m["term"])
            k = (m["site"], g["kind"], m["term"])
            if k in where:
                rp.fail(f"terms.json: {cx.names.get(m['site'], m['site'])}の「{m['term']}」が {where[k]} と {gid} の 2 つの組に入っている")
            where[k] = gid
        for s, ts in per.items():
            if len(ts) > 1:
                rp.fail(f"terms.json {gid}: {cx.names.get(s, s)}の 2 つの用語（{'・'.join(ts)}）を 1 つの組にしている（拠点の中の担当の違いが消える。別の組にする）")
        if g["basis"] == "same_name" and len({norm(m["term"]) for m in g["members"]}) > 1:
            rp.fail(f"terms.json {gid}: 名前が違う用語を same_name でまとめている（訳なら translation、言い換えなら equivalent と理由）")
        if g["basis"] in ("translation", "equivalent") and not g["reason"].strip():
            rp.fail(f"terms.json {gid}: {g['basis']} でまとめた理由（reason）が無い（原典の役割の説明・担う作業を書く）")
    used = sc_terms.used_terms(cx.cmp, cx.sites)
    for (s, kind, name) in used:
        if (s, kind, name) not in where:
            rp.fail(f"terms.json: {cx.names[s]}の{dict(actor='担当', system='システム', document='帳票')[kind]}「{name}」がどの組にも無い（sc_terms.py draft で足し、まとめる）")
    for (s, kind, name), gid in where.items():
        if (s, kind, name) not in used and s in cx.ids:
            rp.warn(f"terms.json {gid}: {cx.names[s]}の「{name}」は、比べる箱に出てこない（消すか、拠点の用語の名前を確かめる）")
    for q in tm["questions"]:
        if q.get("group") and q["group"] not in ids:
            rp.fail(f"terms.json 確認事項 {q['id']}: 組 {q['group']} が無い")


# ---------- [axis] ----------

def model(cx):
    from sc_rows import Model
    if not hasattr(cx, "_mdl"):
        cx._mdl = Model(cx.cmp, cx.sites)
    return cx._mdl


def check_axis(cx, rp):
    if not any(cx.snap(s) for s in cx.ids):
        return
    ax, err = try_load(cx.cmp / "axis.json")
    if err or ax is None:
        rp.fail(f"axis.json が{'読めない（' + err + '）' if err else '無い'}（python3 scripts/sc_axis.py draft で下書き）")
        return
    if not shape(rp, ax, "axis", "axis.json"):
        return
    want = "baseline" if cx.sites.get("baseline") else "side_by_side"
    if ax["mode"] != want:
        rp.fail(f"axis.json: mode が {ax['mode']}（sites.json の基準{'あり' if want == 'baseline' else 'なし'}なら {want}）")
    gids = [g["id"] for g in ax["groups"]]
    if len(set(gids)) != len(gids):
        rp.fail("axis.json: 業務分類の ID が重なっている")
    for g in ax["groups"]:
        if not is_jp(g["name"]):
            rp.fail(f"axis.json {g['id']}「{g['name']}」: 業務分類の名前は日本語にする")
        nf = g.get("name_from")
        if nf and not any(c["id"] == nf["category"] for c in (cx.snap(nf["site"]) or {}).get("categories") or []):
            rp.fail(f"axis.json {g['id']}: name_from の分類 {nf['category']} が、{cx.names.get(nf['site'], nf['site'])}の写しに無い")
    rids = [r["id"] for r in ax["rows"]]
    if len(set(rids)) != len(rids):
        rp.fail("axis.json: 行の ID が重なっている")
    mdl = model(cx)
    lang = {x["id"]: x.get("language", "ja") for x in cx.sites["sites"]}
    cover = {s: {} for s in cx.ids}
    for r in ax["rows"]:
        rid = r["id"]
        if r["group"] is not None and r["group"] not in gids:
            rp.fail(f"axis.json {rid}: 業務分類 {r['group']} が groups に無い")
        if r["parent"] is not None and r["parent"] not in rids:
            rp.fail(f"axis.json {rid}: 親の行 {r['parent']} が無い")
        if not is_jp(r["name"]):
            rp.fail(f"axis.json {rid}「{r['name']}」: 行の名前は日本語にする（日本語でない拠点の名前を取ったときは訳す）")
        nf = r["name_from"]
        if nf["flow"] not in ((cx.snap(nf["site"]) or {}).get("flows") or {}):
            rp.fail(f"axis.json {rid}: name_from のフロー {nf['flow']} が、{cx.names.get(nf['site'], nf['site'])}の写しに無い")
        seen = [x["site"] for x in r["sites"]]
        for s in cx.ids:
            if seen.count(s) != 1:
                rp.fail(f"axis.json {rid}: {cx.names[s]}の当て方が {seen.count(s)} 回ある（全拠点を 1 回ずつ書く）")
        base = cx.sites.get("baseline")
        if ax["mode"] == "baseline":
            bp = next((x for x in r["sites"] if x["site"] == base), {})
            if r["origin"] == "baseline" and (not bp.get("parts") or nf["site"] != base):
                rp.fail(f"axis.json {rid}: origin が baseline なのに、基準の拠点のフローを当てていない・名前を基準から取っていない")
            if r["origin"] == "not_in_baseline" and bp.get("parts"):
                rp.fail(f"axis.json {rid}: origin が not_in_baseline なのに、基準の拠点のフローを当てている")
            if r["origin"] == "merged":
                rp.fail(f"axis.json {rid}: 基準ありでは、origin は baseline か not_in_baseline")
        elif r["origin"] != "merged":
            rp.fail(f"axis.json {rid}: 基準なしでは、origin は merged")
        for x in r["sites"]:
            s, rel = x["site"], x["relation"]
            if s not in cx.ids:
                rp.fail(f"axis.json {rid}: 拠点 {s} が sites.json に無い")
                continue
            who = f"axis.json {rid} {cx.names[s]}"
            if rel in ("same", "part", "split") and not x["parts"]:
                rp.fail(f"{who}: {rel} なのに、当てるフロー（parts）が無い")
            if rel in ("none", "out_of_scope") and x["parts"]:
                rp.fail(f"{who}: {rel} なのに、フローを当てている")
            if rel != "same" and not x["reason"].strip():
                rp.fail(f"{who}: {rel} の理由（reason）が無い" + ("（記載なしは、何を探して無かったかを書く）" if rel == "none" else ""))
            if rel == "split" and not any(isinstance(p.get("nodes"), list) for p in x["parts"]):
                rp.warn(f"{who}: split なのに、箱を分けていない（parts.nodes で、この行に入れる箱を書く）")
            snap = cx.snap(s) or {}
            for p in x["parts"]:
                if p["flow"] not in (snap.get("flows") or {}):
                    rp.fail(f"{who}: フロー {p['flow']} が写しに無い")
                    continue
                if lang.get(s, "ja") != "ja" and not is_jp(p.get("title", "")):
                    rp.fail(f"{who} {p['flow']}: 日本語でない拠点のフローには、日本語の名前（title）を書く（納品物に出る）")
                nodes, _e = mdl.part_graph(s, p)
                full = mdl.part_graph(s, dict(p, nodes=None))[0]
                for k in p.get("nodes") or []:
                    if k not in full:
                        rp.fail(f"{who}: 箱 {k} が、フロー {p['flow']}{'（開いた 1 本）' if p['expand'] else ''}に無い")
                for k in nodes:
                    if (snap.get("nodes", {}).get(k) or {}).get("type") in STEP_TYPES:
                        cover[s].setdefault(k, []).append(rid)
    for x in ax["not_compared"]:
        s = x["site"]
        snap = cx.snap(s) or {}
        if x["flow"] not in (snap.get("flows") or {}):
            rp.fail(f"axis.json not_compared: {cx.names.get(s, s)}のフロー {x['flow']} が写しに無い")
            continue
        fl = snap["flows"][x["flow"]]
        for k in (x["nodes"] if isinstance(x["nodes"], list) else fl["nodes"]):
            if (snap.get("nodes", {}).get(k) or {}).get("type") in STEP_TYPES:
                cover[s].setdefault(k, []).append("比べない")
    for s in cx.ids:
        snap = cx.snap(s) or {}
        left = {}
        for k, n in (snap.get("nodes") or {}).items():
            if n["type"] not in STEP_TYPES:
                continue
            got = cover[s].get(k, [])
            if not got:
                left.setdefault(n["flow"], []).append(k)
            elif len(got) > 1:
                rp.fail(f"axis.json: {cx.names[s]}の箱 {k}「{n['label'][:20]}」が {'・'.join(got)} の 2 か所に入っている（ちょうど 1 つの行か not_compared に）")
        for f, ks in left.items():
            rp.fail(f"axis.json: {cx.names[s]}のフロー {f}「{snap['flows'][f]['name']}」の作業の箱 {len(ks)} がどの行にも入っていない（行に当てるか、not_compared に理由。{', '.join(ks[:4])}）")


# ---------- [rows] ----------

def check_rows(cx, rp):
    import sc_match
    mdl = model(cx)
    if not mdl.rows():
        return
    for r in mdl.rows():
        rid = r["id"]
        with A.guarded(rp, f"rows/{rid}.match.json"):
            check_row(cx, rp, mdl, rid, sc_match)


def check_row(cx, rp, mdl, rid, sc_match):
    p = match_path(cx.cmp, rid)
    mt, err = try_load(p)
    if err or mt is None:
        rp.fail(f"rows/{rid}.match.json が{'読めない（' + err + '）' if err else '無い'}（python3 scripts/sc_match.py draft --row {rid} で下書き。references/match-guide.md）")
        return
    if not shape(rp, mt, "match", f"rows/{rid}.match.json"):
        return
    where = f"rows/{rid}"
    if mt["row"] != rid:
        rp.fail(f"{where}: row が {mt['row']}（ファイルの名前と合わせる）")
    tids = [t["id"] for t in mt["tasks"]]
    if len(set(tids)) != len(tids):
        rp.fail(f"{where}: タスクの ID が重なっている")
    parts = mdl.with_parts(rid)
    steps = {s: mdl.step_nodes(rid, s) for s in parts}
    inside = {s: set(mdl.row_nodes(rid, s)[0]) for s in parts}
    seen = {s: {} for s in parts}
    for t in mt["tasks"]:
        tw = f"{where} {t['id']}「{t['name'][:20]}」"
        if not is_jp(t["name"]) or not ends_with_verb(t["name"]):
            rp.fail(f"{tw}: タスクの名前は、日本語の「目的語＋動詞」で終える")
        mem = t["members"]
        for s in mem:
            if s not in parts:
                rp.fail(f"{tw}: {cx.names.get(s, s)}は、この行にフローを当てていない（axis.json）のに、箱がある")
        for s in parts:
            if s not in mem:
                rp.fail(f"{tw}: {cx.names[s]}の欄が無い（箱が無ければ、空の並びで書き、absent に根拠）")
        for s, ks in mem.items():
            if s not in parts:
                continue
            for k in ks:
                n = (cx.snap(s) or {}).get("nodes", {}).get(k)
                if k not in inside[s]:
                    rp.fail(f"{tw}: {cx.names[s]}の箱 {k} は、この行に当てた箱でない（axis.json の parts）")
                elif n["type"] not in STEP_TYPES:
                    rp.fail(f"{tw}: {cx.names[s]}の箱 {k} は作業の箱でない（別のフローへ進む箱・開始・終了は入れない）")
                else:
                    seen[s].setdefault(k, []).append(t["id"])
        present = [s for s in parts if mem.get(s)]
        if not present:
            rp.fail(f"{tw}: どの拠点の箱も入っていない")
            continue
        cnt = [len(mem[s]) for s in present]
        if len(present) >= 2 and min(cnt) >= 2 and not t["joined"].strip():
            rp.fail(f"{tw}: どの拠点も 2 つ以上の箱でこのタスクを書いている（粗い拠点に合わせてまとめるのは、1 つの箱で書いた拠点があるときだけ。"
                    f"分けるか、分けられない理由を joined に書く）")
        if t["joined"].strip() and min(cnt) == 1:
            rp.warn(f"{tw}: joined があるが、1 つの箱で書いた拠点がある（joined は要らない）")
        for s in present:
            if len(mem[s]) > 1 and not connected(mdl, rid, s, mem[s]):
                rp.warn(f"{tw}: {cx.names[s]}の箱が、線でつながっていない（間に、ほかのタスクの箱を挟んでいないか）")
        for s in parts:
            ab = (t["absent"] or {}).get(s)
            if mem.get(s):
                if ab:
                    rp.fail(f"{tw}: {cx.names[s]}は箱があるのに、absent（記載なし）がある")
                continue
            if not ab:
                rp.fail(f"{tw}: {cx.names[s]}に箱が無いのに、記載なしの根拠（absent）が無い（sc_match.py candidates の候補を確かめて書く）")
                continue
            want = [c[0] for c in sc_match.candidates(mdl, rid, t, s)]
            miss = [k for k in want if k not in ab["checked"]]
            if miss:
                rp.fail(f"{tw}: {cx.names[s]}の記載なしで、機械の候補 {', '.join(miss)} を確かめていない（checked に書き、同じ作業でない理由を note に。同じ作業ならタスクに入れる）")
        for s in (t["absent"] or {}):
            if s not in parts:
                rp.fail(f"{tw}: absent の {cx.names.get(s, s)}は、この行にフローを当てていない")
        for i, c in enumerate(t["content"], 1):
            for s in c["sites"]:
                if s not in present:
                    rp.fail(f"{tw} 内容の差 {i}: {cx.names.get(s, s)}は、このタスクに箱が無い（記載なしは、内容の差に書かない）")
            for e in c["evidence"]:
                if e["node"] not in (mem.get(e["site"]) or []):
                    rp.fail(f"{tw} 内容の差 {i}: 根拠の箱 {e['node']} が、{cx.names.get(e['site'], e['site'])}のこのタスクの箱でない")
            if not is_jp(c["text"]):
                rp.fail(f"{tw} 内容の差 {i}: 日本語で書く")
            if not {e["site"] for e in c["evidence"]} - set(c["sites"]):
                rp.warn(f"{tw} 内容の差 {i}: 根拠が差のある拠点の箱だけ（比べた相手の箱も根拠に書く）")
    for s in parts:
        for k in steps[s]:
            got = seen[s].get(k, [])
            if not got:
                rp.fail(f"{where}: {cx.names[s]}の作業の箱 {k}「{(cx.snap(s)['nodes'][k]['label'])[:20]}」がどのタスクにも入っていない")
            elif len(got) > 1:
                rp.fail(f"{where}: {cx.names[s]}の箱 {k} が {'・'.join(got)} の 2 つのタスクに入っている")
    for q in mt["questions"]:
        if q.get("task") and q["task"] not in tids:
            rp.fail(f"{where} 確認事項 {q['id']}: タスク {q['task']} が無い")
        if q.get("site") and q["site"] not in cx.ids:
            rp.fail(f"{where} 確認事項 {q['id']}: 拠点 {q['site']} が無い")


def connected(mdl, rid, sid, keys):
    """箱の組が、組の中の線だけで（向きを問わず）つながっているか"""
    ks = set(keys)
    adj = {k: set() for k in ks}
    for a, b, *_ in mdl.row_nodes(rid, sid)[1]:
        if a in ks and b in ks:
            adj[a].add(b)
            adj[b].add(a)
    start = next(iter(ks))
    seen, todo = {start}, [start]
    while todo:
        for v in adj[todo.pop()]:
            if v not in seen:
                seen.add(v)
                todo.append(v)
    return seen == ks


# ---------- [compute] ----------

def check_compute(cx, rp):
    import json
    import sc_compute
    mdl = model(cx)
    nb = load(cx.cmp / "numbers.json", None) or {"next": 1, "assigned": {}}
    cfg = sc_compute.config()
    for r in mdl.rows():
        rid = r["id"]
        if not mdl.match(rid):
            continue
        res, err = try_load(result_path(cx.cmp, rid))
        if err or res is None:
            rp.fail(f"rows/{rid}.result.json が無い（python3 scripts/sc_compute.py -w <比較の作業フォルダ>）")
            continue
        fresh = sc_compute.compute_row(mdl, rid, cfg, json.loads(json.dumps(nb)))
        if json.dumps(fresh, ensure_ascii=False, sort_keys=True) != json.dumps(res, ensure_ascii=False, sort_keys=True):
            rp.fail(f"rows/{rid}.result.json が、いまのデータと違う（手で直した・計算し直していない。python3 scripts/sc_compute.py -w <比較の作業フォルダ>）")


CHECKS = {"sites": check_sites, "applies": check_applies, "snapshot": check_snapshot, "terms": check_terms, "axis": check_axis, "rows": check_rows,
          "compute": check_compute}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--stage", choices=STAGES + ["all"], default="all")
    ap.add_argument("--fast", action="store_true")
    ap.add_argument("--report")
    ap.add_argument("--all-lines", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    cx = Ctx(cmp, a.fast)
    rp = Report("拠点間比較の点検（sc_check）", fold=None if a.all_lines else 10)
    for st in (STAGES if a.stage == "all" else [a.stage]):
        with A.guarded(rp, st):
            CHECKS[st](cx, rp)
    rp.finish(a.report)


if __name__ == "__main__":
    main()
