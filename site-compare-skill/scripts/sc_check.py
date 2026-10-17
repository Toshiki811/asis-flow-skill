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
"""
import argparse
import pathlib
import re
import sys

from sc_common import (OTHER_SITE, SWITCH_WORDS, A, Report, load, mentions, norm, shared_work, site_ids, site_names, site_work, sites_of,
                       snap_path, try_load, work_dir)
from sc_validate import validate

STAGES = ["sites", "applies", "snapshot"]


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


CHECKS = {"sites": check_sites, "applies": check_applies, "snapshot": check_snapshot}


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
