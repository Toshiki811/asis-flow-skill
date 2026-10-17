#!/usr/bin/env python3
"""sc_share.py — 共通の資料のページ化と読取（カード）を、拠点の asis の作業フォルダへ写す（工程 1c。共通の資料は 1 回だけ読む）。

  python3 scripts/sc_share.py -w <比較の作業フォルダ> --site TK      # 拠点の棚卸し（inventory.py）のあと、ページ化（render_pages.py）の前に
  python3 scripts/sc_share.py -w <比較の作業フォルダ> --all

拠点の棚卸しの資料のうち、共通の資料と中身が同じもの（sha1 が同じ）について:
  ページの画像と機械抽出（pages/<資料ID>/）、カード（cards/<資料ID>/。doc_id を拠点の資料 ID に付け替える）、
  棚卸しの描画の項目（rendered・page_list など。道の資料 ID も付け替える）を写す。拠点の render_pages.py は、写した資料を飛ばす。
共通の資料のカードは、共通の作業フォルダ（asis/shared）でだけ直す。拠点の側で直したカードは、写し直すと共通のカードに戻る。
共通の資料の読取（カード）が終わっていないうちは、写さない。
"""
import argparse
import json
import shutil
import sys

from sc_common import load, save, shared_work, site_ids, site_work, sites_of, work_dir

import inventory                                                     # asis（描画の項目の一覧 KEEP）


def shared_docs(sh):
    """共通の作業フォルダの、読取の済んだ資料: {sha1: (資料, 足りないカードの数)}"""
    m = load(sh / "manifest.json", None)
    if not m:
        sys.exit(f"共通の資料の棚卸しが無い（python3 <asis>/scripts/inventory.py <共通の資料> -w {sh}）")
    out = {}
    for d in m["docs"]:
        if d.get("duplicate_of") or not d.get("rendered"):
            continue
        lack = sum(1 for p in d.get("page_list") or [] if not (sh / "cards" / d["doc_id"] / f"p{p['page']:04d}.json").exists())
        out[d["sha1"]] = (d, lack)
    return out


def card_for(card, doc_id):
    c = dict(card)
    c["doc_id"] = doc_id
    return c


def share(cmp, sites, sid, quiet=False):
    """1 拠点に写す。戻り値は写した資料の数"""
    sh = shared_work(cmp, sites)
    w = site_work(cmp, sites, sid)
    me = next(x for x in sites["sites"] if x["id"] == sid)
    if sh is None or me.get("uses_shared") is False:
        if not quiet:
            print(f"{sid}: 共通の資料を使わない拠点（写すものは無い）")
        return 0
    src = shared_docs(sh)
    m = load(w / "manifest.json", None)
    if not m:
        sys.exit(f"{sid}: 拠点の棚卸しが無い（python3 <asis>/scripts/inventory.py <共通の資料> <拠点の資料> -w {w}）")
    done = 0
    for d in m["docs"]:
        if d.get("sha1") not in src or d.get("duplicate_of"):
            continue
        s, lack = src[d["sha1"]]
        if lack:
            sys.exit(f"共通の資料「{s['name']}」の読取が終わっていない（カードの無いページ {lack}）。共通の作業フォルダ {sh} で読取を終えてから写す")
        a, b = s["doc_id"], d["doc_id"]
        for sub in ("pages", "cards"):
            dst = w / sub / b
            if dst.exists():
                shutil.rmtree(dst)
            dst.mkdir(parents=True)
            for f in sorted((sh / sub / a).iterdir()):
                if f.suffix == ".json":
                    data = json.loads(f.read_text(encoding="utf-8"))
                    if isinstance(data, dict) and "doc_id" in data:
                        data["doc_id"] = b
                    save(dst / f.name, data)
                else:
                    shutil.copy2(f, dst / f.name)
        for k in inventory.KEEP:
            if k in s:
                d[k] = json.loads(json.dumps(s[k]).replace(f"pages/{a}/", f"pages/{b}/"))
        done += 1
        if not quiet:
            print(f"{sid}: 共通の資料「{s['name']}」を {b} として写した（ページ {len(d.get('page_list') or [])}）")
    save(w / "manifest.json", m)
    if not done and not quiet:
        print(f"{sid}: 拠点の棚卸しに、共通の資料と同じ中身の資料が無い（inventory.py の入力に、共通の資料のフォルダを入れる）")
    return done


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--site")
    g.add_argument("--all", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    for sid in (site_ids(sites) if a.all else [a.site]):
        if sid not in site_ids(sites):
            sys.exit(f"拠点 {sid} が sites.json に無い")
        share(cmp, sites, sid)


if __name__ == "__main__":
    main()
