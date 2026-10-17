#!/usr/bin/env python3
"""sc_init.py — 入力のフォルダの形から、拠点の一覧（sites.json）の下書きを書き、比較の作業フォルダを作る（工程 0）。

  python3 scripts/sc_init.py <入力フォルダ> -w <比較の作業フォルダ> [--baseline <拠点のフォルダ名>]

入力の形: <入力フォルダ>/<拠点のフォルダ>/…（拠点ごとの資料）と、<入力フォルダ>/共通/…（全社共通の資料。無くてよい）。
共通の資料とみなすフォルダ名: 共通・共通資料・全社共通・common・shared。
下書きのあと、利用者と次を決めて sites.json を直す: 基準の拠点（baseline。無ければ null）・拠点の名前・別名（aliases。共通の資料の
書き分け「大阪支店の場合は…」を拾うのに使う）・資料の言葉（language。ja 以外は、カードとフローの文を日本語に訳して書く）。
拠点 ID は納品物に出ない（S1・S2… のままでよい）。すでに sites.json があれば書かない（--force で書き直す）。
"""
import argparse
import pathlib
import re
import subprocess
import sys

from sc_common import ASIS, save, sites_of
from sc_validate import validate

SHARED_NAMES = ("共通", "共通資料", "全社共通", "common", "shared")


def site_id(name, used):
    """フォルダ名から拠点 ID（英字があれば、その頭の大文字。無ければ S1・S2…）"""
    letters = re.sub(r"[^A-Za-z0-9]", "", name).upper()
    cand = letters[:3] if len(letters) >= 2 and letters[0].isalpha() else ""
    n = 1
    while not cand or cand in used:
        cand, n = f"S{n}", n + 1
    used.add(cand)
    return cand


def asis_version():
    r = subprocess.run(["git", "-C", str(ASIS), "rev-parse", "--short", "HEAD"], capture_output=True, text=True)
    return r.stdout.strip() if r.returncode == 0 else ""


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--baseline", help="基準の拠点のフォルダ名（省けば基準なし）")
    ap.add_argument("--force", action="store_true", help="すでにある sites.json を書き直す")
    a = ap.parse_args()
    src = pathlib.Path(a.input).expanduser().resolve()
    cmp = pathlib.Path(a.work).expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"入力フォルダが無い: {src}")
    if (cmp / "sites.json").exists() and not a.force:
        sites_of(cmp)
        sys.exit(f"sites.json はもうある（{cmp / 'sites.json'}）。書き直すときは --force")
    dirs = sorted(p for p in src.iterdir() if p.is_dir() and not p.name.startswith("."))
    shared = [p for p in dirs if p.name.lower() in SHARED_NAMES]
    sites, used = [], set()
    for p in dirs:
        if p in shared:
            continue
        sites.append({"id": site_id(p.name, used), "name": p.name, "aliases": [], "language": "ja", "inputs": [str(p)]})
    if len(sites) < 2:
        sys.exit(f"拠点のフォルダが 2 つ以上要る（見つかったのは {len(sites)}: {', '.join(x['name'] for x in sites) or 'なし'}）。"
                 f"<入力フォルダ>/<拠点>/… の形に置く")
    base = None
    if a.baseline:
        base = next((x["id"] for x in sites if x["name"] == a.baseline), None)
        if base is None:
            sys.exit(f"基準の拠点「{a.baseline}」が拠点のフォルダに無い（{', '.join(x['name'] for x in sites)}）")
    data = {"baseline": base, "shared": {"inputs": [str(p) for p in shared]} if shared else None, "sites": sites,
            "asis": {"path": str(ASIS), "commit": asis_version()}}
    errs = validate(data, "sites")
    if errs:
        sys.exit("下書きが形に合わない: " + "; ".join(errs[:5]))
    for d in ("asis", "snap", "rows", "reports", "review", "out"):
        (cmp / d).mkdir(parents=True, exist_ok=True)
    save(cmp / "sites.json", data)
    print(f"sites.json の下書きを書いた → {cmp / 'sites.json'}")
    print(f"  共通の資料: {', '.join(p.name for p in shared) or 'なし'}")
    for x in sites:
        print(f"  {x['id']}  {x['name']}{'【基準】' if x['id'] == base else ''}  資料 {x['inputs'][0]}")
    print("\n次: 利用者と、基準の拠点・拠点の名前と別名（aliases）・資料の言葉（language）を決めて sites.json を直し、"
          "python3 scripts/sc_status.py -w " + str(cmp))


if __name__ == "__main__":
    main()
