#!/usr/bin/env python3
"""sc_layout.py — 拠点比較フローの配置の点検と、PNG の書き出し（工程 8）。

  python3 scripts/sc_layout.py -w <比較の作業フォルダ>                 # 点検して、全ページを out/png/ に PNG で書き出す（draw.io デスクトップが要る）
  python3 scripts/sc_layout.py -w <比較の作業フォルダ> --no-png        # 点検だけ
  python3 scripts/sc_layout.py -w <比較の作業フォルダ> --preview       # draw.io の無い環境で、下見の画像を out/下見/ に（納品物にしない）

点検は asis-flow-skill の check_layout.py と同じ（箱の重なり・線が箱を貫く・出どころも行き先も違う線が同じ所を走る は FAIL。交わりの多い線は WARN）。
PNG は、目で全ページを見るためのもの（帯のそろい・空き枠・内訳の枠・色）。納品の門は、PNG が新しくそろっていなければ止まる。
"""
import argparse
import pathlib
import sys

from sc_common import FIG, Report, load, work_dir

from check_layout import drawio_cli, export_png, page_problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--no-png", action="store_true")
    ap.add_argument("--preview", action="store_true")
    ap.add_argument("--report")
    ap.add_argument("--all-lines", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    rp = Report("拠点比較フローの配置の点検（sc_layout）", fold=None if a.all_lines else 10)
    fig = cmp / "out" / FIG
    lp = fig.with_suffix(".layout.json")
    lay = load(lp, None)
    if not fig.exists() or not lay:
        rp.fail(f"{FIG} が無い（python3 scripts/sc_drawio.py -w <比較の作業フォルダ>）")
        rp.finish(a.report)
    for pg in lay["pages"]:
        found = page_problems(pg)
        for kind, text in found:
            if kind == "fail":
                rp.fail(f"{FIG}［{pg['name']}］: {text}", at=found.gi_frame.f_lineno)
            else:
                rp.warn(f"{FIG}［{pg['name']}］: {text}")
    if a.preview:
        import sc_preview
        out = sc_preview.render(fig, (cmp / "out" / "下見").resolve() if (cmp / "out" / "下見").mkdir(parents=True, exist_ok=True) is None else None)
        rp.info(f"下見の画像 {len(out)} 枚 → out/下見/（正確な描画ではない。納品物にしない）")
    elif not a.no_png:
        cli = drawio_cli()
        if cli is None:
            rp.warn("draw.io のコマンドが見つからないので PNG を作れない（draw.io デスクトップを入れる。無い環境の下見は --preview）")
        else:
            made = 0
            for pg in lay["pages"]:
                png = cmp / "out" / "png" / f"{fig.stem}_p{pg['index']:02d}_{pg['id']}.png"
                ok, msg = export_png(cli, fig, pg["index"], png)
                if not ok:
                    rp.fail(f"{FIG}［{pg['name']}］: PNG に書き出せない（{msg}）")
                made += ok
            rp.info(f"PNG {made} 枚 → out/png/（全ページを目で見る）")
    rp.finish(a.report)


if __name__ == "__main__":
    main()
