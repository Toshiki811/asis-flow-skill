#!/usr/bin/env python3
"""sc_layout.py — 拠点比較フローの配置の点検と、PNG の書き出し（工程 8）。

  python3 scripts/sc_layout.py -w <比較の作業フォルダ>                 # 点検して、全ページを out/png/ に PNG で書き出す（draw.io デスクトップが要る）
  python3 scripts/sc_layout.py -w <比較の作業フォルダ> --no-png        # 点検だけ
  python3 scripts/sc_layout.py -w <比較の作業フォルダ> --preview       # draw.io の無い環境で、下見の画像を out/下見/ に（納品物にしない）

点検は asis-flow-skill の check_layout.py と同じ（箱の重なり・線が箱を貫く・出どころも行き先も違う線が同じ所を走る は FAIL。交わりの多い線は WARN）。
PNG は、目で全ページを見るためのもの（帯のそろい・空き枠・内訳の枠・色）。納品の門は、PNG が新しくそろっていなければ止まる。
"""
import argparse
import os
import pathlib
import sys

from sc_common import FIG, Report, load, work_dir

from check_layout import drawio_cli, export_png, page_problems


def preview(cmp, fig, rp):
    """下見の画像を out/下見/ に（draw.io と同じ描画エンジン mxGraph が使えればそれで、無ければ Pillow の近似。どちらも納品物にしない）"""
    import preview_mx
    out_dir = cmp / "out" / "下見"
    out_dir.mkdir(parents=True, exist_ok=True)
    for q in out_dir.glob("*.png"):                                  # 2 つの描き方で名前の形が違うので、前の分は消す
        q.unlink()
    if preview_mx.usable():
        try:
            pages = preview_mx.render(fig, out_dir)
            made = [p["png"] for p in pages if p.get("png")]
            for p in pages:
                if p.get("error"):
                    rp.info(f"{FIG}［{p.get('name', '')}］: 下見を描けない（{p['error']}）")
            rp.info(f"下見の画像 {len(made)} 枚 → out/下見/（draw.io と同じ描画エンジン mxGraph の近似。字の幅・帳票の折れ角が違う。納品物にしない。全ページを Read で開いて目で見る）")
            for p in made:
                rp.info(f"  {p}")
            return
        except preview_mx.PreviewError as e:
            rp.info(f"mxGraph の下見を描けなかった（{e}）。Pillow の近似で描く")
    import sc_preview
    out = sc_preview.render(fig, out_dir)
    rp.info(f"下見の画像 {len(out)} 枚 → out/下見/（Pillow の近似。線のラベル・字の折り返しは draw.io と違う。納品物にしない。mxGraph の下見は python3 {preview_mx.SKILL / 'scripts' / 'preview_mx.py'} --probe）")


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
        preview(cmp, fig, rp)
    elif not a.no_png:
        cli = drawio_cli()
        if cli is None:
            rp.warn("draw.io のコマンドが見つからないので PNG を作れない（draw.io デスクトップを入れる。無い環境の下見は --preview）")
            if os.environ.get("ASIS_NO_PREVIEW") != "1":
                preview(cmp, fig, rp)
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
