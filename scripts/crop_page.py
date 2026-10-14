#!/usr/bin/env python3
"""crop_page.py — ページの一部を拡大した画像を作る（字が小さい・図が細かいページを読み直すため）。

  python3 scripts/crop_page.py -w <作業フォルダ> M03 2                 # おすすめの分割（機械抽出の suggest_grid）で全体を分けて拡大
  python3 scripts/crop_page.py -w <作業フォルダ> M03 2 --grid 3x2      # 横3×縦2 に分ける
  python3 scripts/crop_page.py -w <作業フォルダ> M03 2 --region 0.5,0,1,0.4   # 左上を(0,0)・右下を(1,1)とした範囲 x0,y0,x1,y1

出力は pages/<資料ID>/crops/ に置き、パスを表示する。分割は隣と少し重ねる（境目の文字が切れないように）。
元が PDF（Office の文書を含む）なら PDF から描き直すので、拡大しても字がぼけない。
"""
import argparse
import sys

from common import load, page_json, page_stem, work_dir

LONG_SIDE = 1568
OVERLAP = 0.08


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("doc_id")
    ap.add_argument("page", type=int)
    ap.add_argument("--grid", help="横x縦（例 2x3）")
    ap.add_argument("--region", help="x0,y0,x1,y1（0〜1）")
    a = ap.parse_args()
    work = work_dir(a.work)
    pj = load(page_json(work, a.doc_id, a.page))
    if not pj or not pj.get("render"):
        sys.exit("このページには画像がない（文字だけのページ）")
    regions = []
    if a.region:
        x0, y0, x1, y1 = [float(v) for v in a.region.split(",")]
        regions.append(("r", x0, y0, x1, y1))
    else:
        cols, rows = [int(v) for v in a.grid.lower().split("x")] if a.grid else (pj.get("suggest_grid") or [2, 2])
        for r in range(rows):
            for c in range(cols):
                regions.append((f"r{r + 1}c{c + 1}", max(0, c / cols - OVERLAP / 2), max(0, r / rows - OVERLAP / 2),
                                min(1, (c + 1) / cols + OVERLAP / 2), min(1, (r + 1) / rows + OVERLAP / 2)))
    out_dir = work / "pages" / a.doc_id / "crops"
    out_dir.mkdir(parents=True, exist_ok=True)
    rd = pj["render"]
    for tag, x0, y0, x1, y1 in regions:
        name = f"{page_stem(a.page)}_{tag}" + (f"_{x0:.2f}-{y0:.2f}-{x1:.2f}-{y1:.2f}" if tag == "r" else "") + ".png"
        out = out_dir / name
        if rd.get("pdf"):
            import fitz
            doc = fitz.open(rd["pdf"])
            page = doc[rd["pdf_page"]]
            R = page.rect
            clip = fitz.Rect(R.x0 + R.width * x0, R.y0 + R.height * y0, R.x0 + R.width * x1, R.y0 + R.height * y1)
            zoom = min(6.0, LONG_SIDE / max(clip.width, clip.height))
            page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False).save(str(out))
            doc.close()
        else:
            from PIL import Image
            im = Image.open(rd["image"]).convert("RGB")
            box = (round(im.width * x0), round(im.height * y0), round(im.width * x1), round(im.height * y1))
            im = im.crop(box)
            scale = LONG_SIDE / max(im.size)
            if scale < 1:
                im = im.resize((round(im.width * scale), round(im.height * scale)))
            im.save(out)
        print(out)


if __name__ == "__main__":
    main()
