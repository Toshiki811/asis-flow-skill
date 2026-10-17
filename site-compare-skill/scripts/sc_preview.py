#!/usr/bin/env python3
"""sc_preview.py — 拠点比較フロー（.drawio）の下見の画像を、Pillow で描く（draw.io デスクトップの無い環境で、形とそろい方を目で確かめるため）。

  python3 scripts/sc_preview.py <図.drawio> <出力フォルダ> [倍率]

正確な描画ではない（字の折り返し・矢じり・色の細部は draw.io と違う）。納品物の PNG は、draw.io デスクトップで書き出す（sc_layout.py）。
下見の画像は、納品物にしない。
"""
import html
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont

FONTS = ["/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
         "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", "C:/Windows/Fonts/meiryo.ttc"]
FONT = next((f for f in FONTS if pathlib.Path(f).exists()), None)


def style_of(s):
    return dict(p.split("=", 1) for p in (s or "").split(";") if "=" in p)


def text_of(v):
    v = re.sub(r"<br\s*/?>", "\n", v or "")
    v = re.sub(r"<[^>]+>", "", v)
    return html.unescape(v)


def render(path, out_dir, scale=0.8):
    tree = ET.parse(path)
    lay = json.loads(pathlib.Path(str(path).replace(".drawio", ".layout.json")).read_text(encoding="utf-8"))
    out = []
    for pi, dg in enumerate(tree.getroot().findall("diagram")):
        cells = []
        root = dg.find("mxGraphModel/root")
        for el in root:
            if el.tag == "object":
                c = el.find("mxCell")
                cells.append((el.get("id"), el.get("label"), c))
            elif el.tag == "mxCell" and el.get("vertex") == "1":
                cells.append((el.get("id"), el.get("value"), el))
        geos = []
        for cid, val, c in cells:
            if c is None or c.get("vertex") != "1":
                continue
            g = c.find("mxGeometry")
            if g is None:
                continue
            x, y, w, h = (float(g.get(k, 0)) for k in ("x", "y", "width", "height"))
            geos.append((cid, val, c.get("style"), x, y, w, h))
        if not geos:
            continue
        W = int(max(x + w for _, _, _, x, _, w, _ in geos) + 60)
        H = int(max(y + h for _, _, _, _, y, _, h in geos) + 60)
        img = Image.new("RGB", (int(W * scale), int(H * scale)), "white")
        d = ImageDraw.Draw(img)
        f = ImageFont.truetype(FONT, max(9, int(10 * scale))) if FONT else ImageFont.load_default()
        S = lambda v: v * scale
        for cid, val, st, x, y, w, h in geos:
            s = style_of(st)
            fill = s.get("fillColor", "#ffffff")
            fill = None if fill == "none" else fill
            stroke = s.get("strokeColor", "#000000")
            stroke = None if stroke == "none" else stroke
            box = [S(x), S(y), S(x + w), S(y + h)]
            if (st or "").startswith("line"):
                d.line([S(x), S(y + h / 2), S(x + w), S(y + h / 2)], fill=stroke or "#999", width=1)
                continue
            if (st or "").startswith("rhombus"):
                pts = [(S(x + w / 2), S(y)), (S(x + w), S(y + h / 2)), (S(x + w / 2), S(y + h)), (S(x), S(y + h / 2))]
                d.polygon(pts, fill=fill, outline=stroke)
            elif (st or "").startswith("ellipse"):
                d.ellipse(box, fill=fill, outline=stroke)
            elif not (st or "").startswith("text"):
                try:
                    d.rectangle(box, fill=fill, outline=stroke, width=2 if "dashed" not in (st or "") else 1)
                except ValueError:
                    pass
            t = text_of(val)
            if t:
                lines = t.split("\n")
                yy = S(y) + 3
                for ln in lines[:8]:
                    d.text((S(x) + 4, yy), ln[:40], fill=s.get("fontColor", "#000000") if s.get("fontColor", "").startswith("#") else "#000", font=f)
                    yy += 12 * scale + 2
        pg = lay["pages"][pi]
        for e in pg["edges"]:
            pts = [(S(a), S(b)) for a, b in e["points"]]
            d.line(pts, fill="#333333", width=1)
            if len(pts) >= 2:
                (x1, y1), (x2, y2) = pts[-2], pts[-1]
                d.ellipse([x2 - 3, y2 - 3, x2 + 3, y2 + 3], fill="#333333")
        p = pathlib.Path(out_dir) / f"{pathlib.Path(path).stem}_p{pi + 1:02d}.png"
        img.save(p)
        out.append(p)
    return out


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    pathlib.Path(sys.argv[2]).mkdir(parents=True, exist_ok=True)
    for p in render(sys.argv[1], sys.argv[2], float(sys.argv[3]) if len(sys.argv) > 3 else 0.8):
        print(p)
