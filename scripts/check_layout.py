#!/usr/bin/env python3
"""check_layout.py — 書き出した draw.io の配置を点検し、目で見るための画像（PNG）を作る。

  python3 scripts/check_layout.py -w <作業フォルダ>             # 配置の点検＋各ページの PNG（out/png/）
  python3 scripts/check_layout.py -w <作業フォルダ> --no-png    # 点検だけ

check_flow.py は中身（出典・つながり）を見る。ここは見た目を見る。配置は build_drawio.py が書いた *.layout.json を使う。
  FAIL  ノードどうしが重なる
  FAIL  線がノードを貫く（両端のノードは除く）
  FAIL  枠の外に出した出典の字が、ノード・ほかの出典と重なる／線が、出典の字を通る
  FAIL  PNG に書き出せない（draw.io で開けない図になっている）
  FAIL  出どころも行き先も違う線どうしが、同じ所を重なって走る・端どうしがつながって見える（別の箱へ進むように読める）
  WARN  1本の線が、ほかの線と4回以上交わる
  WARN  画像の大きさが配置から求めた大きさと合わない（照合シートの赤枠がずれる）
PNG は必ず自分の目で全ページ見る（Read で開く）。機械は重なりと貫通しか見ない。文字のあふれ・線の回り込みの分かりにくさ・
レーンの順は、機械が、線がまたぐレーンの数が少なくなる順に並べる（決まり 4.2）。それでも読みにくければ、直しの一覧の lane_order で並びを決める（JSON をまるごと書いたフローは、flow.json の lanes を入れ替えるか renumber.py --lanes）。ノードの順は、直しの一覧・flow.json で入れ替えて書き出し直す。
終了コード: FAIL があれば 1。
"""
import argparse
import pathlib
import shutil
import subprocess

from common import Report, load, work_dir
from validate import validate

MAC_CLI = "/Applications/draw.io.app/Contents/MacOS/draw.io"


def drawio_cli():
    return shutil.which("drawio") or shutil.which("draw.io") or (MAC_CLI if pathlib.Path(MAC_CLI).exists() else None)


def export_png(cli, drawio, index, png, scale=1.0):
    png.parent.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([cli, "-x", "-f", "png", "-p", str(index), "-s", str(scale), "-o", str(png), str(drawio)],
                       capture_output=True, text=True, timeout=300)
    return png.exists(), (r.stderr or r.stdout).strip()[-200:]


def seg_hits_rect(p, q, r, pad=2):
    """水平・垂直の線分 p-q が、四角 r=[x,y,w,h] の内側（pad だけ内へ入ったところ）を通るか"""
    x0, y0, x1, y1 = r[0] + pad, r[1] + pad, r[0] + r[2] - pad, r[1] + r[3] - pad
    ax, ay, bx, by = min(p[0], q[0]), min(p[1], q[1]), max(p[0], q[0]), max(p[1], q[1])
    return ax < x1 and bx > x0 and ay < y1 and by > y0 if (ax != bx or ay != by) else False


NEAR, PAD = 4, 10                                                    # 同じ直線とみなす近さ／端どうしがこれより近いと、1本につながって見える


def runs_together(a, b, c, d):
    """線分 a-b と c-d が、同じ直線の上で重なるか、端どうしが近くて1本につながって見えるか（水平どうし・垂直どうしのときだけ）"""
    for ax in (0, 1):
        o = 1 - ax
        if abs(a[ax] - b[ax]) < 0.6 and abs(c[ax] - d[ax]) < 0.6 and abs(a[ax] - c[ax]) < NEAR and abs(a[o] - b[o]) > 0.6 and abs(c[o] - d[o]) > 0.6:
            return min(a[o], b[o]) < max(c[o], d[o]) + PAD and min(c[o], d[o]) < max(a[o], b[o]) + PAD
    return False


def crosses(a, b, c, d):
    """水平の線分と垂直の線分が、端でないところで交わるか"""
    for (p, q), (r, s) in (((a, b), (c, d)), ((c, d), (a, b))):
        if p[1] == q[1] and r[0] == s[0]:
            if min(p[0], q[0]) < r[0] < max(p[0], q[0]) and min(r[1], s[1]) < p[1] < max(r[1], s[1]):
                return True
    return False


def page_problems(pg):
    """1ページの配置の点検。("fail" か "warn", 文) を、見つけた順に返す。
    pg は {"nodes": {id: [x, y, 幅, 高さ]}, "edges": [{"from", "to", "points"}], "captions": {ノードの id: 枠の外に出した出典の [x, y, 幅, 高さ]}}"""
    rects = pg["nodes"]
    caps = pg.get("captions") or {}
    over = lambda r, s: r[0] < s[0] + s[2] - 1 and s[0] < r[0] + r[2] - 1 and r[1] < s[1] + s[3] - 1 and s[1] < r[1] + r[3] - 1
    ids = list(rects)
    for i, u in enumerate(ids):
        for v in ids[i + 1:]:
            if over(rects[u], rects[v]):
                yield "fail", f"ノード {u} と {v} が重なる"
    cids = list(caps)
    for i, u in enumerate(cids):
        for v in ids:
            if over(caps[u], rects[v]):
                yield "fail", f"ノード {u} の出典の字が、ノード {v} と重なる"
        for v in cids[i + 1:]:
            if over(caps[u], caps[v]):
                yield "fail", f"ノード {u} の出典の字が、ノード {v} の出典の字と重なる"
    segs = []
    for e in pg["edges"]:
        pts = e["points"]
        mine = []
        for p, q in zip(pts, pts[1:]):
            mine.append((p, q))
            for nid, r in rects.items():
                if nid in (e["from"], e["to"]):
                    continue
                if seg_hits_rect(p, q, r):
                    yield "fail", f"線 {e['from']}→{e['to']} がノード {nid} を貫く（レーンの順かノードの順を入れ替える）"
            for nid, r in caps.items():
                if seg_hits_rect(p, q, r, pad=1):
                    yield "fail", f"線 {e['from']}→{e['to']} が、ノード {nid} の出典の字を通る（レーンの順かノードの順を入れ替える）"
        segs.append((e, mine))
    for i, (e1, s1) in enumerate(segs):
        n_cross = 0
        for j, (e2, s2) in enumerate(segs):
            if i == j:
                continue
            for x in s1:
                for y in s2:
                    if crosses(x[0], x[1], y[0], y[1]):
                        n_cross += 1
            if j > i and e1["from"] != e2["from"] and e1["to"] != e2["to"] and any(runs_together(x[0], x[1], y[0], y[1]) for x in s1 for y in s2):
                yield "fail", (f"線 {e1['from']}→{e1['to']} と {e2['from']}→{e2['to']} が、同じ所を重なって走る（別の箱へ進むように読める。"
                               f"線の通り道が足りない。箱・分岐を減らす（畳む・まとめる・図を分ける）か、レーンの順・ノードの順を入れ替える）")
        if n_cross >= 4:
            yield "warn", f"線 {e1['from']}→{e1['to']} がほかの線と {n_cross} 回交わる（レーンの順を見直す。直しの一覧のフローは lane_order、JSON をまるごと書いたフローは lanes の順か renumber.py --lanes）"


def stale_pngs(folder, layouts):
    """いまのページの並びに無い PNG（前の書き出しの残り）。ページが増減すると番号がずれて、古い画像が納品物に残る"""
    out = []
    for lp in layouts:
        lay = load(lp)
        stem = pathlib.Path(lay["file"]).stem
        keep = {f"{stem}_p{pg['index']:02d}_{pg['id']}.png" for pg in lay["pages"]}
        for q in sorted(folder.glob("*.png")):
            rest = q.name[len(stem):]
            if q.name.startswith(stem) and rest[:2] == "_p" and rest[2:4].isdigit() and rest[4:5] == "_" and q.name not in keep:
                out.append(q)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--no-png", action="store_true")
    ap.add_argument("--report")
    ap.add_argument("--all-lines", action="store_true", help="同じ種類の行を畳まずに全部出す")
    a = ap.parse_args()
    work = work_dir(a.work)
    rp = Report("配置の点検（check_layout）", fold=None if a.all_lines else 10)
    layouts = sorted((work / "out").glob("*.layout.json"))
    if not layouts:
        rp.fail("out/*.layout.json が無い（build_drawio.py を実行する）")
        rp.finish(a.report)
    cli = None if a.no_png else drawio_cli()
    if not a.no_png and cli is None:
        rp.warn("draw.io のコマンドが見つからないので PNG を作れない（draw.io デスクトップを入れる。見た目は draw.io で開いて確かめる）")
    made = []
    for lp in layouts:
        lay = load(lp)
        errs = validate(lay, "layout")
        if errs:
            rp.fail(f"{lp.name}: 形が崩れている（{'／'.join(errs[:3])}）。build_drawio.py で書き出し直す")
            continue
        drawio = lp.with_name(lay["file"])
        for pg in lay["pages"]:
            where = f"{lay['file']}［{pg['name']}］"
            for kind, text in page_problems(pg):
                (rp.fail if kind == "fail" else rp.warn)(f"{where}: {text}")
            if cli:
                png = work / "out" / "png" / f"{drawio.stem}_p{pg['index']:02d}_{pg['id']}.png"
                ok, msg = export_png(cli, drawio, pg["index"], png)
                if not ok:
                    rp.fail(f"{where}: PNG に書き出せない（{msg}）")
                    continue
                made.append(png)
                try:
                    from PIL import Image
                    w, h = Image.open(png).size
                    b = pg["bounds"]
                    ew, eh = b[2] - b[0], b[3] - b[1]                 # bounds は四隅の見えない点まで含む
                    if abs(w - ew) > 6 or abs(h - eh) > 6:
                        rp.warn(f"{where}: 画像 {w}×{h} が配置から求めた {ew:.0f}×{eh:.0f} と合わない（照合シートの赤枠がずれることがある）")
                except ImportError:
                    pass
    rp.info(f"点検したページ {sum(len(load(lp)['pages']) for lp in layouts)}")
    if cli:
        stale = stale_pngs(work / "out" / "png", layouts)
        for q in stale:
            q.unlink()
        if stale:
            rp.info(f"前の書き出しの PNG {len(stale)} 枚を消した（ページが増減して、番号がずれた分）")
    if made:
        rp.info(f"PNG {len(made)} 枚 → {work / 'out' / 'png'}（全ページを Read で開いて目で見る）")
        for p in made:
            rp.info(f"  {p}")
    rp.finish(a.report)


if __name__ == "__main__":
    main()
