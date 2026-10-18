#!/usr/bin/env python3
"""preview_mx.py — draw.io デスクトップの無い環境で、図（.drawio）の下見の PNG を描く（draw.io の描画エンジン mxGraph を Chromium で動かす近似描画）。

  python3 scripts/preview_mx.py <図.drawio> <出力フォルダ> [--scale 2]
  python3 scripts/preview_mx.py --probe                                  # 使えるか（node・Playwright・mxgraph・Chromium）と、使えない理由

要るもの: node、Playwright（npm -g。Chromium は PLAYWRIGHT_BROWSERS_PATH の下）、mxgraph（npm install --no-save --prefix <このスキル> mxgraph@4.2.2。
環境変数 MXGRAPH_DIR でも指せる）。playwright install は要らない（入っている Chromium を使う）。
draw.io と同じ描画エンジンの公開版（mxGraph 4.2.2）で描くので、配置・線の通し方は同じになるが、字体（字の幅・折り返し）・帳票の折れ角・
サブプロセスの二重線は違う。下見は納品物にしない（納品の門は、draw.io デスクトップで書き出した PNG を求める）。
出力の名前は check_layout.py の PNG と同じ（<図>_p<番号>_<ページ ID>.png）。圧縮された diagram（draw.io で保存し直したもの）は描けない。
"""
import argparse
import json
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile

from common import font_env

SKILL = pathlib.Path(__file__).resolve().parent.parent
CAP = 16000                                                          # Chromium が 1 枚で撮れる長辺のおおよその上限（px）

# Chromium の中で動かすハーネス（node で実行。Playwright と mxClient.js の場所は環境変数で受ける）
HARNESS = r"""
const fs = require('fs'), path = require('path');
const { chromium } = require(process.env.ASIS_PW_MODULE);
const [, , drawio, outDir, scaleArg, capArg] = process.argv;
const want = parseFloat(scaleArg || '2'), cap = parseInt(capArg || '16000', 10);
const xml = fs.readFileSync(drawio, 'utf8');
const stem = path.basename(drawio, '.drawio');
const font = process.env.ASIS_MX_FONT || 'Yu Gothic UI,IPAPGothic,sans-serif';

function renderOne({ xml, i, scale, font, cap }) {
  const container = document.getElementById('g');
  container.innerHTML = '';
  const doc = mxUtils.parseXml(xml);
  const diag = doc.documentElement.getElementsByTagName('diagram')[i];
  const id = diag.getAttribute('id') || String(i + 1), name = diag.getAttribute('name') || '';
  const model = diag.getElementsByTagName('mxGraphModel')[0];
  if (!model) return { id, name, error: '圧縮された diagram（draw.io で保存し直したもの）は描けない' };
  const C = mxConstants;
  C.DEFAULT_FONTFAMILY = font;
  // draw.io だけの shape は、近い形で代える（帳票の折れ角・サブプロセスの二重線は出ない）
  mxCellRenderer.registerShape('note', mxRectangleShape);
  mxCellRenderer.registerShape('process', mxRectangleShape);
  mxCellRenderer.registerShape('startState', mxEllipse);
  mxCellRenderer.registerShape('endState', mxDoubleEllipse);
  const graph = new mxGraph(container);
  graph.setEnabled(false);
  graph.setTooltips(false);
  graph.foldingEnabled = false;                                      // レーンの折りたたみの画像を取りに行かない
  graph.convertValueToString = function (c) {                        // <object label=…> の label（素のままだと「object」になる）
    const v = c.value;
    if (v != null && v.nodeType === 1) return v.getAttribute('label') || '';
    return v == null ? '' : String(v);
  };
  graph.isHtmlLabel = function (c) {                                 // draw.io と同じく、style の html=1 で決める
    const s = this.getCurrentCellStyle(c);
    return s != null && String(s.html) === '1';
  };
  // draw.io の既定のスタイル（素の mxGraph は青い箱で、ellipse・rhombus・swimlane の名前も知らない）
  const ss = graph.getStylesheet();
  Object.assign(ss.getDefaultVertexStyle(), { [C.STYLE_FILLCOLOR]: '#ffffff', [C.STYLE_STROKECOLOR]: '#000000', [C.STYLE_FONTCOLOR]: '#000000',
    [C.STYLE_FONTSIZE]: 12, [C.STYLE_FONTFAMILY]: font });
  Object.assign(ss.getDefaultEdgeStyle(), { [C.STYLE_STROKECOLOR]: '#000000', [C.STYLE_FONTCOLOR]: '#000000', [C.STYLE_FONTSIZE]: 11,
    [C.STYLE_FONTFAMILY]: font, [C.STYLE_ENDARROW]: C.ARROW_CLASSIC, [C.STYLE_LABEL_BACKGROUNDCOLOR]: '#ffffff' });
  const text = { [C.STYLE_FILLCOLOR]: 'none', [C.STYLE_STROKECOLOR]: 'none', [C.STYLE_ALIGN]: C.ALIGN_LEFT, [C.STYLE_VERTICAL_ALIGN]: C.ALIGN_TOP };
  ss.putCellStyle('text', text);
  ss.putCellStyle('edgeLabel', Object.assign({}, text, { [C.STYLE_ALIGN]: C.ALIGN_CENTER, [C.STYLE_VERTICAL_ALIGN]: C.ALIGN_MIDDLE,
    [C.STYLE_LABEL_BACKGROUNDCOLOR]: '#ffffff', [C.STYLE_FONTSIZE]: 11 }));
  ss.putCellStyle('label', {});
  ss.putCellStyle('swimlane', { [C.STYLE_SHAPE]: C.SHAPE_SWIMLANE, [C.STYLE_FONTSTYLE]: 1, [C.STYLE_STARTSIZE]: 23, [C.STYLE_FONTSIZE]: 12,
    [C.STYLE_VERTICAL_ALIGN]: C.ALIGN_TOP, [C.STYLE_ALIGN]: C.ALIGN_CENTER });
  ss.putCellStyle('group', { [C.STYLE_FILLCOLOR]: 'none', [C.STYLE_STROKECOLOR]: 'none', [C.STYLE_VERTICAL_ALIGN]: C.ALIGN_TOP });
  ss.putCellStyle('ellipse', { [C.STYLE_SHAPE]: C.SHAPE_ELLIPSE, [C.STYLE_PERIMETER]: 'ellipsePerimeter' });
  ss.putCellStyle('rhombus', { [C.STYLE_SHAPE]: C.SHAPE_RHOMBUS, [C.STYLE_PERIMETER]: 'rhombusPerimeter' });
  ss.putCellStyle('line', { [C.STYLE_SHAPE]: C.SHAPE_LINE, [C.STYLE_STROKEWIDTH]: 4, [C.STYLE_VERTICAL_ALIGN]: C.ALIGN_TOP, [C.STYLE_SPACING_TOP]: 8 });
  const mdoc = mxUtils.createXmlDocument();                         // ページごとに別の文書に（id 0・1 は全ページにあるので、1 つの文書だと重なる）
  mdoc.appendChild(mdoc.importNode(model, true));
  const codec = new mxCodec(mdoc);
  const gm = graph.getModel();
  gm.beginUpdate();
  try { codec.decode(mdoc.documentElement, gm); } finally { gm.endUpdate(); }
  const b = graph.getGraphBounds();
  let s = scale;
  const side = Math.max(b.width, b.height, 1);
  if (side * s > cap) s = cap / side;
  const w = Math.max(1, Math.ceil(b.width * s)), h = Math.max(1, Math.ceil(b.height * s));
  container.style.width = w + 'px';
  container.style.height = h + 'px';
  graph.view.scaleAndTranslate(s, -b.x, -b.y);
  return { id, name, w, h, scale: s, cells: gm.getChildCount(gm.getRoot()) };
}

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 800, height: 600 }, deviceScaleFactor: 1 });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.setContent('<!doctype html><html><head><meta charset="utf-8"></head>'
    + '<body style="margin:0;background:#fff"><div id="g" style="position:absolute;left:0;top:0;overflow:hidden"></div></body></html>');
  await page.evaluate(() => { window.mxLoadStylesheets = false; window.mxLoadResources = false; window.mxForceIncludes = false; window.mxBasePath = '.'; });
  await page.addScriptTag({ path: process.env.ASIS_MXCLIENT });
  const n = await page.evaluate(x => mxUtils.parseXml(x).documentElement.getElementsByTagName('diagram').length, xml);
  const out = [];
  for (let i = 0; i < n; i++) {
    let m;
    try { m = await page.evaluate(renderOne, { xml, i, scale: want, font, cap }); }
    catch (e) { m = { id: '', name: '', error: String(e.message || e).slice(0, 200) }; }
    if (m.error) { out.push({ index: i + 1, id: m.id, name: m.name, error: m.error }); continue; }
    await page.setViewportSize({ width: m.w, height: m.h });
    const png = path.join(outDir, `${stem}_p${String(i + 1).padStart(2, '0')}_${m.id}.png`);
    await page.screenshot({ path: png, clip: { x: 0, y: 0, width: m.w, height: m.h } });
    out.push({ index: i + 1, id: m.id, name: m.name, png, w: m.w, h: m.h, scale: m.scale });
  }
  await browser.close();
  for (const e of errors) console.error('page: ' + e);
  console.log(JSON.stringify(out));
})().catch(e => { console.error(e.stack || String(e)); process.exit(1); });
"""


class PreviewError(Exception):
    pass


def node_bin():
    return shutil.which("node")


def npm_root():
    """npm の全体の node_modules（Playwright がここにある）"""
    npm = shutil.which("npm")
    if npm:
        try:
            r = subprocess.run([npm, "root", "-g"], capture_output=True, text=True, timeout=30)
            p = pathlib.Path(r.stdout.strip())
            if r.returncode == 0 and p.exists():
                return p
        except (OSError, subprocess.SubprocessError):
            pass
    nb = node_bin()
    if nb:
        p = pathlib.Path(nb).resolve().parent.parent / "lib" / "node_modules"
        if p.exists():
            return p
    return None


def find_mxclient():
    cands = []
    if os.environ.get("MXGRAPH_DIR"):
        cands.append(pathlib.Path(os.environ["MXGRAPH_DIR"]).expanduser())
    cands.append(SKILL / "node_modules" / "mxgraph")
    root = npm_root()
    if root:
        cands.append(root / "mxgraph")
    for d in cands:
        p = d / "javascript" / "mxClient.js"
        if p.exists():
            return p
    return None


def find_playwright():
    env = os.environ.get("ASIS_PW_MODULE")
    if env and (pathlib.Path(env) / "package.json").exists():
        return pathlib.Path(env)
    for d in [SKILL / "node_modules", npm_root()]:
        if d and (d / "playwright" / "package.json").exists():
            return d / "playwright"
    return None


def browsers_dir():
    return pathlib.Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH") or (pathlib.Path.home() / ".cache" / "ms-playwright"))


_PROBE = None


def probe(fresh=False):
    """{node, playwright, mxclient, browsers, ok, why}"""
    global _PROBE
    if _PROBE is not None and not fresh:
        return _PROBE
    node, pw, mx, br = node_bin(), find_playwright(), find_mxclient(), browsers_dir()
    why = []
    if not node:
        why.append("node が無い")
    if not pw:
        why.append("Playwright が無い（npm install -g playwright）")
    if not mx:
        why.append(f"mxgraph が無い（npm install --no-save --prefix {SKILL} mxgraph@4.2.2）")
    if not (br.exists() and any(p.name.startswith("chromium") for p in br.iterdir())):
        why.append(f"Chromium が無い（{br}。npx playwright install chromium）")
    _PROBE = {"node": node, "playwright": str(pw) if pw else None, "mxclient": str(mx) if mx else None, "browsers": str(br), "ok": not why, "why": "・".join(why)}
    return _PROBE


def usable():
    return probe()["ok"]


def render(drawio, out_dir, scale=2.0, timeout=600):
    """図の全ページを描く → [{index, id, name, png, w, h, scale}]（描けないページは png の代わりに error）"""
    pr = probe()
    if not pr["ok"]:
        raise PreviewError(pr["why"])
    drawio, out_dir = pathlib.Path(drawio), pathlib.Path(out_dir)
    if not drawio.exists():
        raise PreviewError(f"{drawio} が無い")
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="asis_mx_") as td:
        js = pathlib.Path(td) / "harness.js"
        js.write_text(HARNESS, encoding="utf-8")
        env = dict(font_env(), ASIS_PW_MODULE=pr["playwright"], ASIS_MXCLIENT=pr["mxclient"])   # 書体が無い機械は、代わりの和文の書体で
        try:
            r = subprocess.run([pr["node"], str(js), str(drawio), str(out_dir), str(scale), str(CAP)], capture_output=True, text=True, timeout=timeout, env=env)
        except subprocess.TimeoutExpired:
            raise PreviewError(f"{timeout} 秒で終わらない")
    if r.returncode != 0:
        raise PreviewError((r.stderr or r.stdout).strip()[-300:] or f"node が終了コード {r.returncode}")
    lines = [x for x in r.stdout.splitlines() if x.strip()]
    try:
        pages = json.loads(lines[-1])
    except (IndexError, ValueError):
        raise PreviewError("描画の結果が読めない: " + (r.stderr or r.stdout).strip()[-200:])
    return pages


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("drawio", nargs="?")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--scale", type=float, default=2.0)
    ap.add_argument("--probe", action="store_true")
    a = ap.parse_args()
    if a.probe or not (a.drawio and a.out):
        pr = probe(fresh=True)
        for k in ("node", "playwright", "mxclient", "browsers"):
            print(f"{k}: {pr[k] or '—'}")
        print("使える" if pr["ok"] else f"使えない: {pr['why']}")
        sys.exit(0 if pr["ok"] else 1)
    try:
        pages = render(a.drawio, a.out, a.scale)
    except PreviewError as e:
        sys.exit(f"下見を描けない: {e}")
    bad = 0
    for p in pages:
        if p.get("png"):
            print(p["png"])
        else:
            bad += 1
            print(f"p{p['index']:02d}［{p.get('name', '')}］: {p.get('error')}", file=sys.stderr)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
