#!/usr/bin/env python3
"""make_compare.py — 照合シートを作る。出典ページの画像と、そこを根拠にしたノードを並べた PDF（フローごとに1本）。

  python3 scripts/make_compare.py -w <作業フォルダ>              # out/照合シート_<フローID>.pdf
  python3 scripts/make_compare.py -w <作業フォルダ> --only F03
  python3 scripts/make_compare.py -w <作業フォルダ> --max-pages 8   # 1本の PDF に入れる出典ページの数（既定 12）。超えたら _1・_2… に分ける
  python3 scripts/make_compare.py -w <作業フォルダ> --only F03 --nodes P-003,P-007   # 差分照合のシート out/差分照合シート_F03_S-<n>.pdf
                                                                                   # （図全体の項目を照合し直すときは、番号に「図の項目」も並べる）
  python3 scripts/make_compare.py -w <作業フォルダ> --skip-reviewed   # 照合に出した図（採否表のある図）のシートは作らない（納品の門の書き出し直し）

1ページ＝出典の1ページ。左にフロー図（そのページを出典にしたノードを赤枠で囲む）とノードの一覧（引用つき）、
右に出典ページの画像。「このページに本当にそう書いてあるか」「このページの手順でフローに無いものはないか」を、
人と、作り方を知らない別のエージェント（references/source-review-prompt.md）が見るための紙。
最後に、このフローに割り当てたのに使わなかったページと理由、カードにあるのにフローに入れなかった手順と理由を載せる。
見出しのページ番号は、資料に刷られた番号（Excel はシート名）。出典ページが多いフローは、レビュー担当が1回で読み切れるように PDF を分ける。
作るたびに、照合に出す版（箱ごとの要約と図の印）を review/sheets.json に控え（review.record）、見出しの頭にシートの番号［S-<n>］を出す。
レビュー担当は、この番号を報告の sheet に写す（どの版の図を照合したかは、この番号で決まる）。図・種類・箱が、その図の最後の控えと同じなら、
前と同じ番号（作り直しでは増やさない）。書き出したあとに flow.json を直していれば、作らずに止まる（シートの図と、控える版をそろえる）。
差分照合（--nodes）: 照合のあとに変わった・足した箱（check_flow.py --stage review が挙げる）だけを照合し直すためのシート。
1ページ目に、照合し直す箱の一覧（図の中で赤枠）。続けて、その箱の出典ページだけ（赤枠と一覧も、その箱だけ）。使わなかったページの一覧は付けない。
差分照合のシートは、番号ごとに残す（照合した記録として、納品の門が作業の記録に写す）。
先に check_layout.py を実行しておく（フロー図の PNG を使う。draw.io が無い環境では out/下見/ の近似の画像を使う）。
"""
import argparse
import io
import pathlib
import re

from common import (NODE_LABEL, CardItems, cited_pages, docs_by_id, item_reasons, load, load_card, load_flows, load_page, manifest, page_json,
                    page_png, question_ids, src_label, unused_reasons, why_unused, work_dir)
from review import FIG

W, H = 1190, 842           # A3 横（pt）
MARGIN, TOP = 24, 44       # 余白と、見出しの下の y
LEFT_W, FLOW_H = 500, 440  # 左の欄の幅と、フロー図の高さ
CROP_MARGIN = 170          # 赤枠のノードのまわりに残す余白（図の座標）
BIG_MB = 10                # 照合シート1本が、これを超えたら知らせる（レビュー担当が Read で開ける大きさ）
RED = (0.8, 0.1, 0.1)
KIND = NODE_LABEL
JP_FONTS = ["/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "/Library/Fonts/Arial Unicode.ttf", "C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/YuGothR.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]
FONT = {"fontname": "japan"}                                         # 和文の書体ファイルが見つかれば差し替える（半角の字間が詰まる）


def pick_font():
    import fitz
    for f in JP_FONTS:
        if pathlib.Path(f).exists():
            try:
                t = fitz.open().new_page()
                t.insert_text(fitz.Point(10, 10), "試験 P-001", fontname="jp", fontfile=f, fontsize=9)
                FONT.clear()
                FONT.update({"fontname": "jp", "fontfile": f})
                return
            except Exception:
                continue


def put(page, x, y, text, size=9, color=(0, 0, 0), width=None, max_lines=3):
    """折り返して書き、次の y を返す"""
    import fitz
    per = max(8, int((width or 500) / (size * 0.95)))
    lines = []
    for para in str(text).split("\n"):
        while len(para) > per:
            lines.append(para[:per])
            para = para[per:]
        lines.append(para)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1][:-1] + "…"
    for ln in lines:
        page.insert_text(fitz.Point(x, y), ln, fontsize=size, color=color, **FONT)
        y += size * 1.35
    return y


def fit(rect_w, rect_h, w, h):
    k = min(rect_w / w, rect_h / h)
    return w * k, h * k, k


def flow_images(work):
    """フローのページ id → (配置, 書き出した PNG)。build_drawio.py の layout.json と check_layout.py の PNG から"""
    out = {}
    from validate import validate
    for lp in sorted((work / "out").glob("*.layout.json")):
        lay = load(lp)
        errs = validate(lay, "layout")
        if errs:
            raise SystemExit(f"{lp.name} の形が崩れている（{'／'.join(errs[:3])}）。build_drawio.py で書き出し直す")
        for pg in lay["pages"]:
            name = f"{lp.name[:-len('.layout.json')]}_p{pg['index']:02d}_{pg['id']}.png"
            png = work / "out" / "png" / name
            if not png.exists() and (work / "out" / "下見" / name).exists():   # draw.io が無い環境: 近似の下見（配置と同じ大きさなので赤枠は合う）
                png = work / "out" / "下見" / name
            out[pg["id"]] = (pg, png)
    return out


def nodes_by_page(fl):
    """出典のページ → そのページを出典にした [(ノード, 出典)]。ページは、フローに出てくる順"""
    by_page = {}
    for n in fl["nodes"]:
        for s in n.get("sources") or []:
            by_page.setdefault((s.get("doc_id"), s.get("page")), []).append((n, s))
    return by_page


class Sheets:
    """フロー1本の照合シート。出典ページ1枚につき1ページ（左上: フロー図、左下: ノードの一覧、右: 出典ページ）と、最後に使わなかったものの一覧"""

    def __init__(self, work, m, fl, layout, flow_im, questions, flows):
        self.work, self.m, self.fl, self.pg, self.flow_im = work, m, fl, layout, flow_im
        self.reasons, self.item_why, self.cited = unused_reasons(work, flows), item_reasons(flows), cited_pages(flows)
        self.F = fl["flow_id"]
        self.dmap = docs_by_id(m)
        self.q_node, _ = question_ids(questions, self.F)
        self.items_of = CardItems(work)
        self.sheet = ""                                              # 照合に出した版の番号（S-<n>。見出しの頭に出す）

    def src(self, s):
        return src_label(self.m, s)

    def source_page(self, page, k, uses):
        """出典ページ1枚分。k は (資料, ファイルの番号)、uses はそのページを出典にした [(ノード, 出典)]"""
        head = f"［{self.sheet}］{self.F} {self.fl['name']}　／　出典: {self.src({'doc_id': k[0], 'page': k[1]})}　{self.dmap.get(k[0], {}).get('rel_path', '')}"
        put(page, 24, 28, head, size=12, width=W - 48, max_lines=1)
        self.flow_crop(page, uses)
        self.node_list(page, k, uses)
        self.source_image(page, k, min([s["lines"][0] for _n, s in uses if isinstance(s.get("lines"), list) and s["lines"]] or [1]))

    def flow_crop(self, page, uses):
        """左上: フロー図。このページを出典にしたノードを赤枠で囲む。図が大きいときは、赤枠のまわりだけを切り出して大きく見せる"""
        import fitz
        pg, im, b = self.pg, self.flow_im, self.pg["bounds"]
        sx, sy = im.width / (b[2] - b[0]), im.height / (b[3] - b[1])
        rs = [pg["nodes"][n["id"]] for n, _s in uses if n["id"] in pg["nodes"]]
        cx0, cy0, cx1, cy1 = 0, 0, im.width, im.height
        if rs:
            x0 = max(0, (min(r[0] for r in rs) - b[0] - CROP_MARGIN) * sx)
            y0 = max(0, (min(r[1] for r in rs) - b[1] - CROP_MARGIN) * sy)
            x1 = min(im.width, (max(r[0] + r[2] for r in rs) - b[0] + CROP_MARGIN) * sx)
            y1 = min(im.height, (max(r[1] + r[3] for r in rs) - b[1] + CROP_MARGIN) * sy)
            if (x1 - x0) * (y1 - y0) < 0.6 * im.width * im.height:
                cx0, cy0, cx1, cy1 = int(x0), int(y0), int(x1), int(y1)
        part = im.crop((cx0, cy0, cx1, cy1))
        fw, fh, kf = fit(LEFT_W, FLOW_H, part.width, part.height)
        buf = io.BytesIO()
        part.resize((max(1, int(fw * 2)), max(1, int(fh * 2)))).save(buf, "PNG")
        page.insert_image(fitz.Rect(MARGIN, TOP, MARGIN + fw, TOP + fh), stream=buf.getvalue())
        for r in rs:
            x0, y0 = MARGIN + ((r[0] - b[0]) * sx - cx0) * kf, TOP + ((r[1] - b[1]) * sy - cy0) * kf
            page.draw_rect(fitz.Rect(x0 - 2, y0 - 2, x0 + r[2] * sx * kf + 2, y0 + r[3] * sy * kf + 2), color=RED, width=1.6)

    def node_list(self, page, k, uses):
        """左下: このページを出典にしたノードと、その引用"""
        y = TOP + FLOW_H + 22
        put(page, MARGIN, y - 16, f"開始: {self.fl.get('start_condition') or '（空）'}　／　終了: {self.fl.get('end_condition') or '（空）'}", size=8,
            color=(0.3, 0.3, 0.3), width=LEFT_W, max_lines=1)              # 図の情報欄が切り出しで見えないことがあるので、文字でも出す
        put(page, MARGIN, y - 6, "このページを出典にしたノード（番号・種別・ラベル ／「引用」）", size=9, color=(0.3, 0.3, 0.3), max_lines=1)
        y += 8
        items = self.items_of(*k)
        for n, s in uses:
            if y > H - 30:
                put(page, MARGIN, y, "…（続きは出典対応表を参照）", size=8, color=(0.4, 0.4, 0.4))
                break
            mark = ("【推定】" if n.get("inferred") else "") + ("【補足の中の作業】" if items.get(s.get("item"), {}).get("origin") == "note" else "")
            qs = "　" + ",".join(self.q_node.get(n["id"], [])) if self.q_node.get(n["id"]) else ""
            y = put(page, MARGIN, y, f"{n['id']}［{KIND.get(n['type'], n['type'])}］{mark}{n.get('label', '')}{qs}", size=9, width=LEFT_W, max_lines=2)
            y = put(page, MARGIN + 12, y, f"「{s.get('quote', '')}」" + ("（画像から読取）" if s.get("from_image") else ""), size=8,
                    color=(0.1, 0.25, 0.6), width=LEFT_W - 12, max_lines=2) + 4

    def source_image(self, page, k, first=1):
        """右: 出典ページの画像。画像の無いページ（マクロのソースなど）は文字を書く"""
        import fitz
        from PIL import Image
        x, w, h = MARGIN + LEFT_W + 20, W - (MARGIN + LEFT_W + 20) - MARGIN, H - TOP - MARGIN
        png = page_png(self.work, k[0], k[1]) if isinstance(k[1], int) else None
        if png is not None and png.exists():
            im = Image.open(png)
            iw, ih, _ = fit(w, h, im.width, im.height)
            page.insert_image(fitz.Rect(x, TOP, x + iw, TOP + ih), filename=str(png))
            page.draw_rect(fitz.Rect(x, TOP, x + iw, TOP + ih), color=(0.6, 0.6, 0.6), width=0.5)
            return
        pj = load_page(page_json(self.work, k[0], k[1])) if isinstance(k[1], int) else {}
        put(page, x, 58, "（このページは画像がない。文字だけ）", size=9, color=(0.4, 0.4, 0.4))
        y = 76
        if pj.get("code"):                                           # プログラムのソース: 引いた行の少し前から、モジュールの行の番号つきで
            at, base = max(0, first - 4), int(pj["code"].get("line_from") or 1)
            for i, ln in enumerate((pj.get("text") or [])[at:at + 68], at):
                y = put(page, x, y, f"{base + i:>4}| {ln}", size=8, width=w, max_lines=1)
            return
        lines = list(pj.get("text") or []) + [s["text"] for s in pj.get("shapes") or [] if isinstance(s, dict) and s.get("text")]
        for ln in lines[:70]:                                        # 非表示のスライドは、文字が図形の側にある
            y = put(page, x, y, ln, size=8, width=w, max_lines=1)

    def diff_page(self, page, targets):
        """差分照合の1ページ目: 照合し直す箱（図の中で赤枠）と、その名前・レーン・出典・前後の線"""
        nodes = [n for n in self.fl["nodes"] if n["id"] in targets]
        lanes = {x.get("id"): x.get("name") for x in self.fl.get("lanes") or []}
        put(page, MARGIN, 28, f"［{self.sheet}］{self.F} {self.fl['name']}　／　差分照合: 照合のあとに変わった・足した箱（この箱だけを判定する）", size=12, width=W - 2 * MARGIN, max_lines=1)
        self.flow_crop(page, [(n, {}) for n in nodes])
        x = MARGIN + LEFT_W + 20
        y = put(page, x, TOP + 12, f"照合のあとに図を直したので、赤枠の箱だけを照合し直す。ほかの箱は、前の照合で済んでいる。報告の sheet に「{self.sheet}」と書く。",
                size=10, width=W - x - MARGIN)
        put(page, x, y + 6, "次のページから、この箱の出典ページ（赤枠と一覧も、この箱だけ）。箱の前後の線（つなぎ）も、原典どおりかを見る。出典の無い箱（推定）は、図とフローの定義で見る。",
            size=10, width=W - x - MARGIN, max_lines=4)
        y = TOP + FLOW_H + 22
        if FIG in targets:
            fl = self.fl
            y = put(page, MARGIN, y, f"{FIG}（図全体）: 開始条件「{fl.get('start_condition') or '（空）'}」／終了条件「{fl.get('end_condition') or '（空）'}」。"
                    "使わないページ・対象外にした手順の理由は、最後のページ。原典どおりかを見る（間違いは、報告の findings に書く）",
                    size=9, width=W - 2 * MARGIN, max_lines=3) + 4
        for n in nodes:
            if y > H - 30:
                put(page, MARGIN, y, "…（続きは、次のページからの一覧）", size=8, color=(0.4, 0.4, 0.4))
                break
            srcs = "・".join(dict.fromkeys(self.src(s) for s in n.get("sources") or [])) or "出典なし（推定）"
            near = "　".join([f"→{e['to']}{'（' + e['label'] + '）' if e.get('label') else ''}" for e in self.fl.get("edges") or [] if e.get("from") == n["id"]]
                            + [f"←{e['from']}" for e in self.fl.get("edges") or [] if e.get("to") == n["id"]])
            y = put(page, MARGIN, y, f"{n['id']}［{KIND.get(n['type'], n['type'])}］{'【推定】' if n.get('inferred') else ''}{n.get('label', '')}　レーン: {lanes.get(n.get('lane'), '')}",
                    size=9, width=W - 2 * MARGIN, max_lines=2)
            y = put(page, MARGIN + 12, y, f"出典: {srcs}　線: {near}", size=8, color=(0.1, 0.25, 0.6), width=W - 2 * MARGIN - 12, max_lines=2) + 4

    def leftover_page(self, page, unused, pages=()):
        """最後のページ: 割り当てたのに出典にしなかったページと、カードにあるのにフローに入れなかった手順・分岐（どちらも理由つき）。
        pages（このフローのページ）のカードで、作業として扱わなかった補足（notes の why_not_action）も載せる"""
        fl, wide = self.fl, W - 2 * MARGIN
        put(page, MARGIN, 28, f"［{self.sheet}］{self.F} {fl['name']}　／　この業務分類に割り当てたが、ノードの出典にしなかったページ", size=12, width=wide, max_lines=1)
        y = 56
        if not unused:
            put(page, MARGIN, y, "なし（割り当てたページはすべて、どれかのノードの出典になっている）", size=10)
        for k in unused:
            others = [f for f in self.cited.get(k, []) if f != self.F]       # 点検（check_flow.py）と同じ見方で書く
            why = (f"ほかのフロー（{'・'.join(others)}）の出典になっている" if others
                   else why_unused(k, self.reasons, self.item_why) or "理由の記録なし（check_flow.py が FAIL にする）")
            y = put(page, MARGIN, y, f"{self.src({'doc_id': k[0], 'page': k[1]})}　{self.dmap.get(k[0], {}).get('rel_path', '')}　— {why}",
                    size=10, width=wide, max_lines=2) + 4
        y += 20
        put(page, MARGIN, y, "カードにあるのに、フローに入れなかった手順・分岐（理由）", size=12, width=wide, max_lines=1)
        y += 26
        ex = fl.get("excluded_items") or []
        if not ex:
            put(page, MARGIN, y, "なし", size=10)
        for x in ex:
            it = self.items_of(x.get("doc_id"), x.get("page")).get(x.get("item"), {})
            y = put(page, MARGIN, y, f"{self.src(dict(x, quote=it.get('quote', '')))} {x.get('item')}「{it.get('text', '')}」"
                    + ("【補足の中の作業】" if it.get("origin") == "note" else "") + f"　— {x.get('reason', '')}", size=10, width=wide, max_lines=2) + 4
        y += 20
        put(page, MARGIN, y, "読取のときに、作業として扱わなかった補足（理由）", size=12, width=wide, max_lines=1)
        y += 26
        held = [(k, n) for k in sorted(pages, key=str) for n in (load_card(self.work, *k) or {}).get("notes") or []
                if isinstance(n, dict) and str(n.get("why_not_action") or "").strip()]
        if not held:
            put(page, MARGIN, y, "なし", size=10)
        for k, n in held:
            y = put(page, MARGIN, y, f"{self.src({'doc_id': k[0], 'page': k[1], 'quote': n.get('quote', '')})}「{n.get('quote', '')}」　— {n.get('why_not_action', '')}",
                    size=9, width=wide, max_lines=2) + 3


def write_sheets(work, sheets, assigned, max_pages, targets=None):
    """フロー1本の照合シートを、出典 max_pages ページごとの PDF に分けて書く（レビュー担当が1回で読み切れる量にする）。
    targets（箱の番号）があれば、差分照合のシート（その箱と、その箱の出典ページだけ）"""
    import fitz
    F, by_page = sheets.F, nodes_by_page(sheets.fl)
    full = set(by_page)
    if targets:
        by_page = {k: [(n, s) for n, s in uses if n["id"] in targets] for k, uses in by_page.items()}
        by_page = {k: uses for k, uses in by_page.items() if uses}
    order = list(by_page)
    parts = [order[i:i + max_pages] for i in range(0, max(len(order), 1), max_pages)]
    pdfs = [fitz.open() for _ in parts]
    if targets:
        sheets.diff_page(pdfs[0].new_page(width=W, height=H), targets)
    for part, pdf in zip(parts, pdfs):
        for k in part:
            sheets.source_page(pdf.new_page(width=W, height=H), k, by_page[k])
    if not targets or FIG in targets:                                # 図の項目（使わないページ・対象外の項目の理由）を照合し直すときも、最後のページに載せる
        sheets.leftover_page(pdfs[-1].new_page(width=W, height=H), [k for k in assigned if k not in full], pages=set(assigned) | full)
    name = f"差分照合シート_{F}_{sheets.sheet}" if targets else f"照合シート_{F}"
    for old in (work / "out").glob(f"{name}*.pdf"):                  # 前に作った分（分け方が変わると名前が変わる）。下位フロー F03-1 の分・ほかの差分照合のシートは消さない
        if re.fullmatch(rf"{re.escape(name)}(_\d+)?", old.stem):
            old.unlink()
    for i, pdf in enumerate(pdfs, 1):
        out = work / "out" / (f"{name}.pdf" if len(pdfs) == 1 else f"{name}_{i}.pdf")
        try:                                                         # 和文の書体を、使った字だけにする（しないと、ページごとに書体が丸ごと入って 1 ページ 7〜10 MB になる）
            pdf.subset_fonts()
        except Exception as e:
            print(f"  ※ 書体を小さくできなかった（{type(e).__name__}。fontTools が要る）。ファイルが大きくなる")
        pdf.save(str(out), garbage=4, deflate=True)
        mb = out.stat().st_size / 1e6
        print(f"{out}  （出典 {len(parts[i - 1])} ページ" + (f"・{i}/{len(pdfs)}" if len(pdfs) > 1 else "") + f"・{mb:.1f} MB）"
              + (f"  ※ {BIG_MB} MB を超える。レビュー担当が開けないときは --max-pages を小さくする" if mb > BIG_MB else ""))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only")
    ap.add_argument("--max-pages", type=int, default=12, help="1本の PDF に入れる出典ページの数（レビュー担当が1回で読み切れる量）")
    ap.add_argument("--nodes", help="差分照合: 照合し直す箱の番号（カンマ区切り。--only に 1 本だけ）。out/差分照合シート_<フローID>_S-<n>.pdf を作る")
    ap.add_argument("--skip-reviewed", action="store_true",
                    help="照合に出した図（採否表のある図）のシートは作らない（作り直すと、照合していない版が控えられ、照合したシートが消える。納品の門が使う）")
    a = ap.parse_args()
    targets = [x.strip() for x in a.nodes.split(",") if x.strip()] if a.nodes else None
    if targets is not None and (not a.only or "," in a.only or not targets):
        ap.error("--nodes は、--only にフローを 1 本だけ指定して使う（箱の番号も 1 つ以上）")
    from PIL import Image
    import review as RV
    pick_font()
    work = work_dir(a.work)
    m = manifest(work)
    questions = (load(work / "questions.json", {}) or {}).get("questions", [])
    cat_pages = {c["id"]: [(ps["doc_id"], p) for ps in c.get("pages") or [] for p in ps.get("pages") or []]
                 for c in (load(work / "taxonomy.json", {}) or {}).get("categories") or []}
    images = flow_images(work)
    if any("下見" in png.parts for _pg, png in images.values()):
        print("フロー図の PNG が無いので、下見の画像（mxGraph の近似）でシートを作る（納品の PNG は draw.io のある環境で check_layout.py）")
    only = set(a.only.split(",")) if a.only else None
    flows, skipped = load_flows(work), []
    for fl in flows:
        F = fl["flow_id"]
        if (only and F not in only) or (a.skip_reviewed and (work / "review" / f"{F}.review.json").exists()):
            continue
        layout, png = images.get(f"flow-{F}", (None, None))
        if layout is None or not png.exists():
            print(f"{F}: フロー図の PNG が無い（build_drawio.py → check_layout.py を先に実行する）")
            skipped.append(F)
            continue
        if layout.get("mark") != RV.figure_mark(fl):                 # 書き出したあとに flow.json を直した（シートの図と、控える版が食い違う）
            print(f"{F}: 書き出した図が、いまの flow.json と違う（書き出したあとに直した）。build_drawio.py → check_layout.py で書き出し直す")
            skipped.append(F)
            continue
        if targets:
            have = {n["id"]: n for n in fl["nodes"]}
            bad = [x for x in targets if x != FIG and (x not in have or have[x]["type"] in ("start", "end"))]
            if bad:
                raise SystemExit(f"--nodes の番号が、{F} の箱に無い（開始・終了は照合しない）: {'・'.join(bad)}")
        sheets = Sheets(work, m, fl, layout, Image.open(png).convert("RGB"), questions, flows)
        try:
            sheets.sheet = RV.record(work, {F: fl}, "差分" if targets else "全体", targets)[F]   # 照合に出す版を控える（報告の sheet と結ぶ）
        except ValueError as e:
            raise SystemExit(f"{e}。消さずに、ファイルを前の版に戻す（照合に出した版の記録は、作り直せない）")
        write_sheets(work, sheets, cat_pages.get(F, []), a.max_pages, targets)
        if targets:
            print(f"{F}: シート {sheets.sheet}。差分照合の報告の置き場は {RV.next_report(work, F)}（sheet に「{sheets.sheet}」と書かせる）")
    ids = {fl["flow_id"] for fl in flows}
    for old in sorted((work / "out").glob("*照合シート_*.pdf")) if not only else []:   # 消えた・番号を変えたフローのシート（残すと、照合に出してしまう）
        if re.sub(r"(_S-\d+)?(_\d+)?$", "", old.stem.split("照合シート_", 1)[1]) not in ids:
            old.unlink()
            print(f"{old.name}: いまの業務フローに無いので消した")
    if skipped:
        raise SystemExit(f"照合シートを作れなかったフロー {len(skipped)} 本（{'・'.join(skipped)}）。build_drawio.py → check_layout.py で PNG を作ってから、もう一度")


if __name__ == "__main__":
    main()
