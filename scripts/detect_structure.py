#!/usr/bin/env python3
"""detect_structure.py — 入力が階層構造かどうかを判定するための手がかりを、機械的に集める。

  python3 scripts/detect_structure.py -w <作業フォルダ>

集めるもの（structure/signals.json に保存し、要約を表示する）:
  フォルダの階層・フォルダ名とファイル名の番号／PDF のしおり・PowerPoint のセクション・Excel のシート名／
  目次らしいページと、そこから機械的に拾えた項目／各ページの見出しと見出し番号の深さ／
  本文の中の、番号で始まる短い行（第◯章・1・1.1・（1）・①）＝本文の見出しの候補と、番号の書き方から見立てた段／
  他の資料を指す言い回し（別紙・◯◯参照・資料名の言及）＝資料どうしの関係の候補／資料の役割の見立て／同一ファイル・版違いらしい組
これは手がかりで、判定ではない。判定（structure/hierarchy.json）は、この結果と目次ページの画像を見て LLM が書く
（references/structure-guide.md）。画像の中にしかない目次は、ここでは拾えない。
"""
import argparse
import re

from common import ROLE_LABEL, cover_title, file_stem, load, load_card, manifest, norm, page_text, save, work_dir

KANJI_NUM = "一二三四五六七八九十"
H_RULES = [
    (re.compile(rf"^第\s*([0-9０-９{KANJI_NUM}]+)\s*[部編]"), lambda m: 1),
    (re.compile(rf"^第\s*([0-9０-９{KANJI_NUM}]+)\s*章"), lambda m: 1),
    (re.compile(rf"^第\s*([0-9０-９{KANJI_NUM}]+)\s*[節項]"), lambda m: 2),
    (re.compile(r"^([0-9０-９]+(?:[\.\-－．][0-9０-９]+)*)[\.．、:：\s]\s*\S"), lambda m: len(re.split(r"[\.\-－．]", m.group(1)))),
    (re.compile(r"^[（(]\s*[0-9０-９]+\s*[)）]"), lambda m: 3),
    (re.compile(r"^[①-⑳]"), lambda m: 4),
]
NUM_STYLES = [                                                       # 目次の項目の番号の書き方 → 段を決めるのに使う
    ("chapter", re.compile(rf"^(第\s*[0-9０-９{KANJI_NUM}]+\s*[部編章])\s*")),
    ("section", re.compile(rf"^(第\s*[0-9０-９{KANJI_NUM}]+\s*[節項])\s*")),
    ("num", re.compile(r"^([0-9０-９]+(?:[\.\-－．][0-9０-９]+)*)(?:[\.．、:：]\s*|\s+|$)")),
    ("paren", re.compile(r"^([（(]\s*[0-9０-９]+\s*[)）])\s*")),
    ("circled", re.compile(r"^([①-⑳])\s*")),
]
ROLE_RULES = [                                                      # 資料の役割の見立て（名前とフォルダから。決めるのは LLM）
    ("regulation", re.compile(r"規程|規則|規定|要領|要綱|細則|基準|方針|ポリシー")),
    ("appendix", re.compile(r"別紙|別添|別表|様式|付録|チェックシート|チェックリスト|テンプレート|ひな形|雛形|フォーマット")),
    ("procedure", re.compile(r"手順|操作|作業標準|SOP|ガイド|手引")),
    ("manual", re.compile(r"マニュアル|業務フロー|業務一覧|全体")),
]
SHOW_REFS, SHOW_TOC, SHOW_HEAD = 30, 150, 120                        # 画面に出す件数（全部は signals.json）
HEAD_ORDER = ["chapter", "section", "num1", "num2", "num3", "num4", "paren", "circled"]   # 本文の見出しの段: 番号の書き方の、ふつうの入れ子の順
HEAD_LEN, HEAD_TOP = 40, 5                                           # 見出しの候補にする行の長さの上限／柱（ページの上で繰り返す行）を探す行数
HEAD_UNIT = re.compile(r"^[年月日時分秒円件名個回人枚部％%](?:[\s　0-9０-９、。)）]|$)|^[〜~\-－/／:：)）]")   # 「3 月 31 日」「10 年」は見出しでない
TOC_TITLE = re.compile(r"(目\s*次|もくじ|contents|index|索\s*引|業務一覧|フロー一覧|マニュアル一覧)", re.I)
TOC_LINE = re.compile(r"^(.{2,60}?)[\s　…‥\.．・]{2,}\s*([0-9０-９]{1,4})\s*$")
REF_RULES = [
    re.compile(r"(別紙|別添|別表|付録|様式|別冊)\s*[0-9０-９A-Za-zＡ-Ｚ一二三四五六七八九十\-－]*"),
    re.compile(r"「([^」]{2,40})」\s*(?:を|に|の)?[^。\n]{0,10}参照"),
    re.compile(r"([^\s、。「」（）()]{2,30}(?:マニュアル|手順書|要領|規程|規則|細則|ガイド|基準書|フロー図))"),
    re.compile(r"(?:詳細|詳しく)は[^。\n]{2,40}(?:参照|を見る|をご覧)"),
]
NUM_PREFIX = re.compile(r"^([0-9０-９]{1,3}(?:[\-_－][0-9０-９]{1,3})*)[\s_\-－\.．)）]")


def heading_of(pj):
    """ページの見出しの候補（PowerPoint はタイトル、ほかは最初の短い行）"""
    if pj.get("title"):
        return pj["title"].split("\n")[0].strip()
    for ln in (pj.get("text") or []) + (pj.get("cells") or []):
        ln = ln.strip()
        if 2 <= len(ln) <= 50:
            return ln
    return ""


def heading_level(text):
    for rx, lv in H_RULES:
        m = rx.match(text)
        if m:
            return lv(m), m.group(0).strip()
    return None, ""


def number_style(text):
    """行の頭の番号 → (書き方, 番号, 残りの文字)。番号が無ければ (None, "", text)"""
    for name, rx in NUM_STYLES:
        m = rx.match(text)
        if m:
            num = m.group(1).strip()
            style = name + (str(len(re.split(r"[\.\-－．]", num))) if name == "num" else "")
            return style, num, text[m.end():].strip()
    return None, "", text


def body_headings_of(lines):
    """本文の行のうち、番号で始まる短い行（見出しの候補）: [(行の位置, 書き方, 番号, 題, 行の文字)]。文（。で終わる・途中に 。がある）と、数量（「3 月」）は除く"""
    out, seen = [], set()
    for i, ln in enumerate(lines):
        ln = re.sub(r"[\s　]+", " ", ln).strip()
        if not 3 <= len(ln) <= HEAD_LEN or "。" in ln or ln[-1] in "、，,":
            continue
        style, num, title = number_style(ln)
        if not style or not title or HEAD_UNIT.match(title) or style not in HEAD_ORDER:
            continue
        if style.startswith("num") and not num.isdecimal() and not re.search(r"[\.\-－．]", num):
            continue
        if style == "num1" and int(num) > 99:                        # 「2020 年度」「100 万円」
            continue
        if norm(ln) not in seen:                                     # 文字層と図形の両方から同じ行が来る
            seen.add(norm(ln))
            out.append((i, style, num, title, ln))
    return out


def toc_entries_of(lines):
    """目次の行（題 … ページ）を拾う。番号が前の行に分かれている組版（「第1 章」「（1）」だけの行）にも対応する"""
    out, prev = [], ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        mm = TOC_LINE.match(ln)
        if mm:
            title, style, num = mm.group(1).strip(), None, ""
            ps, pn, prest = number_style(prev)
            if ps and not prest:                                     # 前の行が番号だけ
                style, num = ps, pn
            else:
                st, n, rest = number_style(title)
                if st and len(rest) >= 2:
                    style, num, title = st, n, rest
            out.append({"title": title, "page": int(norm(mm.group(2))), "style": style, "number": num})
        prev = ln
    return out


def toc_offset(entries, pages, after):
    """目次のページ番号と、資料の実際のページとのずれ。題が見出しとして出てくるページを探し、いちばん多い差を返す"""
    diffs = {}
    top = [e for e in entries if e.get("level") in (1, 2)]           # 下の段の項目は、章の頭の一覧にも出てくるので当てにしない
    for e in (top if len(top) >= 3 else entries):
        key = norm(re.split(r"[（(]", e["title"])[0])
        if len(key) < 4:
            continue
        for pno, lines in pages:
            if pno <= after:
                continue
            if any(norm(re.split(r"[（(]", number_style(ln.strip())[2])[0]) == key for ln in lines[:40]):
                diffs[pno - e["page"]] = diffs.get(pno - e["page"], 0) + 1
                break
    if not diffs:
        return None
    best = max(diffs, key=diffs.get)
    return best if diffs[best] >= 3 and diffs[best] * 2 >= sum(diffs.values()) else None


def name_key(name):
    """資料名を本文の中から探すための幹（番号の頭・拡張子・版の表記を落とす）"""
    stem = re.sub(r"\.[A-Za-z0-9]+$", "", name)
    stem = NUM_PREFIX.sub("", stem)
    stem = re.sub(r"(?i)[_\-\s（(]*(v(er)?\.?\s*\d+(\.\d+)*|第?\d+版|\d{6,8}|最新|旧|改訂|コピー|final|draft)[)）]*$", "", stem)
    return norm(stem)


def scan_doc(work, d, keys, cross):
    """資料1件から、階層の手がかり（見出し・目次・しおり・セクション・シート）を集める。参照の言い回しは cross に足す"""
    info = {"doc_id": d["doc_id"], "name": d["name"], "folder": d["folder"], "kind": d["kind"], "pages": d.get("pages"),
            "number_prefix": (NUM_PREFIX.match(d["name"]) or [None, None])[1],
            "outline": d.get("outline") or [], "sections": d.get("sections") or [], "sheets": d.get("sheets") or [],
            "hidden_sheets": d.get("hidden_sheets") or [],
            "vba": [{"module": v["name"], "procedures": [p["name"] for p in v["procedures"]]} for v in d.get("vba_modules") or []],
            "toc_pages": [], "toc_entries": [], "headings": [], "body_headings": [], "render_note": d.get("render_note", "")}
    levels, page_lines, prev_top = set(), [], set()
    for p in d.get("page_list") or []:
        pj = load(work / p["json"], {})
        if p["kind"] == "vba":
            continue
        head = heading_of(pj)
        lv, num = heading_level(head)
        if head:
            info["headings"].append({"page": p["page"], "text": head[:60], "level": lv, "number": num})
        if lv:
            levels.add(lv)
        lines = (pj.get("text") or []) + (pj.get("cells") or []) + [s.get("text", "") for s in pj.get("shapes") or []]
        lines = [ln for blk in lines for ln in str(blk).split("\n")]
        page_lines.append((p["page"], lines))
        entries = toc_entries_of(lines)
        is_toc = bool(TOC_TITLE.search(head) or len(entries) >= 4)
        for i, style, num, title, ln in [] if is_toc or d["kind"] == "xlsx" else body_headings_of(lines):   # 表計算の番号つきの行は、表の行（見出しはシート名で見る）
            if not (i < HEAD_TOP and norm(ln) in prev_top):          # 前のページの上にもある行は柱（章名の繰り返し）。初めて出たページだけ残す
                info["body_headings"].append({"page": p["page"], "text": ln, "style": style, "number": num, "title": title, "level": None})
        prev_top = {norm(ln) for ln in lines[:HEAD_TOP]}
        if is_toc:
            info["toc_pages"].append(p["page"])
            for e in entries:
                e = dict(e, toc_page=p["page"], level=None)
                if e not in info["toc_entries"]:                     # 文字層と図形の両方から同じ行が来る
                    info["toc_entries"].append(e)
        text = page_text(pj)
        for rx in REF_RULES:
            for mm in rx.finditer(text):
                phrase = mm.group(0).strip()[:60]
                if not any(c["doc_id"] == d["doc_id"] and c["page"] == p["page"] and c["text"] == phrase for c in cross):
                    cross.append({"doc_id": d["doc_id"], "page": p["page"], "text": phrase, "to_doc": None})
        nt = norm(text)
        for other, key in keys.items():
            if other != d["doc_id"] and len(key) >= 4 and key in nt:
                cross.append({"doc_id": d["doc_id"], "page": p["page"], "text": f"資料名の言及: {key}", "to_doc": other})
    info["heading_levels"] = sorted(levels)
    order = []                                                       # 目次の段: 番号の書き方が初めて出た順（第◯章 → 1 → （1） なら 1・2・3 段）
    for e in info["toc_entries"]:
        if e["style"] and e["style"] not in order:
            order.append(e["style"])
    for e in info["toc_entries"]:
        e["level"] = order.index(e["style"]) + 1 if e["style"] else None
    info["toc_levels"] = len(order)
    styles = [x for x in HEAD_ORDER if any(h["style"] == x for h in info["body_headings"])]
    for h in info["body_headings"]:                                  # 段は、この資料に出てくる書き方を、ふつうの入れ子の順に並べた何番目か（見立て）
        h["level"] = styles.index(h["style"]) + 1
    info["body_heading_styles"] = styles
    info["toc_page_offset"] = toc_offset(info["toc_entries"], page_lines, max(info["toc_pages"])) if info["toc_pages"] else None
    info["role_guess"] = role_guess(d)
    info["title_guess"] = title_guess(work, d)
    return info


def title_guess(work, d):
    """資料の題名の見立て: 表紙のカードの見出し（カードがあれば）、無ければファイル名（拡張子なし）。決めるのは LLM（表紙の画像を見て docs[].title に写す）"""
    return cover_title(load_card(work, d["doc_id"], p["page"]) for p in d.get("page_list") or []) or file_stem(d.get("name"))


def role_guess(d):
    """資料の役割の見立て（資料名、無ければフォルダ名から。マクロのあるものはツール）"""
    if d.get("has_macro"):
        return "tool"
    path = "/".join(d["folder"]) + "/" + d["name"]
    return next((k for k, rx in ROLE_RULES if rx.search(d["name"])), None) or next((k for k, rx in ROLE_RULES if rx.search(path)), "other")


def settle_cross(cross, docs, keys):
    """参照の言い回しから、柱（ヘッダー・フッター）を除き、言い回しの中に資料名があれば行き先を当てる"""
    seen_on = {}                                                     # 多くのページに同じ文言が出るのはヘッダー・フッター（資料名の柱）。参照ではない
    for c in cross:
        seen_on.setdefault((c["doc_id"], c["text"]), set()).add(c["page"])
    n_pages = {d["doc_id"]: d.get("pages") or 0 for d in docs}
    cross = [c for c in cross if c["to_doc"] or len(seen_on[(c["doc_id"], c["text"])]) <= max(3, n_pages[c["doc_id"]] * 0.3)]
    for c in cross:
        if c["to_doc"] is None:
            t = norm(c["text"])
            for other, key in keys.items():
                if other != c["doc_id"] and len(key) >= 4 and (key in t or (len(t) >= 4 and t in key)):
                    c["to_doc"] = other
                    break
    return cross


def make_hints(m, infos, cross, folders, folder_depth):
    """階層の判定に使う手がかりを、人が読む文にする"""
    numbered_folders = sum(1 for f in folders if f != "（直下）" and all(NUM_PREFIX.match(seg + " ") or NUM_PREFIX.match(seg) for seg in f.split("/")))
    hints = []
    if folder_depth >= 1:
        hints.append(f"フォルダが {folder_depth} 階層（{len(folders)} フォルダ）" + ("。番号付きのフォルダ名あり" if numbered_folders else ""))
    if any(x["toc_pages"] for x in infos):
        hints.append("目次らしいページ: " + "、".join(f"{x['doc_id']}/" + "・".join(f"p{n:04d}" for n in x["toc_pages"]) for x in infos if x["toc_pages"])
                     + "（画像を見て確かめる。項目が業務フローの一覧になっているかを判定する）")
    else:
        hints.append("文字から拾える目次は無い（画像の中にある場合は、表紙の次あたりのページを見て確かめる）")
    for x in infos:
        if x.get("toc_levels", 0) >= 2:
            hints.append(f"{x['doc_id']} の目次は {x['toc_levels']} 段（番号の書き方から）。どの段が業務（フローの単位）かを画像で確かめる")
        if x.get("toc_page_offset"):
            off = x["toc_page_offset"]
            hints.append(f"{x['doc_id']} の目次の番号は、ファイルの番号（pNNNN）と {off:+d} ずれている（目次の N ＝ ファイルの N{off:+d}。表紙などを数えていない）。"
                         f"ページの割当にはファイルの番号を書く。読取のあと pages.py pagemap が対応を出す")
    for x in infos:
        if x.get("body_headings"):
            names = {"chapter": "第◯章", "section": "第◯節", "num1": "1", "num2": "1.1", "num3": "1.1.1", "num4": "1.1.1.1", "paren": "（1）", "circled": "①"}
            hints.append(f"{x['doc_id']} の本文に、番号で始まる短い行（見出しの候補）が {len(x['body_headings'])} 件（書き方: "
                         + " → ".join(names[s] for s in x["body_heading_styles"]) + "）。図の箱の題・箇条書きも混じる。"
                         + ("目次が拾えていないので、下の一覧を画像で確かめて、章・節の見出しを階層に使う" if not (x["toc_entries"] or x["outline"])
                            else "目次と食い違う見出し・目次に無い見出しを探すときに signals.json の body_headings を見る"))
    deep = [x for x in infos if len(x["heading_levels"]) >= 2]
    if deep:
        hints.append("見出し番号が2階層以上ある資料: " + "、".join(f"{x['doc_id']}（{x['heading_levels']}）" for x in deep))
    if any(x["outline"] for x in infos):
        hints.append("PDF のしおりがある資料: " + "、".join(x["doc_id"] for x in infos if x["outline"]))
    if any(x["sections"] for x in infos):
        hints.append("PowerPoint のセクションがある資料: " + "、".join(x["doc_id"] for x in infos if x["sections"]))
    linked = sorted({(c["doc_id"], c["to_doc"]) for c in cross if c["to_doc"]})
    if linked:
        hints.append("資料間の参照: " + "、".join(f"{a}→{b}" for a, b in linked))
    vg = sorted({tuple(d["version_group"]) for d in m["docs"] if d.get("version_group")})
    if vg:
        hints.append("版違いらしい組: " + "、".join("/".join(g) for g in vg) + "（どれが現行かを確認事項に出す）")
    dup = [(d["doc_id"], d["duplicate_of"]) for d in m["docs"] if d.get("duplicate_of")]
    if dup:
        hints.append("同一ファイル: " + "、".join(f"{a}={b}" for a, b in dup))
    unread = [d for d in m["docs"] if not d.get("duplicate_of") and not d.get("page_list")]
    if unread:
        hints.append("読めていない資料: " + "、".join(f"{d['doc_id']}（{d.get('render_note') or d.get('note') or '未描画'}）" for d in unread))
    return hints, [list(g) for g in vg], dup


def print_summary(sig):
    """親（進行役）が読む要約。全部は signals.json にある"""
    infos = {x["doc_id"]: x for x in sig["docs"]}
    print("■ フォルダと資料")
    for f, ids in sig["folders"].items():
        print(f"  {f}")
        for i in ids:
            x = infos[i]
            print(f"    {i} {x['name']}  {x['pages']} ページ  役割の見立て: {ROLE_LABEL[x['role_guess']]}  題名の見立て: {x.get('title_guess', '')}"
                  + (f"  ※{x['render_note']}" if x["render_note"] else ""))
    print("\n■ 手がかり")
    for h in sig["hints"]:
        print("  ・" + h)
    cand = [c for c in sig["cross_refs"] if c["to_doc"]]
    if cand:
        print(f"\n■ 資料どうしの関係の候補（{len(cand)} 件。根拠の文を読んで採否を決める。根拠の無い関係は書かない）")
        for c in cand[:SHOW_REFS]:
            print(f"    {c['doc_id']}/p{c['page']:04d} → {c['to_doc']}｜「{c['text'][:50]}」")
        if len(cand) > SHOW_REFS:
            print(f"    …（残り {len(cand) - SHOW_REFS} 件は signals.json の cross_refs）")
    for x in sig["docs"]:
        if x["toc_entries"]:
            print(f"\n■ {x['doc_id']} の目次から機械的に拾えた項目（{len(x['toc_entries'])} 件。画像で確かめる）")
            for e in x["toc_entries"][:SHOW_TOC]:
                print(f"    {'  ' * ((e['level'] or 1) - 1)}{(e['number'] + ' ') if e['number'] else ''}{e['title']} … {e['page']}")
            if len(x["toc_entries"]) > SHOW_TOC:
                print(f"    …（残り {len(x['toc_entries']) - SHOW_TOC} 件は signals.json）")
        elif x["outline"]:
            print(f"\n■ {x['doc_id']} のしおり")
            for e in x["outline"][:40]:
                print(f"    {'  ' * (e['level'] - 1)}{e['title']} … {e['page']}")
        if x.get("body_headings") and not (x["toc_entries"] or x["outline"]):
            print(f"\n■ {x['doc_id']} の本文の見出しの候補（{len(x['body_headings'])} 件。番号で始まる短い行。段は番号の書き方からの見立て。画像で確かめる）")
            for h in x["body_headings"][:SHOW_HEAD]:
                print(f"    p{h['page']:04d} {'  ' * (h['level'] - 1)}{h['text']}")
            if len(x["body_headings"]) > SHOW_HEAD:
                print(f"    …（残り {len(x['body_headings']) - SHOW_HEAD} 件は signals.json の body_headings）")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    a = ap.parse_args()
    work = work_dir(a.work)
    m = manifest(work)
    docs = [d for d in m["docs"] if not d.get("duplicate_of")]
    keys = {d["doc_id"]: name_key(d["name"]) for d in docs}          # 資料 → 本文の中から資料名を探すための幹
    cross = []
    infos = [scan_doc(work, d, keys, cross) for d in docs]
    cross = settle_cross(cross, docs, keys)
    folder_depth = max([len(d["folder"]) for d in docs] + [0])
    folders = {}
    for d in docs:
        folders.setdefault("/".join(d["folder"]) or "（直下）", []).append(d["doc_id"])
    hints, version_groups, dup = make_hints(m, infos, cross, folders, folder_depth)
    sig = {"folder_depth": folder_depth, "folders": folders, "docs": infos, "cross_refs": cross,
           "version_groups": version_groups, "duplicates": dup, "hints": hints}
    save(work / "structure" / "signals.json", sig)
    print(f"signals → {work / 'structure' / 'signals.json'}\n")
    print_summary(sig)
    print("\n次: 目次ページ（と表紙の次のページ）の画像を見て、structure/hierarchy.json を書く（references/structure-guide.md）。"
          "役割・題名の見立てと関係の候補は機械的な当たり。画像と文で確かめてから書く（題名は、表紙の題名を写す。納品物では、資料を題名で書く）")


if __name__ == "__main__":
    main()
