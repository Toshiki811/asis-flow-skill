#!/usr/bin/env python3
"""asis-flow-skill の各スクリプトが共有する小道具（作業フォルダの読み書き・文字の照合・結果の出し方）。

作業フォルダ（-w で渡す）の中身:
  manifest.json            入力の棚卸し（inventory.py）。ページ単位の進捗もここ
  pages/<資料ID>/pNNNN.png  ページ画像（render_pages.py）
  pages/<資料ID>/pNNNN.json 機械抽出（文字・表・図形・コネクタ）
  structure/signals.json   階層の手がかり（detect_structure.py）
  structure/hierarchy.json 階層の判定（LLM が書く）
  cards/<資料ID>/pNNNN.json ページカード（LLM が画像を読んで書く）
  glossary.json            用語の正本（LLM が書く）
  taxonomy.json            業務分類とページの割当（目次があれば機械が下書き。LLM が決める）
  flows/<フローID>.edits.json 直しの一覧（LLM が書く）→ flows/<フローID>.flow.json フロー（apply_edits.py が作る）
  structure/systems.json・code_terms.json、code/、activities/  プログラムがあるとき（台帳・用語の対応・構造・アクティビティ図）
  reports/・review/        読取・フロー化・照合の報告（サブエージェント）と、照合の採否表
  questions.json           確認事項の Q 番号の台帳（build_reports.py が集めて採番）
  times.json・answers.json ブックから取り込んだ、人が入れた処理時間・回答と足した列（書き出し直しても引き継ぐ）
  out/                     納品物と点検の結果（deliver.py が、門を通ったものを出力先に写す）
"""
import collections
import contextlib
import difflib
import json
import os
import pathlib
import re
import sys
import unicodedata

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent
ASSETS = SKILL_DIR / "assets"


# 認証情報を伏せる（プログラムのページ・Access の直接読みの書き出しで、同じ決まりを使う）

MASK = "＊＊＊"
SECRETS = [
    (re.compile(r"""(?i)\b([\w.]*(?:password|passwd|pwd|secret|token|api_?key|access_?key)\w*)(\s*(?::=|=|:)\s*)(["'])(?!\s*["'])(.+?)\3"""),
     lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}{MASK}{m.group(3)}"),                        # 名前に password などを含む項目に、文字を入れている
    (re.compile(r"""(?i)\b(pwd|password|passwd)(\s*=\s*)([^;"'\s&]+)"""), lambda m: f"{m.group(1)}{m.group(2)}{MASK}"),   # 接続文字列の PWD=…
    (re.compile(r"(?i)\b(Bearer\s+)[A-Za-z0-9._\-]{12,}"), lambda m: f"{m.group(1)}{MASK}"),
    (re.compile(r"(://[^/\s:@\"']+:)([^@\s/\"']+)(@)"), lambda m: f"{m.group(1)}{MASK}{m.group(3)}"),            # URL の中の user:pass@
]


def mask_secrets(lines):
    """認証情報を伏せた行の並びと、伏せた行の番号（1 から）"""
    out, hit = [], []
    for k, ln in enumerate(lines, 1):
        new = ln
        for rx, fn in SECRETS:
            new = rx.sub(fn, new)
        if new != ln:
            hit.append(k)
        out.append(new)
    return out, hit


def schema_enum(kind, *path):
    """スキーマ（assets/schemas）の決まった語。コードの定数はここから読む（決まった語を書く所を、スキーマの 1 か所にする）"""
    o = json.loads((ASSETS / "schemas" / f"{kind}.schema.json").read_text(encoding="utf-8"))
    for k in path:
        o = o[k]
    return list(o["enum"])


PAGE_TYPES = schema_enum("card", "properties", "page_type")
# フローのノードに使われなくてよいページ種別（網羅チェックの対象外）
NON_CONTENT_TYPES = {"cover", "toc", "blank"}
NODE_TYPES = schema_enum("flow", "$defs", "node", "properties", "type")
# 出典が要るノード種別（開始・終了は要らない）
SOURCED_TYPES = {"process", "decision", "system", "file", "subprocess"}
STEP_TYPES = ("process", "decision", "system", "subprocess")        # ステップとして数えるノード
MATERIAL_CAP = 40000                                                # 1つの分類の材料（カードの文字数）の上限。超えたら下位の分類に分ける
# 人に見せる呼び名（図と報告書で同じ語を使う）
ROLE_LABEL = {"regulation": "規程", "manual": "全体マニュアル", "procedure": "手順書", "appendix": "別紙・様式", "tool": "ツール", "other": "その他"}
REL_LABEL = {"parent_of": "上位→下位", "appendix_of": "別紙", "basis_of": "根拠", "refers_to": "参照", "version_of": "版違い",
             "duplicate_of": "同一"}
SEV_LABEL = {"high": "高", "mid": "中", "low": "低"}
EFFECTS = schema_enum("common", "$defs", "question", "properties", "effect")   # 確認事項の「答えで変わる所」。どれにも当たらない問いは、確認事項にしない
# 動作の種類（処理時間の仮の値を、表から引く鍵）。人の作業（process）は、LLM がこの中から選ぶ。数字は、表（業務マトリクス一覧の「処理時間の基準」）からだけ
ACTIONS = schema_enum("common", "$defs", "action")
ACT_JOIN = "／"                                                      # 2 つの動きのある箱の動作の種類（確認する／承認する）。仮の値は、基準の表で大きいほう
IO_KINDS = schema_enum("common", "$defs", "io_kind")                 # プログラムが触る相手の種類（機械の取り出しも、ブロックの木も）
IO_KIND_WAS = {"レポート": "帳票", "印刷": "帳票", "フォーム": "画面"}       # 機械の取り出しの語 → 業務の語


def io_kind(kind):
    """入出力の種類を、決まった語（IO_KINDS）にそろえる"""
    return IO_KIND_WAS.get(kind, kind)
TYPE_ACTION = {"decision": "判断する", "system": "システムが処理する", "subprocess": "別のフローへ進む"}   # 人の作業でないノードは、種別から機械が決める（既定は、判断 2 分・ほかは 0 分）
VOLUMES = schema_enum("flow", "$defs", "node", "properties", "volume")   # 量。書かなければ、機械が数えて決める（time_basis。小・中・大だけ）
VOLUME_SPARE = {"特大": "大"}                                          # 無くてもよい量（機械は付けない。LLM が、理由つきで）と、基準の表にその列が無いとき（前の形のテンプレート）に引く列
VERB_END = "うくぐすつぬぶむる"                                       # 動詞の終止形の、終わりの字（う段のひらがな）
NODE_LABEL = {"process": "プロセス", "decision": "判断", "system": "システム処理", "file": "ファイル", "subprocess": "別フローへ",
              "start": "開始", "end": "終了"}
HIERARCHY_LABEL = {"yes": "あり", "partial": "一部あり", "no": "なし"}
LABEL_LONG = 30                                                      # 作業の箱のラベルの長さの目安（決まり 3.1）。超えると点検が WARN、下書きはかっこ書きを補足に回す
QUOTE_OK = 0.75                                                       # 引用が原文と合っているとみなす、文字の一致の度合い（引用照合・見出しの照合）
SHIFT_MIN, SHIFT_SHARE = 3, 0.3                                      # 同じ向きのずれが、この件数以上・この割合以上なら「系統的なずれ」
# お客様に渡す納品物（out/ の中。PNG は out/png/）と、作業の記録。納品の門（deliver.py）・文字の走査・進み具合が使う
FLOW_FIG, ACT_FIG, MATRIX_BOOK = "業務フロー_as-is", "アクティビティ図.drawio", "業務マトリクス一覧.xlsx"   # 業務フローの図は、分けると 業務フロー_as-is_<分類>.drawio
REPORTS = ("確認事項一覧.md", "確認事項一覧.xlsx", "出典対応表.md", "出典対応表.csv", "階層と業務分類.md")
DELIVERABLES = (f"{FLOW_FIG}*.drawio", ACT_FIG, MATRIX_BOOK) + REPORTS
NUMBERS_FILE = "納品の数"                                              # 納品の門が書く数（out/納品の数.md・.json）
BASELINE_FILE = "現状の処理時間.json"                                    # 納品の門が書く、現状の処理時間の写し（BPR の比較の基準。表の版つき）
RECORDS = ("チェック_*.txt", "照合シート_*.pdf", "差分照合シート_*.pdf", "原典照合レビュー_採否表.md", f"{NUMBERS_FILE}.md", f"{NUMBERS_FILE}.json", BASELINE_FILE)
REVIEW_MIN = 0.95                                                    # 原典照合の支持率の下限（業務フロー・アクティビティ図それぞれ）。下回れば、図を直して照合し直す
MAX_STEPS, MAX_LANES = 25, 8                                         # 1 ページの図の目安（作業の箱の数・レーンの数）。業務フローは、ステップ・レーンの数だけでは分けない（決まり 2.10）


class BadJson(ValueError):
    """JSON として読めないファイル。LLM が手で書く JSON は、途中で切れる・閉じ忘れる。どのファイルかを必ず言う"""


def load(path, default=None):
    p = pathlib.Path(path)
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except ValueError as e:
        raise BadJson(f"JSON として読めない: {p}（{e}）。書きかけ・閉じ忘れを直す") from None


def try_load(path):
    """(中身, None) か (None, 読めない理由)。点検のスクリプトが、読めないファイルを FAIL にして先へ進むのに使う"""
    try:
        return load(path), None
    except BadJson as e:
        return None, str(e).split("（", 1)[-1].rsplit("）", 1)[0]


def load_card(work, doc_id, page):
    """ページカード。無い・JSON として読めない・辞書でないときは {}（1枚壊れていても、ほかのページの仕事は続ける。
    壊れたカードは check_extract.py が FAIL にする）"""
    try:
        c = load(card_json(work, doc_id, page), {})
    except BadJson:
        return {}
    return c if isinstance(c, dict) else {}


def load_page(path):
    """ページの機械抽出（pages/<資料ID>/pNNNN.json）。無い・読めない・辞書でないときは {}"""
    d = try_load(path)[0]
    return d if isinstance(d, dict) else {}


def _one_line_errors(kind, err, tb):
    """読めない JSON は、Traceback でなく、ファイル名つきの1行で止める"""
    if issubclass(kind, BadJson):
        sys.stderr.write(f"{err}\n")
    else:
        sys.__excepthook__(kind, err, tb)


sys.excepthook = _one_line_errors


def save(path, data):
    p = pathlib.Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


def work_dir(arg):
    w = pathlib.Path(arg).expanduser().resolve()
    w.mkdir(parents=True, exist_ok=True)
    return w


def split_figs(out):
    """業務フローの図を、分類ごとに分けて書き出しているか（1 つの図が無く、分けた図がある）。書き出し直すときに --split を付けるか"""
    out = pathlib.Path(out)
    return not (out / f"{FLOW_FIG}.drawio").exists() and any(out.glob(f"{FLOW_FIG}_*.drawio"))


_CTX = {"work": None, "labels": None}


def manifest(work):
    m = load(work / "manifest.json")
    if m is None:
        sys.exit(f"manifest.json が無い: {work}（先に inventory.py を実行する）")
    _CTX["work"], _CTX["labels"] = pathlib.Path(work), None       # 出典の表記（src_label）が、この作業フォルダのページ番号を引く
    return m


def page_stem(page):
    return f"p{int(page):04d}"


def page_json(work, doc_id, page):
    return work / "pages" / doc_id / f"{page_stem(page)}.json"


def page_png(work, doc_id, page):
    return work / "pages" / doc_id / f"{page_stem(page)}.png"


def card_json(work, doc_id, page):
    return work / "cards" / doc_id / f"{page_stem(page)}.json"


def docs_by_id(m):
    return {d["doc_id"]: d for d in m["docs"]}


CODE_PAGE_KINDS = ("vba", "code")                                    # プログラムのソースのページ（カードにしない。アクティビティ図の側で読む）


def all_pages(m, code=False):
    """(doc, page_entry) を、読む対象の資料（重複でないもの）について順に返す。
    プログラムのソースのページは、code=True のときだけ含める（カード・業務分類・ページの網羅の対象にしない）。
    道具のブック（マクロ付きの Excel で、利用者が「シートは手順書ではない」とした資料。tool）のシートも、同じ扱い（出典には引ける）"""
    for d in m["docs"]:
        if d.get("duplicate_of") or not d.get("page_list"):
            continue
        for p in d["page_list"]:
            if code or not (is_code_page(p) or d.get("tool")):
                yield d, p


def is_code_page(p):
    """プログラムのソースのページか（いまの版でページ化したもの。前の版のマクロのページは、カードにして読んでいるので含めない）"""
    return p.get("kind") in CODE_PAGE_KINDS and bool(p.get("code"))


def parse_pages(spec):
    """「13-24」「3,5,8-10」をページ番号の集合にする。空なら None"""
    if not spec:
        return None
    out = set()
    for part in re.split(r"[,、]", unicodedata.normalize("NFKC", str(spec))):
        mt = re.fullmatch(r"\s*p?0*(\d+)\s*(?:[-~〜]\s*p?0*(\d+))?\s*", part)     # 束の一覧の書き方（p0013〜p0018）も受ける
        if not mt or int(mt.group(2) or mt.group(1)) < int(mt.group(1)):
            sys.exit(f"ページの範囲「{spec}」が読めない。13-24 や 3,5,8-10 の形で書く（ファイル名 pNNNN の番号）")
        out.update(range(int(mt.group(1)), int(mt.group(2) or mt.group(1)) + 1))
    return out


PRIVATE = re.compile("[\ue000-\uf8ff\U000f0000-\U000ffffd\U00100000-\U0010fffd]")   # 外字（フォントごとの記号。Wingdings の箇条書きの印など）。文字としての意味を持たない


def norm(text):
    """照合用の正規化: 全角半角をそろえ、空白・改行・外字を落とす"""
    if not isinstance(text, str):                                    # 名前の項目に数や配列が書かれていても落ちない（形の違反は validate が言う）
        text = "" if text is None else str(text)
    return re.sub(r"\s+", "", PRIVATE.sub("", unicodedata.normalize("NFKC", text)))


def tree_nodes(nodes, parent=None):
    """木（hierarchy.json の tree など、children を持つ入れ子）を上から順にたどる。(ノード, 親) を返す"""
    for n in nodes or []:
        if isinstance(n, dict):
            yield n, parent
            yield from tree_nodes(n.get("children"), n)


def style_dict(style):
    """draw.io のスタイル文字列（"rounded=1;fillColor=#fff;"）を辞書に。値の無い語（"ellipse"）は空文字で入る"""
    d = {}
    for part in (style or "").split(";"):
        if part:
            k, _, v = part.partition("=")
            d[k] = v
    return d


def page_text(pj):
    """機械抽出の JSON から、そのページにある文字をすべて1本にする"""
    if not pj:
        return ""
    parts = list(pj.get("text") or []) + list(pj.get("cells") or [])
    for t in pj.get("tables") or []:
        for row in t:
            parts += [c for c in row if c]
    for s in pj.get("shapes") or []:
        if s.get("text"):
            parts.append(s["text"])
    for c in pj.get("connectors") or []:
        if c.get("text"):
            parts.append(c["text"])
    if pj.get("notes"):
        parts.append(pj["notes"])
    return "\n".join(parts)


GAP_CHARS, GAP_SHARE, PATCHWORK = 3, 0.1, 0.5                         # 原典の側の隙間がこれを超えたら継ぎはぎ。継ぎはぎの一致率の上限


def quote_match(quote, haystack):
    """引用が原典の文字にどれだけ出てくるか（0〜1）。1 は、原典の連続した一部そのまま（空白と全角半角の違いだけ）のときに限る。
    言い回しの小さな違いは部分一致で拾うが、1 にはしない。原典の離れた所から拾った断片をつないだ文（主語や条件をすり替えた継ぎはぎ）は、
    断片がどれだけ当たっていても PATCHWORK まで"""
    q, h = norm(quote), norm(haystack)
    if not q:
        return 0.0
    if q in h:
        return 1.0
    if not h:
        return 0.0
    sm = difflib.SequenceMatcher(None, q, h, autojunk=False)
    blocks = [b for b in sm.get_matching_blocks() if b.size >= 2]
    score = min(sum(b.size for b in blocks) / len(q), 0.99)
    gaps = sum(max(0, nxt.b - (prev.b + prev.size)) for prev, nxt in zip(blocks, blocks[1:]))
    return min(score, PATCHWORK) if gaps > max(GAP_CHARS, GAP_SHARE * len(q)) else score


# ---------- カードの項目（手順 s・分岐 d・図の箱 g） ----------

def card_items(card):
    """カードの項目を {id: {"kind", "text", "quote", "origin", "from_image"}} で返す。フローの出典（sources[].item）はこの id を引く"""
    out = {}
    for s in (card or {}).get("steps") or []:
        if s.get("id"):
            out[s["id"]] = {"kind": "step", "text": str(s.get("action", "")), "quote": str(s.get("quote", "")),
                            "origin": s.get("origin", "body"), "from_image": bool(s.get("from_image")), "actor": str(s.get("actor") or "").strip(),
                            "system": str(s.get("system") or "").strip()}
    for d in (card or {}).get("decisions") or []:
        if d.get("id"):
            out[d["id"]] = {"kind": "decision", "text": str(d.get("question", "")), "quote": str(d.get("quote", "")),
                            "origin": d.get("origin", "body"), "from_image": bool(d.get("from_image"))}
    for g in ((card or {}).get("diagram") or {}).get("nodes") or []:
        if g.get("id"):
            out[g["id"]] = {"kind": "box", "text": str(g.get("label", "")), "quote": str(g.get("label", "")),
                            "origin": "body", "from_image": False}
    return out


def same_text(a, b, ratio=0.6):
    """2つの引用が同じ箇所を指しているか（どちらかがもう一方にほぼ含まれる）"""
    return bool(norm(a)) and bool(norm(b)) and (quote_match(a, b) >= ratio or quote_match(b, a) >= ratio)


# ---------- 補足の中の作業（機械で候補を挙げる） ----------

def load_markers(path=None):
    """assets/action_markers.txt を {節の名前: [行]} で読む"""
    out, cur = {"anywhere": [], "at_start": [], "heading": [], "action": [], "directive": [], "not_directive": [], "not_action": []}, None
    for ln in pathlib.Path(path or ASSETS / "action_markers.txt").read_text(encoding="utf-8").splitlines():
        ln = ln.strip()
        if not ln or ln.startswith("#"):
            continue
        if ln.startswith("[") and ln.endswith("]"):
            cur = ln[1:-1]
            out.setdefault(cur, [])
        elif cur:
            out[cur].append(ln)
    return out


BULLET = re.compile(r"^(?:[・･\-–—●○■□◆◇▶➤➲*＊]|[①-⑳]|[(（]?[0-9０-９]{1,2}[)）.．、]|[(（][ア-ンa-zA-Z][)）])")   # 箇条書きの頭


def _cut(pj, breaks=()):
    """機械抽出の文字を文の単位に切った並び: (文, 外字の箇条書きの印で始まっていたか)"""
    out, buf, lead = [], "", False

    def flush(split):
        nonlocal buf, lead
        parts = [x for x in (re.split(r"(?<=[。．])", buf) if split else [buf]) if x.strip()]
        out.extend((x, lead and k == 0) for k, x in enumerate(parts))
        buf, lead = "", False
    for raw in page_text(pj).split("\n"):
        for ln in raw.split(" | "):                                  # Excel は1行に複数のセル
            ln = ln.strip()
            mark = bool(PRIVATE.match(ln))
            ln = PRIVATE.sub("", ln).strip()
            if buf and (mark or ln.startswith(tuple(breaks)) or BULLET.match(ln)):
                flush(False)
            lead = lead or mark                                      # 印だけの行のあとに文が続くときは、その文の印
            if not ln:
                continue
            buf += ln
            if ln.endswith(("。", "．")) or len(norm(ln)) <= 12:
                flush(True)
    if buf.strip():
        flush(True)
    return out


def sentences(pj, breaks=()):
    """機械抽出の文字を、文の単位に切る。PDF の行は文の途中で折れるので、句点か短い行（見出し・ラベル）までをつなぐ。
    breaks の語（※・なお・ただし…）や箇条書きの頭で始まる行の前では、つないでいる途中でも切る（前の行に句点が無くても、別の文）。
    外字（フォントごとの記号）は落とす。行の頭の外字は箇条書きの印なので、そこで切る"""
    return [x for x, _lead in _cut(pj, breaks)]


_MARKERS = {}


def unit_breaks():
    """文の単位に切るとき、前の行に句点が無くても切る語（※・なお・ただし…）。補足の作業の候補（note_actions）と、行の番号（page_units）で同じ切り方をする"""
    if "mk" not in _MARKERS:
        _MARKERS["mk"] = load_markers()
    mk = _MARKERS["mk"]
    return [x for x in mk["anywhere"] + mk["at_start"] if len(x) >= 1 and x != "*"]


def page_units(pj):
    """ページの機械抽出を、番号で指せる単位（文・見出しの行・表のセル）に切った並び。引用を番号で指すときの「行」（番号は 1 から）。
    python3 scripts/pages.py lines が番号つきで見せ、quotes.py がその文字を引用に入れる。
    切り方を変えると、カードに書いてある番号がずれる。変えたら quotes.py locate で引き直す"""
    if (pj or {}).get("code"):                                       # プログラムのソースは、1行が1単位（空の行も数える。番号がソースの行と合う）
        return [str(x).rstrip() for x in pj.get("text") or []]
    return [x.strip() for x in sentences(pj, breaks=unit_breaks())]


def unit_marks(pj):
    """page_units と同じ並びで、その行の頭の印: "number"（1.・(1)・①・(ア)。順のある並び）／"bullet"（・●■ や外字。順の無い箇条書き）／""（印なし）"""
    out = []
    if (pj or {}).get("code"):
        return [""] * len(pj.get("text") or [])
    for text, lead in _cut(pj, unit_breaks()):
        m = BULLET.match(text.strip())
        out.append("bullet" if lead or (m and len(m.group(0)) == 1 and not "①" <= m.group(0) <= "⑳") else "number" if m else "")
    return out


def span_text(units, span):
    """行の番号 [開始, 終了] の文字（つないだもの）。番号が範囲の外・形が違えば None"""
    if not (isinstance(span, list) and len(span) == 2 and all(isinstance(x, int) and not isinstance(x, bool) for x in span)
            and 1 <= span[0] <= span[1] <= len(units)):
        return None
    return "".join(units[span[0] - 1:span[1]])


def action_like(sent, markers=None):
    """作業の言い回しで終わる文か（「〜します。」「〜を行う。」。「〜があります。」「〜とする。」のような、作業でない言い回しは除く）"""
    mk = markers or load_markers()
    t = re.sub(r"\.$", "。", norm(sent))
    if not t.endswith("。") or len(t) < 8:
        return False
    tail = re.sub(r"[(（][^()（）]*[)）](?=。?$)", "", t)
    last = re.split(r"[、,]", tail)[-1]
    return any(re.search(x, tail) for x in mk["action"]) and not any(re.search(x, last) for x in mk["not_action"])


def note_actions(pj, markers=None):
    """補足の目印があり、動作の言い回しで書かれた文（＝補足の中の作業の候補）を返す"""
    mk = markers or load_markers()
    act = [re.compile(x) for x in mk["action"]]
    non = [re.compile(x) for x in mk["not_action"]]
    ask = [re.compile(x) for x in mk.get("directive") or []]
    not_ask = [re.compile(x) for x in mk.get("not_directive") or []]
    heads = {norm(x).lower() for x in mk["heading"]}
    out, block = [], 0                                               # block: 見出し（注意事項 など）のあと、何文目か。箇条書きが続くあいだは補足として見る
    for sent in sentences(pj, breaks=[x for x in mk["anywhere"] + mk["at_start"] if len(x) >= 1 and x != "*"]):      # 切り方は page_units と同じ（unit_breaks）
        t = re.sub(r"\.$", "。", norm(sent))                         # 文末の「．」は NFKC で「.」になる
        bare = re.sub(r"[(（][^()（）]*[)）]$", "", re.sub(r"^[\W_]+|[\W_]+$", "", t)).lower()
        if bare in heads or (len(bare) <= 24 and "。" not in t and any(len(h) >= 3 and bare.endswith(h) for h in heads)):
            block = 1                                                # 「注意事項」だけの行と、それで終わる短い見出し（「支払の手続きの注意事項」）
            continue
        body = re.sub(r"^[\W_]+", "", t) if not t.startswith(("※", "＊", "*", "(注", "（注")) else t
        marked = any(norm(x) in t for x in mk["anywhere"]) or bool(re.search(r"注[0-9]", t)) \
            or any(body.startswith(norm(x)) or t.startswith(norm(x)) for x in mk["at_start"])
        noted = marked or block == 1 or (block > 1 and bool(BULLET.match(sent.strip())))
        block = block + 1 if noted and block else 0
        tail = re.sub(r"[(（][^()（）]*[)）](?=。?$)", "", t)          # 文末のかっこ書き（「…連絡する（内線 1234）。」）は外して文末を見る
        last = re.split(r"[、,]", tail)[-1]                          # 「作業でない」言い回しは、最後の節だけに当てる（「〜場合があるので、…すること。」は作業）
        asked = any(r.search(tail) for r in ask) and not any(r.search(last) for r in not_ask)   # 読み手に頼む文（「〜してください」）は、目印が無くても候補
        acts = any(r.search(tail) for r in act) and not any(r.search(last) for r in non)
        if (asked or (noted and acts)) and len(t) >= (5 if marked else 8) \
                and not any(covers(sent, x, 0.9) for x in out):       # 同じ文が文字層と図形の両方から出てくることがある
            out.append(sent.strip())
    return out


def covers(quote, sentence, ratio=0.8):
    """引用が、その文を指しているか（引用が文にほぼ含まれるか、文が引用にほぼ含まれる）"""
    return bool(norm(quote)) and (quote_match(quote, sentence) >= ratio or quote_match(sentence, quote) >= ratio)


@contextlib.contextmanager
def guarded(rp, what):
    """点検の1単位（カード1枚・フロー1本・ファイル1つ）を囲む。形が崩れた JSON で例外が出ても、そこまでの結果を残して次の単位へ進む。
    点検は最後にまとめて出すので、途中で落ちると、溜めた FAIL が1行も出なくなる"""
    try:
        yield
    except Exception as e:                                           # 何が来ても Traceback だけで終わらせない
        rp.fail(f"{what}: 形が崩れていて、ここから先を点検できない（{type(e).__name__}: {str(e)[:80]}）。「形が違う」の行を先に直す"
                f"（python3 scripts/validate.py -w <作業フォルダ> で形だけを確かめられる）")


FAIL_LOG = "ASIS_FAIL_LOG"                                            # これがあれば、Report.fail・warn が、出した場所（ファイル:行）を、そのファイルに書き足す
FAIL_HELPERS = ("shape_fails", "read_checked", "evidence_pages", "question_sources", "check_question_ids", "check_label_form", "stale_refs")   # 点検の共通の部品（呼び元の行も控える）


def site(msg, up=1):
    """中継する点検の内側の誤りの文（直しの一覧・アクティビティ図の直し・台帳の紐付け・ブロックの木・取り出した構造・用語の対応・照合の報告・
    業務マトリクス一覧と確認事項一覧の読み戻し・処理時間の基準の表・人の値の控え。点検は、その誤りを 1 行で中継する）。
    自己点検の被覆（evals/coverage.py）のために、出した場所（ファイル:行）を控え、文はそのまま返す。up: 何段上の行を控えるか（誤りを溜める部品 bad の中なら 2。呼び元の行）"""
    if FAIL_LOG in os.environ:
        Report._log(sys._getframe(up))
    return msg


class Report:
    """FAIL / WARN / INFO を溜めて出す。FAIL があれば終了コード 1（consulting-pptx-skill の check_deck.py と同じ流儀）。
    画面には、同じ種類の行を fold 件まで出して残りは件数にする（大量の入力で、読む側の文脈をあふれさせないため）。
    全行は save_to のファイルに書く。"""

    def __init__(self, title, fold=10):
        self.title, self.fails, self.warns, self.infos, self.fold = title, [], [], [], fold

    @staticmethod
    def _log(frame, at=None):
        """自己点検の被覆（evals/coverage.py）: FAIL・WARN を出した点検の場所を控える。共通の部品の中なら、呼び元の行も"""
        with open(os.environ[FAIL_LOG], "a", encoding="utf-8") as fh:
            fh.write(f"{os.path.basename(frame.f_code.co_filename)}:{at or frame.f_lineno}\n")
            while frame.f_code.co_name in FAIL_HELPERS and frame.f_back is not None:
                frame = frame.f_back
                fh.write(f"{os.path.basename(frame.f_code.co_filename)}:{frame.f_lineno}\n")

    def fail(self, msg, first=False, at=None):
        """first: 先頭に出す（系統的なずれ）。at: 出した場所の行（配置の点検のように、yield で返した点検の行）"""
        if FAIL_LOG in os.environ:
            self._log(sys._getframe(1), at)
        self.fails.insert(0, msg) if first else self.fails.append(msg)

    def fail_if(self, cond, msg):
        """cond なら FAIL、でなければ WARN（段によって、同じ誤りの重さが変わる点検）。被覆の控えは FAIL のときだけ（WARN の側だけ出ても、FAIL を確かめたことにしない）"""
        if cond and FAIL_LOG in os.environ:
            self._log(sys._getframe(1))
        (self.fails if cond else self.warns).append(msg)

    def warn(self, msg):
        if FAIL_LOG in os.environ:
            self._log(sys._getframe(1))
        self.warns.append(msg)

    def info(self, msg):
        self.infos.append(msg)

    @staticmethod
    def _kind(msg):
        """行の「種類」: 場所（資料・ページ・ノード）と数字と引用を落とした残り"""
        body = msg.split(": ", 1)[1] if ": " in msg[:60] else msg
        body = re.sub(r"「[^」]*」|〔[^〕]*〕|（[^）]*）|\[[^\]]*\]", "", body)
        return re.sub(r"[0-9A-Za-z_.\-]+", "", body)[:16]

    def lines(self, fold=None):
        out = [f"== {self.title} =="]
        out += [f"INFO  {m}" for m in self.infos]
        for tag, msgs in (("WARN", self.warns), ("FAIL", self.fails)):
            seen = {}
            for m in msgs:
                k = self._kind(m)
                seen[k] = seen.get(k, 0) + 1
                if not fold or seen[k] <= fold:
                    out.append(f"{tag}  {m}")
            for k, n in seen.items():
                if fold and n > fold:
                    out.append(f"{tag}  … 同じ種類があと {n - fold} 件（「{k.strip()[:14]}」。全件は --report のファイル）")
        out.append(f"{len(self.fails)} FAIL / {len(self.warns)} WARN")
        return out

    def finish(self, save_to=None):
        print("\n".join(self.lines(self.fold)))
        if save_to:
            p = pathlib.Path(save_to)
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text("\n".join(self.lines()) + "\n", encoding="utf-8")
        sys.exit(1 if self.fails else 0)


# ---------- ページ番号と出典の表記 ----------
# データの鍵は「先頭から数えた番号」（page・ファイル名 pNNNN）。重ならず、必ずあり、ページ画像と一対一だから。
# 人に見せる番号は「ページに刷られた番号」。資料を開いた人が見るのはこちらで、表紙を数えない資料では 1 ずれる。
# 納品物（図・出典対応表・確認事項一覧・照合シート）には刷られた番号だけを出し、書き換えはここ1か所で行う。

TYPE_LABEL = {"cover": "表紙", "toc": "目次", "blank": "白紙"}
NUM_RE = re.compile(r"^(.*?)(\d{1,4})$")


def bare_number(label):
    """刷られた番号から飾りを外す（「- 13 -」「（13）」「P.13」「13 ページ」「13/20」→「13」）。「3-2」「ii」はそのまま"""
    t = norm("" if label is None else str(label))
    t = re.sub(r"^[-‐－—–―~(\[【<]+|[-‐－—–―~)\]】>]+$", "", t)
    t = re.sub(r"^(?:p\.?|page)(?=\d)", "", t, flags=re.I)
    t = re.sub(r"(?<=\d)(?:ページ|頁)$", "", t)
    mt = re.fullmatch(r"(\d+)/\d+", t)
    return mt.group(1) if mt else t


def _split(label):
    """「13」→("", 13)／「3-2」→("3-", 2)／「- 13 -」→("", 13)／数字で終わらなければ None"""
    mt = NUM_RE.match(bare_number(label))
    return (mt.group(1), int(mt.group(2))) if mt else None


def _longest_run(order, got):
    """got のページが、資料の並び order の中で何枚続いているか（いちばん長い所）"""
    best = cur = 0
    for p in order:
        cur = cur + 1 if p in got else 0
        best = max(best, cur)
    return best


def printed_numbers(pages):
    """1資料ぶんの、刷られたページ番号を決める。
    pages: [{"page": 先頭から数えた番号, "cands": [{"n": "13", "x": 0.5, "y": 0.96}], "card": カードの printed_page}]（順に）
    返り値: {page: (番号の文字, どう決めたか)}。どう決めたか＝machine（ページの端の数字）/ card（カードの読み）/ inferred（前後の並びから補った）
    ページの端には番号でない数字もある（見出しの「手順 2」、年）。1ずつ増える並びが 3 ページ以上続くものだけを番号とみなす
    （離れたページにたまたま同じ差の数字があるだけでは、並びにしない）。同じ長さなら、ページの下端の並びを採る。"""
    tracks = {}
    for pg in pages:
        for c in pg.get("cands") or []:
            sp = _split(c.get("n"))
            if sp:                                                   # 位置は上端・下端だけを見る。桁が増えると番号の中心は少し動く
                key = ("m", float(c.get("y", 0.95)) >= 0.5, sp[0], pg["page"] - sp[1])
                tracks.setdefault(key, {})[pg["page"]] = sp
        sp = _split(pg["card"]) if pg.get("card") else None
        if sp:
            tracks.setdefault(("c", True, sp[0], pg["page"] - sp[1]), {})[pg["page"]] = sp
    need = 3 if len(pages) >= 6 else 2
    order = [pg["page"] for pg in pages]
    out = {}
    for src in ("m", "c"):                                           # 機械で拾えたものを先に。無いページだけカードの読みを使う
        good = sorted((t for t in tracks.items() if t[0][0] == src and _longest_run(order, t[1]) >= need),
                      key=lambda t: (-len(t[1]), not t[0][1]))
        for key, got in good:
            for p, (pre, n) in got.items():
                out.setdefault(p, (f"{pre}{n}", "machine" if src == "m" else "card", (key[-2], key[-1])))
    for i, p in enumerate(order):                                    # 間の、番号が刷られていないページを前後の並びから補う
        if p in out:
            continue
        prev = next((out[q] for q in reversed(order[:i]) if q in out), None)
        nxt = next((out[q] for q in order[i + 1:] if q in out), None)
        if prev and nxt and prev[2] == nxt[2] and p - prev[2][1] >= 1:
            out[p] = (f"{prev[2][0]}{p - prev[2][1]}", "inferred", prev[2])
    return {p: (v[0], v[1]) for p, v in out.items()}


def sheet_name(p):
    """ページ（シート）の名前。「手順（2/5）」「台帳（非表示）」の添え書きを外す"""
    return p.get("sheet") or re.sub(r"（(\d+/\d+|非表示|画像なし)）$", "", str(p.get("label", "")))


def cover_title(cards):
    """表紙のカード（page_type が cover。裏表紙は除く）の見出し。無ければ空。資料の題名の、機械の見立て"""
    for c in cards:
        t = str((c or {}).get("title") or "").strip()
        if c.get("page_type") == "cover" and t and "裏表紙" not in t:
            return " ".join(t.split())
    return ""


def file_stem(name):
    """ファイル名から、フォルダと拡張子を外したもの（「手順書/月次の手順.pdf」→「月次の手順」）"""
    return re.sub(r"\.[A-Za-z0-9]{1,6}$", "", pathlib.PurePath(str(name or "")).name)


class Labels:
    """(資料ID, page) → 人に見せる書き方。作業フォルダのページの機械抽出とカードから、その場で引く。
    資料は、題名で書く（資料ID は、機械が読み込んだ順に振った番号なので、納品物には出さない。点検の記録だけが ID で書く — ref）"""

    def __init__(self, work, m):
        self.work, self.m, self.map, self.docs, self.cover = pathlib.Path(work), m, {}, {}, {}
        for d in m["docs"]:
            pl = d.get("page_list") or []
            if d.get("duplicate_of") or not pl:
                continue
            rows = []
            for p in pl:
                pj = load_page(self.work / p["json"]) if p.get("json") else {}
                c = load_card(self.work, d["doc_id"], p["page"])
                rows.append({"page": p["page"], "kind": p.get("kind"), "entry": p, "card": c,
                             "cands": (pj or {}).get("printed_candidates") or []})
            cover = cover_title(r["card"] for r in rows)
            if cover:
                self.cover[d["doc_id"]] = cover
            paged = [r for r in rows if r["kind"] in ("page", "slide")]
            nums = printed_numbers([{"page": r["page"], "cands": r["cands"], "card": r["card"].get("printed_page")} for r in paged])
            seen = {}
            for r in rows:
                k, p, c = (d["doc_id"], r["page"]), r["entry"], r["card"]
                unit = "スライド " if r["kind"] == "slide" else "p."
                off = "（非表示）" if "（非表示）" in str(p.get("label", "")) else ""      # 非表示のスライド・シート。読みはするが、現行の手順とは限らない
                if r["kind"] == "sheet":
                    e = {"text": f"「{sheet_name(p)}」シート{off}", "how": "sheet", "printed": None}
                elif r["kind"] in CODE_PAGE_KINDS:
                    mod = re.sub(r"（\d+/\d+）$", "", str(p.get("label", "")).replace("VBA:", ""))
                    e = {"text": f"マクロ「{mod}」" if r["kind"] == "vba" else f"「{mod}」", "how": "vba", "printed": None,
                         "line_from": (p.get("code") or {}).get("line_from")}
                elif r["kind"] in ("image", "text"):
                    e = {"text": "（画像）" if r["kind"] == "image" else str(p.get("label", "本文")), "how": r["kind"], "printed": None}
                elif c.get("page_type") == "cover":                  # 表紙・裏表紙は、番号ではなく種別で書く（文字としては番号が入っていても、表紙の絵に隠れて見えないことがある）
                    e = {"text": "裏表紙" if "裏表紙" in str(c.get("title") or "") else TYPE_LABEL.get("cover", "表紙"), "how": "unnumbered", "printed": None}
                elif r["page"] in nums:
                    e = {"text": f"{unit}{nums[r['page']][0]}", "how": nums[r["page"]][1], "printed": nums[r["page"]][0]}
                elif nums:                                           # 番号のある資料の、番号が刷られていないページ（表紙・目次・奥付）
                    name = TYPE_LABEL.get(c.get("page_type")) or (f"「{str(c.get('title'))[:14]}」" if c.get("title") else "番号なしのページ")
                    e = {"text": name, "how": "unnumbered", "printed": None}
                else:                                                # 資料全体に番号が無い。先頭から数えた番号（スライド番号・PDF を開いたときの番号）
                    e = {"text": f"{unit}{r['page']}{off}", "how": "count", "printed": None}
                seen.setdefault(e["text"], []).append(k)
                self.map[k] = e
            for text, ks in seen.items():                            # 同じ書き方が重なったら（章ごとに番号を振り直す資料・別紙）、章の見出しを添える
                if len(ks) > 1 and self.map[ks[0]]["how"] not in ("sheet", "vba", "text"):
                    paths = {k: [str(x) for x in load_card(self.work, *k).get("section_path") or []] for k in ks}
                    depth = next((i for i in range(max(map(len, paths.values()), default=0))    # 上から見て、初めて食い違う段の見出しを添える
                                  if len({tuple(v[i:i + 1]) for v in paths.values()}) > 1), None)
                    named = {}
                    for k in ks:
                        sec = paths[k][depth] if depth is not None and len(paths[k]) > depth else ""
                        self.map[k]["text"] = f"{sec[:14]} {text}" if sec and norm(sec) != norm(text) else text
                        named.setdefault(self.map[k]["text"], []).append(k)
                    for same in named.values():                      # 見出しを添えても重なるものには、何番目かを付ける
                        if len(same) > 1:
                            for i, k in enumerate(same, 1):
                                self.map[k]["text"] += f"（{i} つ目）"
            hows = [self.map[(d["doc_id"], r["page"])]["how"] for r in paged]
            self.docs[d["doc_id"]] = {
                "numbered": sum(h in ("machine", "card", "inferred") for h in hows), "paged": len(paged),
                "basis": ("sheet" if not paged else "printed" if nums else "count"),
                "offsets": sorted({r["page"] - _split(nums[r["page"]][0])[1] for r in paged if r["page"] in nums})}
        self.titles = self._titles()

    def _titles(self):
        """資料ID → 資料の題名。hierarchy.json の docs[].title（LLM が表紙から写す）→ 表紙のカードの見出し → ファイル名（拡張子なし）。
        題名が重なる資料は、ファイル名を添えて見分ける（ファイル名も重なれば、フォルダからの場所）。同一の写しは、もとの資料と重ならないように"""
        try:
            hier = load(self.work / "structure" / "hierarchy.json", {}) or {}
        except BadJson:
            hier = {}
        given = {x.get("doc_id"): " ".join(str(x.get("title") or "").split()) for x in hier.get("docs") or [] if isinstance(x, dict)}
        base = {d["doc_id"]: given.get(d["doc_id"]) or self.cover.get(d["doc_id"]) or file_stem(d.get("name")) or d["doc_id"] for d in self.m["docs"]}
        out, taken = {}, set()
        for group in ([d for d in self.m["docs"] if not d.get("duplicate_of")], [d for d in self.m["docs"] if d.get("duplicate_of")]):
            by = {}
            for d in group:
                by.setdefault(norm(base[d["doc_id"]]), []).append(d)
            for key, same in by.items():
                names = collections.Counter(norm(d.get("name")) for d in same)
                for d in same:
                    t = base[d["doc_id"]]
                    if len(same) > 1 or key in taken:
                        t += f"（{d.get('name') if names[norm(d.get('name'))] == 1 and key not in taken else d.get('rel_path') or d.get('name')}）"
                    out[d["doc_id"]] = t
            taken |= set(by)
        return out

    def title(self, doc_id):
        """資料の題名（人に見せる、資料の名前）"""
        return self.titles.get(doc_id) or str(doc_id)

    def page(self, doc_id, page):
        return self.map.get((doc_id, page), {"text": f"p.{page}", "how": "unknown", "printed": None})

    def short(self, doc_id, page):
        """「p.13」「スライド 8」「「手順」シート」"""
        return self.page(doc_id, page)["text"]

    @staticmethod
    def glue(name, text):
        """「業務マニュアル p.13」「請求管理「手順」シート」（かぎかっこで始まる書き方は、間を空けない）"""
        return f"{name}{'' if str(text).startswith(('「', '（')) else ' '}{text}"

    def join(self, doc_id, text):
        """資料の題名と、ページの書き方: 「業務マニュアル p.13」。プログラムのソースで、モジュールの名前が題名と同じなら、1回だけ書く（「「月次集計」」）"""
        name, mod = self.title(doc_id), re.fullmatch(r"(?:マクロ)?「(.+)」", str(text))
        return str(text) if mod and norm(mod.group(1)) == norm(name) else self.glue(name, text)

    def cell(self, doc_id, page, quote):
        """Excel のシートで、引用が載っているセルの番地（引けなければ空）"""
        p = next((x for d in self.m["docs"] if d["doc_id"] == doc_id for x in d.get("page_list") or [] if x["page"] == page), None)
        if not p or p.get("kind") != "sheet" or not norm(quote):
            return ""
        pj = load_page(self.work / p["json"])
        lines, at = pj.get("cells") or pj.get("text") or [], pj.get("cell_at") or []
        q = norm(quote)
        for sh in pj.get("shapes") or []:                            # 図形の文字そのままなら、図形の置かれたセル
            if sh.get("cell") and norm(sh.get("text", "")) == q:
                return sh["cell"]

        def fit(text):
            """引用がこのセル（行）を指している度合い。引用の一部でしかない短いセル（担当名・帳票名だけ）は、その分だけ低くする"""
            t = norm(text)
            if not t:
                return 0.0
            return quote_match(quote, text) if len(t) >= len(q) else quote_match(text, quote) * len(t) / len(q)

        cells = [(fit(x), -abs(len(norm(x)) - len(q)), -i, -j) for i, ln in enumerate(lines[:len(at)]) if not isinstance(at[i], str)
                 for j, x in enumerate(str(ln).split(" | ")[:len(at[i])])]
        best = max(cells, default=(0, 0, 0, 0))                      # まずセル1つに当てる。同点なら、長さが引用に近いセル・上の行・左のセル
        if best[0] >= QUOTE_OK:
            return at[-best[2]][-best[3]]
        row = max(((fit(ln), -i) for i, ln in enumerate(lines[:len(at)])), default=(0, 0))   # 引用が複数のセルにまたがるときは、行の先頭のセル
        if row[0] >= QUOTE_OK:
            return at[-row[1]] if isinstance(at[-row[1]], str) else at[-row[1]][0]
        for sh in pj.get("shapes") or []:
            if sh.get("cell") and same_text(quote, sh.get("text", ""), QUOTE_OK):
                return sh["cell"]
        return ""

    def code_lines(self, doc_id, page, span):
        """プログラムのソースのページで、ページの中の行の番号 [開始, 終了] → モジュールの行の書き方（「31〜37 行」）。ソースのページでなければ空"""
        at = self.page(doc_id, page).get("line_from")
        if not at or not (isinstance(span, list) and len(span) == 2 and all(isinstance(x, int) for x in span)):
            return ""
        a, b = span[0] + at - 1, span[1] + at - 1
        return f"{a} 行" if a == b else f"{a}〜{b} 行"

    def src(self, src):
        """出典の書き方: 「業務マニュアル p.13」「研修資料 スライド 8」「請求管理「手順」シート B12」「請求管理 マクロ「月次集計」 31〜37 行」"""
        d, p = src.get("doc_id"), src.get("page")
        if p is None:
            return self.title(d)
        cell = self.code_lines(d, p, src.get("lines")) or (self.cell(d, p, src.get("quote", "")) if src.get("quote") else "")
        return self.join(d, self.short(d, p)) + (f" {cell}" if cell else "")

    def ref(self, doc_id, page):
        """点検の記録用。資料ID で書き、ファイルの場所を添える: 「M01 p.13〔M01/p0014〕」（作業フォルダのページのファイルを、ID で引くため）"""
        return f"{self.glue(doc_id, self.short(doc_id, page))}〔{doc_id}/{page_stem(page)}〕"

    def find(self, doc_id, printed):
        """刷られた番号 → page（目次の番号から、実際のページを引く）。見つからなければ None"""
        want = norm(str(printed))
        hit = [k[1] for k, e in self.map.items() if k[0] == doc_id and e["printed"] is not None and norm(e["printed"]) == want]
        return hit[0] if len(hit) == 1 else None


def labels(m=None, work=None):
    """いまの作業フォルダの Labels（1回だけ作る）"""
    if work is not None or _CTX["labels"] is None:
        w = pathlib.Path(work) if work is not None else _CTX["work"]
        if w is None:
            raise RuntimeError("作業フォルダが決まっていない（先に manifest(work) を呼ぶ）")
        _CTX["work"] = w
        _CTX["labels"] = Labels(w, m if m is not None else load(w / "manifest.json"))
    return _CTX["labels"]


def src_label(m, src):
    """出典の表示用の書き方。刷られたページ番号で出す（Excel はシート名とセル、マクロはモジュール名）"""
    return labels(m).src(src)


# ---------- フローと確認事項 ----------

class Answers:
    """人が表の納品物（確認事項一覧・業務マトリクス一覧）の、人が埋める欄（テンプレートの、印の無い列）に入れた値の置き場（作業フォルダの answers.json）。
    書き出す前に、いまのブックから行の鍵（確認事項は Q 番号・業務マトリクスはフロー・業務フローのノード一覧はフローとノード番号）で取り込み、書き出しで戻す。人の値が正"""

    def __init__(self, work):
        self.path = pathlib.Path(work) / "answers.json"
        d = load(self.path, {})
        if self.path.exists():                                       # 崩れたまま上書きすると、人が入れた値が消える
            from validate import validate
            errs = validate(d, "answers")
            if errs:
                raise BadJson(site(f"answers.json の形が崩れている（{'／'.join(errs[:3])}）。人が入れた値を消さないよう止める。answers.json を直すか、値の入ったブックから取り込み直す"))
        self.books = (d if isinstance(d, dict) else {}).get("books") or {}
        self.before = json.dumps(self.books, ensure_ascii=False, sort_keys=True)

    def take(self, name, got, seen):
        """ブック name から読んだ値（table_template.manual_back）を取り込む。ブックにある行は、ブックの値に置き換える（ブックで消した値は消す）。
        ブックに無い行の値は残す（確認事項が無くなった・フローが無くなった。点検が知らせる）。seen: 一覧 → 鍵 → いまの行の中身。返すのは、変わった行の数"""
        import datetime
        n, book = 0, self.books.setdefault(name, {})
        for canon, rows in got.items():
            mine = book.setdefault(canon, {})
            for key, cells in rows.items():
                old = mine.get(key)
                if cells and (old is None or old["cells"] != cells):
                    mine[key] = {"cells": cells, "seen": (seen.get(canon) or {}).get(key, (old or {}).get("seen", "")), "date": datetime.date.today().isoformat()}
                    n += 1
                elif not cells and old is not None:
                    del mine[key]
                    n += 1
            if not mine:
                del book[canon]
        if not book:
            del self.books[name]
        return n

    def kept(self, name):
        """ブック name に戻す値: 一覧 → 鍵 → {見出し: 値}"""
        return {canon: {k: x["cells"] for k, x in rows.items()} for canon, rows in (self.books.get(name) or {}).items()}

    def rows(self, name, canon):
        return (self.books.get(name) or {}).get(canon) or {}

    def save(self):
        """変わったときだけ書く（1つも無ければ、ファイルを作らない）"""
        now = json.dumps(self.books, ensure_ascii=False, sort_keys=True)
        if now != self.before and (self.books or self.path.exists()):
            save(self.path, {"version": 1, "books": self.books})
            self.before = now


def answer_gaps(ans, name, canon, now):
    """人が埋める欄の値のうち、(行が無くなったもの, 取り込んだあとで行の中身が変わったもの)。now: 鍵 → いまの行の中身"""
    rows = ans.rows(name, canon)
    gone = sorted(k for k in rows if k not in now)
    changed = sorted(k for k, x in rows.items() if k in now and x.get("seen") and x["seen"] != now[k])
    return gone, changed


class Times:
    """人が業務マトリクス一覧（xlsx）に入れた値の置き場（作業フォルダの times.json）: 処理時間（分）・件数・「処理時間の基準」の直し。
    build_matrix.py が、書き出す前にブックから取り込み、書き出しで戻す。人の値が正（機械は、ここに無い所だけ、仮の値を入れる）"""

    def __init__(self, work):
        self.path = pathlib.Path(work) / "times.json"
        d = load(self.path, {})
        if self.path.exists():                                       # 形の崩れた行を黙って捨てて上書きすると、人が入れた値が消える
            from validate import validate
            errs = validate(d, "times")
            if errs:
                raise BadJson(site(f"times.json の形が崩れている（{'／'.join(errs[:3])}）。人が入れた値を消さないよう止める。"
                              "times.json を直すか、値を入れたブックから取り込み直す（build_matrix.py --keep <ブック>）"))
        d = d if isinstance(d, dict) else {}
        ok = lambda x, *keys: (isinstance(x, dict) and all(isinstance(x.get(k), str) and x[k] for k in keys)
                               and isinstance(x.get("value"), (int, float)) and not isinstance(x.get("value"), bool))
        self.cells = {(x["flow"], x["node"]): dict(x) for x in d.get("cells") or [] if ok(x, "flow", "node")}
        self.counts = {x["flow"]: dict(x) for x in d.get("counts") or [] if ok(x, "flow")}
        self.table = {(x["action"], x["volume"]): dict(x) for x in d.get("table") or [] if ok(x, "action", "volume")}
        self.before = json.dumps(self.data(), ensure_ascii=False, sort_keys=True)     # 中の値を書き換えても、変わったと分かるように、文字で持つ

    def data(self):
        return {"version": 1, "cells": [self.cells[k] for k in sorted(self.cells)], "counts": [self.counts[k] for k in sorted(self.counts)],
                "table": [self.table[k] for k in sorted(self.table)]}

    def save(self):
        """変わったときだけ書く（1つも無ければ、ファイルを作らない）"""
        now = json.dumps(self.data(), ensure_ascii=False, sort_keys=True)
        if now != self.before and (self.cells or self.counts or self.table or self.path.exists()):
            save(self.path, self.data())
            self.before = now

    def gaps(self, flows):
        """人が入れた時間のうち、(箱が無くなったもの, 作業名が変わったもの)。どちらも [(フロー, ノード, 値, 入れたときの作業名, いまの作業名)]"""
        label = {(flow_key(fl), n.get("id")): n.get("label", "") for fl in flows for n in fl.get("nodes") or [] if isinstance(n, dict) and n.get("type") in STEP_TYPES}
        gone = [(f, i, x["value"], x.get("label", ""), "") for (f, i), x in sorted(self.cells.items()) if (f, i) not in label]
        renamed = [(f, i, x["value"], x["label"], label[(f, i)]) for (f, i), x in sorted(self.cells.items())
                   if (f, i) in label and x.get("label") and x["label"] != label[(f, i)]]
        return gone, renamed

    def rename(self, F, ren):
        """番号の振り直し（renumber.py）に合わせて、フロー F の値の番号も振り直す（ren: 前の番号 → 新しい番号）。返すのは、付け替えた数"""
        hit = {k: self.cells.pop(k) for k in [k for k in self.cells if k[0] == F and k[1] in ren]}
        for (f, i), x in hit.items():
            self.cells[(f, ren[i])] = dict(x, node=ren[i])
        return len(hit)

    def note(self, flows):
        """値ごとに、いまの担当の名前と、直前の箱の名前を書いておく（同じ作業名の箱が 2 つ以上あるフローで、番号がずれたときに、どの箱の値かを決める手がかり）。
        作業名が、いまの名前と同じ値だけ（名前の変わった値の手がかりは、入れたときのまま）"""
        clues = node_clues(flows)
        for k, x in self.cells.items():
            c = clues.get(k)
            if c is not None and x.get("label") == c[0]:
                x["lane"], x["after"] = c[1], list(c[2])

    def reconcile(self, flows):
        """フローを作り直して番号がずれたとき（直しの一覧を当て直した）: 入れたときの作業名・担当・直前の箱から、いま付く番号を決めて（locate）、付け替える。
        同じ作業名のノードがそのフローに 1 つだけなら、そこへ。2 つ以上なら、担当 → 直前の箱の順に絞る。返すのは [(フロー, 前の番号, いまの番号, 作業名, 値)]。
        いまの番号が None の行は、捨てた値（同じ作業に、もう値がある／付け替えた先の番号に、無くなった箱の値が残っていた）"""
        clues = node_clues(flows)
        label = {k: c[0] for k, c in clues.items()}
        want = {}
        for (F, i), x in self.cells.items():
            to = locate(clues, F, i, x)
            if to is not None and to != i:
                want.setdefault((F, to), []).append((F, i))
        moves = {src[0]: to for to, src in want.items() if len(src) == 1}      # 同じ行き先を、2 つの値が争うときは、動かさない
        taken, out = {k: self.cells.pop(k) for k in moves}, []
        for k, to in sorted(moves.items()):
            x, there = taken[k], self.cells.get(to)
            if there is not None and there.get("label") == label[to]:        # 同じ作業に、もう値がある（あとから入れたほう）。古いほうを捨てる
                out.append((k[0], k[1], None, x.get("label", ""), x["value"]))
                continue
            if there is not None:                                              # 無くなった箱の値が、この番号に残っていた
                out.append((to[0], to[1], None, there.get("label", ""), there["value"]))
            self.cells[to] = dict(x, node=to[1])
            out.append((k[0], k[1], to[1], x.get("label", ""), x["value"]))
        return out


def node_clues(flows):
    """作業のノードの、付け替えの手がかり: {(フロー, 番号): (作業名, 担当の名前, 直前のノードの名前の並び)}"""
    out = {}
    for fl in flows:
        F, nodes = flow_key(fl), {n.get("id"): n for n in fl.get("nodes") or [] if isinstance(n, dict)}
        lane = {x.get("id"): str(x.get("name") or "") for x in fl.get("lanes") or [] if isinstance(x, dict)}
        before = {}
        for e in fl.get("edges") or []:
            if isinstance(e, dict) and e.get("from") in nodes:
                a = nodes[e["from"]]
                before.setdefault(e.get("to"), set()).add(str(a.get("label") or NODE_LABEL.get(a.get("type"), "")))
        for i, n in nodes.items():
            if n.get("type") in STEP_TYPES:
                out[(F, i)] = (n.get("label", ""), lane.get(n.get("lane"), ""), sorted(before.get(i, ())))
    return out


def locate(clues, F, i, x):
    """人が入れた値 x（入れたときの作業名 label・担当 lane・直前の箱 after）が、いま付く番号。決められなければ None。
    同じ作業名のノードが、そのフローに 1 つだけなら、そこ。2 つ以上なら、担当 → 直前の箱の順に絞る（手がかりの無い値・合うものの無い手がかりは、絞らない）。
    絞り切れないとき、いまの番号 i が残っていれば、そのまま"""
    was = x.get("label")
    same = [k for (f, k), c in clues.items() if f == F and c[0] == was] if was else []
    for pos, key in ((1, "lane"), (2, "after")):
        if len(same) > 1 and x.get(key) is not None:
            same = [k for k in same if clues[(F, k)][pos] == x[key]] or same
    return same[0] if len(same) == 1 else (i if i in same else None)


def whole(v):
    """数を、整数で書けるなら整数に（7.0 → 7）"""
    return int(v) if float(v).is_integer() else round(float(v), 2)


def actions_of(v):
    """動作の種類（1 語か、2 語までの並び）を、決まった語の並びに。決まった語でないものは落とす"""
    return list(dict.fromkeys(x for x in (v if isinstance(v, list) else [v]) if x in ACTIONS))[:2]


def soft_formula(action, ref, col):
    """仮の値の式（基準の表を引く）。2 つの動きのある箱（確認する／承認する）は、2 つを引いて大きいほう"""
    one = [f'VLOOKUP("{a}",{ref},{col},FALSE)' for a in action.split(ACT_JOIN)]
    return one[0] if len(one) == 1 else f"MAX({','.join(one)})"


def time_basis(n, seen=0):
    """作業のノードの (動作の種類, 量, 量の決め方)。処理時間の仮の値は、この 2 つで表を引く。
    種類: 人の作業は action（無ければ「その他」。2 語なら「確認する／承認する」— 表で大きいほうを引く）。判断・システム処理・別のフローへは、種別から。
    量: volume があれば、それ（決め方は volume_reason）。無ければ、同じページの手順を 3 つ以上まとめた箱・見る帳票が 2 つ以上の箱は「大」、ほかは「中」。
    「特大」（時間・日の単位でかかる作業）は、機械は付けない。LLM が、資料からそう読めるときだけ、理由つきで付ける。
    seen は、その作業が見る帳票（direction が in の file）の数。作る帳票は数えない（見て 1 つ作るのは、ふつうの作業。伝票を作る箱だけが「大」になる）。
    まとめた手順は、出典の数では数えない（同じ作業が、概要・図・詳しい手順の 3 か所に書かれているだけの箱が「大」になる）。
    出典が引くカードの手順（item が s1・s2…）を、ページごとに数えた、いちばん多い数"""
    if n.get("type") != "process":
        return TYPE_ACTION.get(n.get("type"), ACTIONS[-1]), "中", ""
    act = ACT_JOIN.join(actions_of(n.get("action"))) or ACTIONS[-1]
    if n.get("volume") in VOLUMES:
        return act, n["volume"], str(n.get("volume_reason") or "").strip() or "指定"
    steps = {}
    for s in n.get("sources") or []:
        if isinstance(s, dict) and re.fullmatch(r"s\d+", str(s.get("item") or "")):
            steps.setdefault((s.get("doc_id"), s.get("page")), set()).add(s["item"])
    k = max((len(v) for v in steps.values()), default=0)
    if k >= 3:
        return act, "大", f"手順 {k} つ"
    return (act, "大", f"見る帳票 {seen} つ") if seen >= 2 else (act, "中", "")


def seen_files(fl):
    """作業のノードごとの、見る帳票（付け先がその作業で、direction が in の file）の数"""
    return collections.Counter(n.get("attach_to") for n in fl.get("nodes") or [] if isinstance(n, dict) and n.get("type") == "file" and n.get("direction") == "in")


def flow_key(fl):
    """フローの id。flow_id が無い・文字でないときは、ファイル名から"""
    i = fl.get("flow_id")
    return i if isinstance(i, str) and i else str(fl.get("_file", "")).split(".")[0]


def load_flows(work, errors=None):
    """flows/*.flow.json を、taxonomy の並び順（無ければファイル名順）で返す。
    読めないファイル・同じ flow_id のファイルは、errors（リスト）が渡されていれば (ファイル名, 理由) を入れて飛ばす（点検が FAIL にして先へ進む）。
    渡されていなければ、ファイル名を言って止まる（1本でも欠けたまま図や報告書を作らない）"""
    flows = {}
    for f in sorted((work / "flows").glob("*.flow.json")):
        fl, err = try_load(f)
        if err is not None:
            err = f"JSON として読めない（{err}）"
        elif not isinstance(fl, dict):
            err = "全体がオブジェクト（{ … }）になっていない"
        if err is None:
            fl["_file"] = f.name
            if flow_key(fl) in flows:
                err = f"flow_id「{flow_key(fl)}」が {flows[flow_key(fl)]['_file']} と重複している"
        if err is not None:
            if errors is None:
                sys.exit(f"flows/{f.name}: {err}。先に python3 scripts/check_flow.py -w <作業フォルダ> --stage flow を通す")
            errors.append((f.name, err))
            continue
        flows[flow_key(fl)] = fl
    tax = try_load(work / "taxonomy.json")[0]
    cats = tax.get("categories") if isinstance(tax, dict) else None
    order = [c["id"] for c in cats if isinstance(c, dict) and isinstance(c.get("id"), str) and c["id"] in flows] if isinstance(cats, list) else []
    return [flows[i] for i in order] + [flows[i] for i in flows if i not in order]


def unused_reasons(work, flows):
    """使わない理由のあるページ: {(資料, page): 理由}。分類の対象外（taxonomy の unassigned_pages）と、フローの unused_pages。
    点検・出典対応表・照合シートが同じ見方をするよう、ここ1か所で決める。項目ごとの対象外（excluded_items）は item_reasons"""
    tax = try_load(work / "taxonomy.json")[0]
    out = {}
    for u in (tax.get("unassigned_pages") if isinstance(tax, dict) else None) or []:
        if isinstance(u, dict):
            out[(u.get("doc_id"), u.get("page"))] = str(u.get("reason") or "")
    for fl in flows:
        for u in fl.get("unused_pages") or []:
            if isinstance(u, dict):
                out.setdefault((u.get("doc_id"), u.get("page")), str(u.get("reason") or ""))
    return out


def cited_pages(flows):
    """出典になっているページ: {(資料, page): [フローの id]}"""
    out = {}
    for fl in flows:
        for n in fl.get("nodes") or []:
            for s in (n.get("sources") or []) if isinstance(n, dict) else []:
                if isinstance(s, dict):
                    users = out.setdefault((s.get("doc_id"), s.get("page")), [])
                    if flow_key(fl) not in users:
                        users.append(flow_key(fl))
    return out


def why_unused(k, reasons, items):
    """出典になっていないページ k の、使わない理由の書き方。理由が無ければ空（点検では FAIL になるもの）"""
    if reasons.get(k, "").strip():
        return reasons[k].strip()
    if any(key[:2] == k and why.strip() for key, (_f, why) in items.items()):
        return "項目ごとに対象外の理由がある（「フローに入れなかった手順・分岐」の表）"
    return ""


def item_reasons(flows):
    """フローに入れなかった項目: {(資料, page, item): (フロー, 理由)}（excluded_items）"""
    out = {}
    for fl in flows:
        for x in fl.get("excluded_items") or []:
            if isinstance(x, dict):
                out.setdefault((x.get("doc_id"), x.get("page"), x.get("item")), (flow_key(fl), str(x.get("reason") or "")))
    return out


def added_category_questions(tax):
    """目次に無く足した分類（目次に準拠した分類の中の basis: llm）ごとの、決まった形の確認事項（決まり 2.21）。LLM がその分類を指す問いを書いていれば、作らない"""
    if tax.get("basis") not in ("toc", "mixed"):
        return []
    cats = [c for c in tax.get("categories") or [] if isinstance(c, dict)]
    asked = {q.get("flow") for q in tax.get("questions") or [] if isinstance(q, dict) and "分類" in str(q.get("category", ""))}
    out = []
    for c in cats:
        if c.get("basis") == "llm" and c.get("id") not in asked:
            ref = c.get("name_ref") if isinstance(c.get("name_ref"), dict) else {}
            src = [{"doc_id": ref["doc_id"], "page": ref["page"]}] if ref.get("doc_id") and isinstance(ref.get("page"), int) else \
                [{"doc_id": p["doc_id"], "page": p["pages"][0]} for p in c.get("pages") or [] if isinstance(p, dict) and p.get("pages")][:1]
            out.append((f"m-added-{c.get('id')}", {"category": "分類", "effect": "フローの範囲", "flow": c.get("id"),
                        "question": f"目次に無い「{c.get('name')}」を、1つの業務分類として足しました。このまま 1 つの業務としてよいですか、目次のどれかに含めますか。",
                        "assumption": "目次の同じ段の末尾に、業務分類として足した", "sources": src, "ask": ""}))
    return out


def collect_questions(work, write=True):
    """全工程の確認事項を1つに集め、通し番号（Q-001…）を振って questions.json に保存する（write=False なら、保存せずに返すだけ。点検用）。
    番号は一度振ったら変えない（出どころと元の id で引き継ぐ）。消えた確認事項の番号は再利用しない。
    いま無い確認事項（書き直している途中のフローの分など）の番号も、対応表（assigned）に残す。戻ってきたら同じ番号になる。"""
    old = load(work / "questions.json")                             # JSON として読めなければ止まる（番号を Q-001 から振り直すと、お客様に渡した番号と食い違う）
    if old is None:
        old = {"next": 1, "questions": [], "assigned": {}}
    else:
        from validate import validate
        errs = validate(old, "questions")
        if errs:
            raise BadJson(f"questions.json の形が崩れている（{'／'.join(errs[:3])}）。Q 番号を振り直さないよう止める。"
                          "直せないときだけ、questions.json を消す（Q-001 から振り直す。お客様に渡した番号と食い違う）")
    known = {tuple(k.split("|", 1)): v for k, v in (old.get("assigned") or {}).items() if "|" in k} if isinstance(old.get("assigned"), dict) else {}
    known.update({(q["origin"], q["local"]): q["id"] for q in old["questions"] if isinstance(q, dict) and {"origin", "local", "id"} <= set(q)})
    nxt = old.get("next", 1)
    items = []

    def add(origin, local, q, fixed=False, **extra):
        n = sum(1 for x in items if x["origin"] == origin and x["local"].split("#")[0] == str(local))
        local = str(local) if not n else f"{local}#{n + 1}"          # 同じファイルの中で id が重なっても、別の番号にする（点検は重なりを FAIL にする）
        # 機械が付けた重要度（資料・台帳の決まった問い、apply_edits・activity.py が決まり 5.4 で付けたもの）は、そのまま。
        # LLM が書くファイル（階層・用語・分類・台帳）の問いは、ここで決まり 5.4 から決める（書いてあっても使わない）
        sev = q.get("severity") if fixed and q.get("severity") in SEV_LABEL else severity_of(q)
        item = {"origin": origin, "local": str(local), "severity": sev, "category": q.get("category", ""), "effect": q.get("effect", ""),
                "flow": extra.get("flow"), "node": q.get("node"), "nodes": q_targets(q), "edge": q.get("edge"), "question": q.get("question", ""),
                "assumption": q.get("assumption", ""), "sources": q.get("sources") or [], "ask": q.get("ask", "")}
        items.append(item)

    m = load(work / "manifest.json", {"docs": []})
    title = labels(m, work).title if m["docs"] else str                # 資料は題名で書く（ファイルとの対応は「資料一覧」）
    for d in m["docs"]:
        if d.get("duplicate_of"):
            continue
        if not d.get("page_list") and d.get("read_via"):             # Access のファイル。書き出したテキストのほうを読んでいる
            continue
        if not d.get("page_list"):
            why = d.get("render_note") or d.get("note") or "読めていない"
            add("inventory", f"{d['doc_id']}:unread", fixed=True, q={"severity": "high", "category": "読取不能", "effect": "箱",
                "question": f"資料「{title(d['doc_id'])}」を読めていない（{why}）。別の形式で提供できるか。",
                "assumption": "この資料の内容はフローに反映していない", "sources": [{"doc_id": d["doc_id"]}]})
        elif d.get("render_note"):
            add("inventory", f"{d['doc_id']}:partial", fixed=True, q={"severity": "mid", "category": "読取不能", "effect": "箱",
                "question": f"資料「{title(d['doc_id'])}」の一部を読めていない（{d['render_note']}）。",
                "assumption": "読めた範囲だけをフローに反映した", "sources": [{"doc_id": d["doc_id"]}]})
    for d in m["docs"]:                                              # 非表示のスライド・シートは、現行の手順とは限らない。読んだことと扱いを利用者に知らせる
        hidden = [f"スライド {x}" for x in d.get("hidden_slides") or []] + [f"「{x}」シート" for x in d.get("hidden_sheets") or []]
        if hidden and not d.get("duplicate_of"):
            add("inventory", f"{d['doc_id']}:hidden", fixed=True, q={"severity": "low", "category": "資料の版", "effect": "箱",
                "question": f"資料「{title(d['doc_id'])}」に非表示のものがある（{'、'.join(hidden)}）。現行の手順に含めるか。",
                "assumption": "読みはしたが、現行の手順かどうかは分からない。フローに入れなかったものは、出典対応表の「使っていない理由」に書いた",
                "sources": [{"doc_id": d["doc_id"]}]})
    seen = set()
    for d in m["docs"]:
        g = tuple(d.get("version_group") or [])
        if g and g not in seen:
            seen.add(g)
            names = "、".join(title(x["doc_id"]) for x in m["docs"] if x["doc_id"] in g)
            add("inventory", "version:" + "/".join(g), fixed=True, q={"severity": "mid", "category": "資料の版", "effect": "箱",
                "question": f"版違いらしい資料がある（{names}）。どれが現行か。", "assumption": "すべて読み、食い違いは別の確認事項にした",
                "sources": [{"doc_id": x} for x in g]})
    for origin, path in (("hierarchy", work / "structure" / "hierarchy.json"), ("glossary", work / "glossary.json"),
                         ("taxonomy", work / "taxonomy.json"), ("systems", work / "structure" / "systems.json")):
        for i, q in enumerate((load(path, {}) or {}).get("questions") or []):
            add(origin, q.get("id") or i + 1, q, flow=q.get("flow") if origin == "taxonomy" else None)   # 分類を指す問いは、その行に出す
    for local, q in added_category_questions(load(work / "taxonomy.json", {}) or {}):
        add("taxonomy", local, q, flow=q["flow"])
    if (work / "structure" / "systems.json").exists():               # システム台帳とプログラムの構造から、機械が作る確認事項
        from systems import machine_questions
        for local, q in machine_questions(work, m):
            add("systems", local, q, fixed=True)
    if (work / "activities").exists():                               # アクティビティ図の確認事項（図の番号を、フローの欄に入れる）
        for f in sorted((work / "activities").glob("*.activity.json")):
            act = try_load(f)[0]
            for i, q in enumerate((act.get("questions") if isinstance(act, dict) else None) or []):
                add(f"activity:{act.get('flow_id')}", q.get("id") or i + 1, q, fixed=True, flow=act.get("flow_id"))
    for fl in load_flows(work):
        for i, q in enumerate(fl.get("questions") or []):
            add(f"flow:{fl['flow_id']}", q.get("id") or i + 1, q, fixed=True, flow=fl["flow_id"])
    for it in items:
        key = (it["origin"], it["local"])
        if key not in known:
            known[key] = f"Q-{nxt:03d}"
            nxt += 1
        it["id"] = known[key]
    items.sort(key=lambda x: (int(re.sub(r"\D", "", x["id"]) or 0), x["id"]))
    if write:
        save(work / "questions.json", {"next": nxt, "questions": items, "assigned": {f"{o}|{l}": i for (o, l), i in sorted(known.items())}})
    return items


def order_lanes(names, pairs, keep=(), right=()):
    """レーンの並び（決まり 4.2）。keep: 左から、この順に固定する名前。right: 右端に置く名前（出てくる順）。
    残りは、線がまたぐレーンの数の合計（線の本数 × 位置の差）が最小になる順。同じなら、出てくる順（names の順）。
    7 つまでは全部の順を試し、多ければ、隣どうしの入れ替えで詰める。決まりなので、毎回同じ並びになる"""
    import itertools
    head = list(dict.fromkeys(x for x in keep if x in names))
    tail = [x for x in names if x in right and x not in head]
    mid = [x for x in names if x not in head and x not in tail]
    w = collections.Counter(tuple(sorted((a, b))) for a, b in pairs if a != b and a in names and b in names)

    def cost(seq):
        pos = {x: i for i, x in enumerate(seq)}
        return sum(n * abs(pos[a] - pos[b]) for (a, b), n in w.items())
    best, low = mid, cost(head + mid + tail)
    if len(mid) <= 7:
        for p in itertools.permutations(mid):
            c = cost(head + list(p) + tail)
            if c < low:
                best, low = list(p), c
    else:
        moved = True
        while moved:
            moved = False
            for i in range(len(best) - 1):
                trial = best[:i] + [best[i + 1], best[i]] + best[i + 2:]
                c = cost(head + trial + tail)
                if c < low:
                    best, low, moved = trial, c, True
    return head + best + tail


def flow_heads(c):
    """業務分類（taxonomy.json の categories[]）から写す、フローの見出しの項目。apply_edits.py が作るとき・renumber.py が付いていくときに写し、
    check_flow.py が分類と同じかを見る（図の見出し・業務マトリクス一覧の行に出る）"""
    return {"name": c.get("name"), "number": str(c.get("number") or ""), "parent": c.get("parent"), "level": c.get("level", 1), "basis": c.get("basis", "toc")}


def label_core(label):
    """箱の名前から、末尾のかっこ書き（最後に置いた目的語・補足）と句点を除いたもの"""
    t = str(label or "").strip().rstrip("。")
    while t and t[-1] in "）)":
        depth = 0
        for i in range(len(t) - 1, -1, -1):
            depth += (t[i] in "）)") - (t[i] in "（(")
            if depth == 0:
                break
        else:
            return t                                                 # かっこが対になっていない
        t = t[:i].rstrip().rstrip("。")
    return t


def ends_with_verb(label):
    """箱の名前が、動詞の終止形で終わるか（「何を・どうする」。末尾のかっこ書きは除いて見る）"""
    return label_core(label)[-1:] in VERB_END


def ends_as_question(label):
    """判断の名前が、問いの形（「〜か」）で終わるか"""
    return label_core(label).rstrip("？?").endswith("か")


ELSE = "その他"                                                      # 分岐の、書かれていない側のラベル（反対の語が表に無いとき。branch で直せる）
# 分岐の書かれていない側のラベル（決まり 3.4）: 書かれた側のラベルの、反対の語。語全体の対と、語の終わりの対（両向きに引く。長い語から）
OPPOSITE_WORDS = {"はい": "いいえ", "Yes": "No", "YES": "NO", "OK": "NG", "○": "×", "〇": "×", "可": "不可", "有": "無", "済": "未済"}
OPPOSITE_ENDS = [("があれば", "がなければ"), ("がある", "がない"), ("はある", "はない"), ("あり", "なし"), ("有り", "無し"), ("すれば", "しなければ"), ("する", "しない"),
                 ("できる", "できない"), ("以上", "未満"), ("一致", "不一致"), ("可能", "不可能"), ("正常", "異常"), ("適正", "不適正"), ("完了", "未完了"), ("必要", "不要")]
_WORDS = dict(OPPOSITE_WORDS, **{b: a for a, b in OPPOSITE_WORDS.items() if b not in OPPOSITE_WORDS})
_ENDS = sorted(OPPOSITE_ENDS + [(b, a) for a, b in OPPOSITE_ENDS], key=lambda p: -len(p[0]))


def opposite(label):
    """書かれた側のラベル → 反対の側のラベル（「不備がある」→「不備がない」・「あり」→「なし」）。表に無ければ None（「その他」にする）"""
    t = str(label or "").strip()
    if t in _WORDS:
        return _WORDS[t]
    return next((t[:-len(a)] + b for a, b in _ENDS if t.endswith(a) and len(t) >= len(a)), None)


def ask_form(label):
    """判断の名前の「？」を「か」にする（決まり 3.4: 問いの形は「〜か」。「〜か？」は「？」を外す）"""
    t = str(label or "").strip()
    if not re.search(r"[？?]$", t):
        return t
    t = re.sub(r"[？?]+$", "", t).rstrip()
    return t if t.endswith("か") else t + "か"


def q_targets(q):
    """確認事項が指すノードの id（`node` と `nodes` を合わせ、重なりを除いた並び）"""
    one = [q["node"]] if isinstance(q.get("node"), str) and q["node"] else []
    many = [x for x in q["nodes"] if isinstance(x, str) and x] if isinstance(q.get("nodes"), list) else []
    return list(dict.fromkeys(one + many))


ICHIDAN_KANJI = set("見得出経寝着似居煮干射来")                    # 漢字 1 字＋「る」で、「〜た」になる動詞（ほかの漢字＋「る」は「〜った」: 送る・作る・取る）


def when_done(label):
    """作業の名前（「〜する」「〜を行う」の形）→ 終了条件の「〜したとき」（決まり 3.6）。動詞で終わらない名前は「「…」を終えたとき」"""
    t = re.sub(r"[。．.\s]+$", "", str(label or "").strip())
    if not t:
        return ""
    last, prev = t[-1], t[-2:-1]
    if t.endswith("する"):
        return t[:-2] + "したとき"
    if t.endswith("行く"):
        return t[:-1] + "ったとき"
    tail = {"う": "った", "つ": "った", "く": "いた", "ぐ": "いだ", "す": "した", "む": "んだ", "ぶ": "んだ", "ぬ": "んだ"}.get(last)
    if last == "る":                                                 # え・い段の仮名＋る（受ける・確かめる）と、見る・得る は「〜た」。ほか（送る・終わる）は「〜った」
        tail = "た" if prev in "えけげせぜてでねへべぺめれいきぎしじちぢにひびぴみり" or prev in ICHIDAN_KANJI else "った"
    return t[:-1] + tail + "とき" if tail and re.match(r"[ぁ-ん一-龥]", prev or "") else f"「{t}」を終えたとき"


def severity_of(q):
    """確認事項の重要度を、決まり 5.4 で機械が決める（LLM には書かせない。書いてあっても使わない）。
    答えで変わる所（effect）と、指す箱の数から: 担当 → 中／システム・帳票 → 低／フローの範囲 → 高（指す所が開始・終了だけなら、
    開始条件・終了条件の問いなので低）／箱・順序・分岐 → 指す箱が 2 つ以上なら高、ほか（箱 1 つ・線 1 本・指す所なし）は中"""
    eff, ids = q.get("effect"), q_targets(q)
    if eff == "担当":
        return "mid"
    if eff == "システム・帳票":
        return "low"
    if eff == "フローの範囲":
        return "low" if ids and all(re.fullmatch(r"(start|end)[0-9]*(-[0-9]+)?", i) for i in ids) else "high"
    return "high" if len(ids) >= 2 else "mid"


ENDS_ASKED = {"start": ("m-start", "start_condition", "この業務の開始条件（いつ・何を受けて始めるか）は何ですか。", "開始条件は空のままにした"),
              "end": ("m-end", "end_condition", "この業務の終了条件（何ができたら終わるか）は何ですか。", "終了条件は空のままにした")}


def settle_flow(fl, ask=True):
    """機械が決まりで決める所を、flow.json に当てる（apply_edits.py・renumber.py が呼ぶ。何度当てても同じ）。変えた所の説明の並びを返す:
    判断の名前の「？」→「か」（3.4）／書かれていない側のラベル「その他」→ 書かれた側の反対の語（3.4）／空の終了条件 → 終了に入る最後の作業の「〜したとき」（3.6）／
    それでも空の開始条件・終了条件に、開始・終了を指す確認事項が無ければ、機械の問い m-start・m-end（3.6。ask=False なら足さない — 直しの一覧の道は、問いを自分で作る）／
    確認事項の重要度（5.4）"""
    notes, nodes = [], [n for n in fl.get("nodes") or [] if isinstance(n, dict)]
    edges = [e for e in fl.get("edges") or [] if isinstance(e, dict)]
    for n in nodes:
        if n.get("type") == "decision" and ask_form(n.get("label")) != n.get("label"):
            notes.append(f"{n.get('id')}: 判断の名前を「〜か」にした（「{n.get('label')}」→「{ask_form(n.get('label'))}」）")
            n["label"] = ask_form(n.get("label"))
        outs = [e for e in edges if e.get("from") == n.get("id")] if n.get("type") == "decision" else []
        other = [e for e in outs if e.get("label") == ELSE and e.get("inferred")]
        said = [e for e in outs if e.get("label") and e.get("label") != ELSE]
        if len(other) == 1 and len(said) == 1 and opposite(said[0]["label"]):
            other[0]["label"] = opposite(said[0]["label"])
            notes.append(f"{n.get('id')}: 書かれていない側のラベルを「{other[0]['label']}」にした（書かれた側「{said[0]['label']}」の反対の語）")
    if not str(fl.get("end_condition") or "").strip():
        last = {e.get("from") for e in edges if e.get("to") == "end" and not e.get("inferred")}
        tail = next((n for n in nodes if n.get("id") in last), None) if len(last) == 1 else None
        if tail and tail.get("type") in ("process", "system", "subprocess"):
            fl["end_condition"] = when_done(tail.get("label"))
            notes.append(f"終了条件が空なので、終了に入る最後の作業「{str(tail.get('label'))[:30]}」から「{fl['end_condition']}」にした（原典に成果が書かれていれば、それを書く）")
    qs = fl.setdefault("questions", []) if isinstance(fl.get("questions", []), list) else []
    for end, (qid, key, text, assume) in ENDS_ASKED.items() if ask else ():
        if str(fl.get(key) or "").strip() or any(isinstance(q, dict) and end in q_targets(q) for q in qs) or not any(n.get("id") == end for n in nodes):
            continue
        src = next((s for n in nodes for s in n.get("sources") or [] if isinstance(s, dict) and s.get("doc_id")), None)
        lanes = [x.get("name") for x in fl.get("lanes") or [] if isinstance(x, dict) and x.get("name")]
        q = {"id": qid, "severity": "", "category": "記載なし", "effect": "フローの範囲", "node": end, "question": text, "assumption": assume,
             "sources": [{"doc_id": src["doc_id"], "page": src.get("page")}] if src else [], "ask": lanes[0] if lanes else ""}
        q["severity"] = severity_of(q)
        qs.append(q)
        notes.append(f"{'開始' if end == 'start' else '終了'}条件が空で、それを聞く確認事項も無いので、機械の問い {qid} を足した")
    for q in fl.get("questions") or []:
        if isinstance(q, dict) and q.get("effect") and q.get("severity") != severity_of(q):
            notes.append(f"確認事項 {q.get('id')}: 重要度を、決まり 5.4 で{SEV_LABEL[severity_of(q)]}にした（前は「{SEV_LABEL.get(q.get('severity'), q.get('severity'))}」）")
            q["severity"] = severity_of(q)
    return notes


def question_ids(questions, flow_id):
    """フロー flow_id について、ノード → 確認事項番号、(from, to) → 確認事項番号 の対応を返す"""
    by_node, by_edge = {}, {}
    for q in questions:
        if q.get("flow") != flow_id:
            continue
        for nid in q_targets(q):
            by_node.setdefault(nid, []).append(q["id"])
        if q.get("edge") and len(q["edge"]) == 2:
            by_edge.setdefault(tuple(q["edge"]), []).append(q["id"])
    return by_node, by_edge


# ---------- ノードの並びと明細 ----------

def order_of(fl):
    """作業のノード（ファイル以外）の、時系列の深さ {id: 深さ}。開始から線をたどった深さで、戻りの線は数えない"""
    main = [n["id"] for n in fl["nodes"] if n.get("type") != "file"]
    out = {i: [] for i in main}
    for e in fl.get("edges") or []:
        if e.get("from") in out and e.get("to") in out:
            out[e["from"]].append(e["to"])
    depth, state = {}, {}

    def visit(u, d):                                                 # 深さ優先でたどり、たどっている途中のノードへ戻る線（差し戻しなど）は数えない
        if state.get(u) == 1:
            return
        if depth.get(u, -1) >= d and state.get(u) == 2:
            return
        depth[u] = max(depth.get(u, -1), d)
        state[u] = 1
        for v in out[u]:
            visit(v, depth[u] + 1)
        state[u] = 2
    for s in [n["id"] for n in fl["nodes"] if n.get("type") == "start"] or main[:1]:
        visit(s, 0)
    last = max(depth.values(), default=0)
    for i in main:                                                   # 開始から届かないノード（点検が FAIL にする）は末尾へ
        depth.setdefault(i, last + 1)
    return depth


WORK_TYPES = ("process", "decision", "system")                       # カードの手順・分岐を「作業として描いた」と数えるノード


def infer_nodes(fl):
    """作業のノードの推定の印（inferred）を、出典から機械で決める: カードの項目（item）を1つも引かないノードが、原典に無いノード＝推定。
    LLM に印を付けさせると、同じ種類のノードで有無が割れる。変えたノードを [(id, 新しい印)] で返す"""
    changed = []
    for n in fl.get("nodes") or []:
        if not isinstance(n, dict) or n.get("type") not in WORK_TYPES:
            continue
        want = not any(isinstance(x, dict) and x.get("item") for x in n.get("sources") or [])
        if bool(n.get("inferred")) != want:
            changed.append((n.get("id"), want))
            if want:
                n["inferred"] = True
            else:
                n.pop("inferred", None)
    return changed


# ---------- システム（札） ----------

TAGGED_TYPES = ("process", "system")                                 # 札（その作業で使うシステムの名前）を付けるノード
SYSTEM_SEP = re.compile(r"[、,，・／/]")


def system_names(work):
    """用語の正本のシステム: {正規化した呼び方: 正本の名前}（別の表記を含む）"""
    g = try_load(pathlib.Path(work) / "glossary.json")[0]
    out = {}
    for t in (g.get("terms") if isinstance(g, dict) else None) or []:
        if isinstance(t, dict) and t.get("kind") == "system" and isinstance(t.get("name"), str) and t["name"]:
            for k in [t["name"]] + [x for x in (t.get("aliases") if isinstance(t.get("aliases"), list) else []) if isinstance(x, str)]:
                out[norm(k)] = t["name"]
    return out


def systems_in(text, names):
    """カードの手順の system（文字）→ システムの名前の並び。用語の正本にあれば、正本の名前。無ければ、書かれた文字のまま（点検が知らせる）"""
    text = str(text or "").strip()
    if not text:
        return []
    if norm(text) in names:
        return [names[norm(text)]]
    parts = [x.strip() for x in SYSTEM_SEP.split(text) if x.strip()]
    if len(parts) > 1 and all(norm(x) in names for x in parts):      # 「A・B」と並べて書かれている
        return list(dict.fromkeys(names[norm(x)] for x in parts))
    return [text]


def systems_named(text, names):
    """文の中に出てくる、用語の正本のシステムの名前（出てきた順）。長い名前から当て、当てた所は短い名前に当てない（「請求システム等」の中の「請求システム」）"""
    t, found = norm(text), []
    for k in sorted(names, key=len, reverse=True):
        at = t.find(k) if k else -1
        if at >= 0:
            found.append((at, names[k]))
            t = t.replace(k, "\0" * len(k))
    return list(dict.fromkeys(nm for _at, nm in sorted(found)))


def node_systems(n, lane_name, items_of, names):
    """ノードの札の、あるべき中身: 出典に引いたカードの手順の system（用語の正本の名前にしたもの）と、その手順の引用・作業の文に出てくる
    用語の正本のシステム名（引いた順。重なりは除く）。自分のレーンと同じ名前は入れない（レーンが示している）"""
    out = []
    if n.get("type") not in TAGGED_TYPES:
        return out
    for s in n.get("sources") or []:
        if isinstance(s, dict) and isinstance(s.get("item"), str) and isinstance(s.get("page"), int):
            it = items_of(s.get("doc_id"), s["page"]).get(s["item"], {})
            if it.get("kind") != "step":
                continue
            for x in systems_in(it.get("system"), names) + systems_named(f"{it.get('quote', '')}\n{it.get('text', '')}", names):
                if x not in out and norm(x) != norm(lane_name):
                    out.append(x)
    return out


def tag_systems(fl, items_of, names):
    """作業のノードの札（systems）を、出典から機械で付ける。LLM に書かせると、補足の文に入れたり入れなかったりして、
    システムを使う作業なのに図に名前が出ないノードが出る。変えたノードを [(id, 新しい札)] で返す"""
    lane = {l.get("id"): l.get("name", "") for l in fl.get("lanes") or [] if isinstance(l, dict)}
    changed = []
    for n in fl.get("nodes") or []:
        if not isinstance(n, dict):
            continue
        want = node_systems(n, lane.get(n.get("lane"), ""), items_of, names)
        if (n.get("systems") or []) != want:
            changed.append((n.get("id"), want))
            if want:
                n["systems"] = want
            else:
                n.pop("systems", None)
        own = {norm(x) for x in want} | {k for k, v in names.items() if v in want}
        parts = [x for x in str(n.get("detail") or "").split("・") if x]
        if own and any(norm(x) in own for x in parts):               # 補足に、札と同じシステム名だけの区切りがある（札が無かった版の書き方）。重ねて出さない
            n["detail"] = "・".join(x for x in parts if norm(x) not in own)
            if not n["detail"]:
                del n["detail"]
            if (n.get("id"), want) not in changed:
                changed.append((n.get("id"), want))
    return changed


NOT_IN_FLOW = "業務フローに出てこない（マニュアルに、使う作業が見つからない）"
NO_MANUAL = "業務フローは作っていない（マニュアルが入力に無い）"


def unlinked_note(m):
    """業務フローに紐付いていない起動点の備考。マニュアルが入力に無いときは、そう書く"""
    return NOT_IN_FLOW if any(True for _x in all_pages(m)) else NO_MANUAL


SYSTEM_KIND = {"core": "基幹・パッケージ", "euc_excel": "EUC（Excel）", "euc_access": "EUC（Access）", "script": "スクリプト・バッチ",
               "external": "外部サービス", "other": "その他"}


def load_systems(work):
    """システム台帳（structure/systems.json）。無ければ・読めなければ空"""
    x = try_load(pathlib.Path(work) / "structure" / "systems.json")[0]
    return x if isinstance(x, dict) else {}


def system_ledger(work, flows):
    """システムの一覧（業務マトリクス一覧の「システム一覧」と、点検が使う）。1件＝
    {"id", "name", "aliases", "kind", "docs", "tags": [(フロー, ノード)], "boxes": [(フロー, ノード, アクティビティ図の番号)], "known"}
    並びと番号は、システム台帳（structure/systems.json）。台帳が無ければ、用語の正本のシステムを書かれた順に並べて SYS-01… を振る。
    台帳にも用語の正本にも無い名前（札・レーンにだけ出てくる）は、番号なしで末尾（known: False。点検が知らせる）"""
    g = try_load(pathlib.Path(work) / "glossary.json")[0]
    terms = [t for t in (g.get("terms") if isinstance(g, dict) else None) or [] if isinstance(t, dict) and t.get("kind") == "system" and t.get("name")]
    alias = {t["name"]: [x for x in t.get("aliases") or [] if isinstance(x, str)] for t in terms}
    listed = [x for x in load_systems(work).get("systems") or [] if isinstance(x, dict) and x.get("name")]
    out = [{"id": str(x.get("id") or ""), "name": x["name"], "kind": str(x.get("kind") or ""), "docs": list(x.get("docs") or [])} for x in listed]
    have = {norm(x["name"]) for x in out}
    for k, t in enumerate([t for t in terms if norm(t["name"]) not in have], 1):
        out.append({"id": "" if listed else f"SYS-{k:02d}", "name": t["name"], "kind": "", "docs": []})
    by = {norm(x["name"]): x for x in out}
    for x in out:
        x.update(aliases=alias.get(x["name"], []), tags=[], boxes=[], known=True)
    for fl in flows:
        lane = {l.get("id"): l for l in fl.get("lanes") or [] if isinstance(l, dict)}

        def row(name):
            if norm(name) not in by:
                by[norm(name)] = {"id": "", "name": name, "kind": "", "docs": [], "aliases": [], "tags": [], "boxes": [], "known": False}
                out.append(by[norm(name)])
            return by[norm(name)]
        for n in fl.get("nodes") or []:
            if not isinstance(n, dict):
                continue
            for name in n.get("systems") or []:
                row(name)["tags"].append((flow_key(fl), n.get("id")))
            l = lane.get(n.get("lane")) or {}
            if n.get("type") == "system" and l.get("kind") == "system":
                row(l.get("name", ""))["boxes"].append((flow_key(fl), n.get("id"), n.get("activity") or ""))
    return out


def step_order_by_branch(fl):
    """作業のノード（ファイル以外。開始・終了を含む）を、枝ごとに続けて並べた順。横に読む表（業務マトリクス）の列の順に使う。
    開始から線をたどり、前の作業がすべて並んだノードだけを置く（合流は、どの枝よりもあと）。戻りの線は数えない。枝の順は edges に書いた順。
    深さの順（order_of。番号の振り直しに使う）は、排他の枝のノードが交互に並ぶので、左隣が前の作業にならない"""
    depth = order_of(fl)
    succ, pred = {i: [] for i in depth}, {i: [] for i in depth}
    for e in fl.get("edges") or []:
        u, v = e.get("from"), e.get("to")
        if u in succ and v in succ and depth[v] > depth[u] and v not in succ[u]:
            succ[u].append(v)
            pred[v].append(u)
    done, out = set(), []
    todo = [n["id"] for n in fl["nodes"] if n.get("type") == "start"][::-1] or list(depth)[:1]
    while todo:                                                      # 深さ優先（再帰を使わない。ノードの多いフローでも落ちない）
        u = todo.pop()
        if u in done or any(p not in done for p in pred[u]):
            continue
        done.add(u)
        out.append(u)
        todo += succ[u][::-1]
    return out + [i for i in depth if i not in done]                 # 開始から届かないノードは、nodes に書いた順で末尾へ


def category_path(cats, cid):
    """業務分類 cid の、最上位からの並び [c0, …, ck]（ck が cid の分類）。親が見つからない・輪になっているときは、たどれた所まで"""
    by = {c.get("id"): c for c in cats or [] if isinstance(c, dict)}
    out, seen = [], set()
    while cid in by and cid not in seen:
        seen.add(cid)
        out.append(by[cid])
        cid = by[cid].get("parent")
    return out[::-1]


def own_number(cat, parent=None):
    """分類の番号のうち、その段だけの部分（「第8章 3 (1)」→ 親が「第8章 3」なら「(1)」）"""
    num, up = str(cat.get("number") or "").strip(), str((parent or {}).get("number") or "").strip()
    return num[len(up):].strip() if up and num.startswith(up) else num


class CardItems:
    """(資料, page) → そのページのカードの項目（card_items）。1回読んだら覚えておく"""

    def __init__(self, work):
        self.work, self.cache = work, {}

    def __call__(self, doc, page):
        if (doc, page) not in self.cache:
            self.cache[(doc, page)] = card_items(load_card(self.work, doc, page)) if isinstance(page, int) else {}
        return self.cache[(doc, page)]


def node_details(fl, m, questions, items_of):
    """フロー1本の、開始・終了を除くノードの明細（nodes に書いた順）。出典対応表（build_reports.py）と
    業務マトリクスのノード一覧（build_matrix.py）が同じ明細を使うので、出典・引用・印の書き方が食い違わない"""
    q_node, _ = question_ids(questions, flow_key(fl))
    lane = {l["id"]: l["name"] for l in fl.get("lanes") or []}
    out = []
    for n in fl["nodes"]:
        if n["type"] in ("start", "end"):
            continue
        ss = n.get("sources") or []
        out.append({"id": n["id"], "type": n["type"], "lane": lane.get(n.get("lane"), ""), "label": n.get("label", ""),
                    "detail": n.get("detail", ""), "reference": n.get("reference", ""), "systems": list(n.get("systems") or []),
                    "sources": [src_label(m, s) for s in ss],
                    "quote": " ／ ".join(("〔画像〕" if s.get("from_image") else "") + s.get("quote", "") for s in ss),
                    "inferred": bool(n.get("inferred")),
                    "note": any(items_of(s.get("doc_id"), s.get("page")).get(s.get("item"), {}).get("origin") == "note" for s in ss),
                    "questions": q_node.get(n["id"], [])})
    return out
