#!/usr/bin/env python3
"""terms.py — プログラムの中の名前と、業務の言葉の対応（用語の対応。structure/code_terms.json）。

アクティビティ図を読むのは、業務部門の担当。箱・判断・出口・図の名前には、プログラムの中だけの名前（変数・項目・手続き・モジュール）を
書かず、業務の言葉で書く。担当の目に見える名前（シート・ファイル・テーブル・画面・ボタンの文字・メッセージ）は、ソースの文字のまま「」で書く。
言い換えは、用語の対応に1回だけ書いて、どの図でも同じ言葉を使う（LLM が書く。形は references/schemas.md の「用語の対応」）。

  python3 scripts/activity.py terms -w <作業フォルダ>               # 名前の候補（箱が受け持つ行に出てくる、プログラムの中の名前）と、登録の状況
  python3 scripts/activity.py terms -w <作業フォルダ> --id A-02-1   # その図の分だけ

ここは、名前の見分け方を1か所にまとめた部品（activity.py・check_flow.py・build_matrix.py が使う）。
  プログラムの中の名前   ソースの、文字列・コメントの外にある名前のうち、機械が知っている名前（手続き・モジュール・用語の対応に載せた名前）か、
                         名前らしい形のもの（記号でつないである・小文字のあとに大文字がある・大文字だけで 4 文字以上）
  目に見える名前         箱に付いた入出力の名前（シート・ファイル・テーブル）と、ソースの文字列の中の文字
"""
import pathlib
import re

from common import load, load_page, norm, page_json, page_units, save

FILE = "structure/code_terms.json"
BASIS = {"comment": "コメント", "definition": "定義", "spelling": "つづり（推定）", "unknown": "不明"}
WORD = re.compile(r"(?:[0-9]+-)*[A-Za-z0-9_]*[A-Za-z][A-Za-z0-9_]*(?:-[A-Za-z0-9_]+)*")   # 名前1つ（COBOL の、- でつないだ形も1つに取る。9000-MAIN）
PART = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
DOUBLE = r'"((?:[^"\\\n]|\\.|"")*)"'
LITERAL = re.compile(DOUBLE + r"|'((?:[^'\\\n]|\\.)*)'")
PLAIN = r'"((?:[^"\n]|"")*)"'                                       # \ が、ただの文字の言語（VBA・COBOL・SQL）: 文字列の中の " は、2つ重ねる
LITERAL_VB = re.compile(PLAIN + r"|((?!))")                          # ' がコメントの始まりの言語: 文字列は " だけ
LITERAL_PLAIN = re.compile(PLAIN + r"|'((?:[^'\n]|'')*)'")
QUOTED = re.compile(r"「([^」]*)」")
APOSTROPHE = ("vba", "vbscript", "vbnet", "accessdef", "accessmacro")
COMMENT = {"vba": "'", "vbscript": "'", "vbnet": "'", "accessdef": "'", "java": "//", "csharp": "//", "javascript": "//", "typescript": "//",
           "cobol": "*>", "python": "#", "shell": "#", "powershell": "#", "sql": "--", "batch": "::"}
MIN = 3                                                              # これより短い名前は、見ない（ID・OK・No など）
CONSTANT = re.compile(r"\bConst\b|\bstatic\s+final\b|\bfinal\s+static\b|\bVALUE\b|#define\b|^\s*const\b", re.I)   # 定数の定義の行（値から言葉が読める）
# 言語の語（名前ではない）。大文字だけの語が、名前らしい形に当たる言語の分
RESERVED = {
    "cobol": set("""ACCEPT ADD AFTER ALL AND ASSIGN BEFORE CALL CLOSE COMPUTE CONTINUE CORRESPONDING DELETE DELIMITED DISPLAY DIVIDE ELSE END END-CALL END-COMPUTE END-EVALUATE
        END-IF END-PERFORM END-READ END-SEARCH END-STRING END-WRITE EQUAL ERROR EVALUATE EXIT EXTEND FALSE FROM FUNCTION GIVING GOBACK GREATER HIGH-VALUES INITIALIZE INPUT INSPECT
        INTO INVALID LENGTH LESS LOW-VALUES MOVE MULTIPLY NEXT NOT NUMERIC OPEN OTHER OUTPUT PERFORM PROGRAM READ REMAINDER REPLACING RETURN REWRITE ROUNDED SEARCH SECTION SENTENCE
        SIZE SORT SPACE SPACES START STOP STRING SUBTRACT TALLYING THAN THEN THROUGH THRU TIMES TRUE UNSTRING UNTIL UPON USING VARYING WHEN WITH WRITE ZERO ZEROES ZEROS""".split()),
    "sql": set("""SELECT FROM WHERE INSERT INTO VALUES UPDATE DELETE JOIN LEFT RIGHT INNER OUTER GROUP ORDER HAVING UNION DISTINCT CREATE TABLE ALTER DROP INDEX VIEW NULL LIKE BETWEEN
        EXISTS CASE WHEN THEN ELSE BEGIN COMMIT ROLLBACK DECLARE EXEC EXECUTE MERGE USING LIMIT COUNT""".split()),
}


def shaped(tok):
    """プログラムの中の名前らしい形か: 記号（_ -）でつないである／小文字のあとに大文字がある／大文字と数字だけで 4 文字以上"""
    return bool(re.search(r"[A-Za-z0-9][_\-][A-Za-z0-9]", tok) or re.search(r"[a-z][A-Z]", tok) or (len(tok) >= 4 and re.fullmatch(r"[A-Z][A-Z0-9]*", tok)))


def literal_re(lang):
    return LITERAL_VB if lang in APOSTROPHE else LITERAL_PLAIN if lang in ("cobol", "sql") else LITERAL


def literals(text, lang=""):
    """ソースの文字列の中の文字を、つないだもの（照合用に、空白などをならす）"""
    return norm(" ".join(a or b or "" for a, b in literal_re(lang).findall(text)))


def bare(text, lang=""):
    """ソースから、文字列と、行末のコメント・コメントだけの行を除いたもの（プログラムの中の名前だけが残る）"""
    out, mark, lit = [], COMMENT.get(lang), literal_re(lang)
    for line in text.split("\n"):
        if lang == "cobol" and line[6:7] in ("*", "/"):
            continue
        line = lit.sub(" ", line)
        out.append(line[:line.index(mark)] if mark and mark in line else line)
    return "\n".join(out)


def code_names(text, lang=""):
    """ソースの、文字列・コメントの外にある名前（- でつないだ形と、ばらした形の両方）"""
    b = bare(text, lang)
    return (set(WORD.findall(b)) | set(PART.findall(b))) - RESERVED.get(lang, set())


def plain(label):
    """「」の外の文字"""
    return QUOTED.sub(" ", str(label or ""))


def inner_names(label, code, known):
    """箱の名前に、そのまま出ている、プログラムの中の名前（「」の外。code は code_names の結果、known は機械が知っている名前）"""
    out = []
    for tok in WORD.findall(plain(label)):
        if len(tok) >= MIN and tok not in out and tok in code and (shaped(tok) or tok in known):
            out.append(tok)
    return out


def quoted_inner(label, text, lang, held, unknown):
    """「」の中に書いた、プログラムの中だけの名前: ソースの文字列・コメントの外にあって、文字列の中の文字でも、箱に付いた入出力の名前でもないもの
    （held は、ならした入出力の名前。unknown は、用語の対応で「不明」にした名前＝もとの名前のまま書いてよい）"""
    lit, code, out = literals(text, lang), norm(bare(text, lang)), []
    for q in QUOTED.findall(str(label or "")):
        n = norm(q)
        if n and n not in lit and n not in held and q not in unknown and q not in out and n in code:
            out.append(q)
    return out


class Terms:
    """用語の対応（structure/code_terms.json）と、まとまり（プロジェクト）のソース"""

    def __init__(self, work, structs):
        self.work, self.structs = pathlib.Path(work), structs
        raw = load(self.work / FILE, {}) or {}
        self.rows = [t for t in raw.get("terms") or [] if isinstance(t, dict)] if isinstance(raw, dict) else []
        self.by = {(t.get("project"), t.get("name")): t for t in self.rows}
        self._lines, self._all, self._defs = {}, {}, {}

    def exists(self):
        return (self.work / FILE).exists()

    def lines(self, proj, module):
        """モジュールの全部の行（[0] が 1 行目）"""
        key = (proj, module)
        if key not in self._lines:
            out, st = [], self.structs.get(proj) or {}
            m = next((x for x in st.get("modules") or [] if x["name"] == module), None)
            for pg in (m or {}).get("pages") or []:
                units = page_units(load_page(page_json(self.work, pg.get("doc_id") or m["doc_id"], pg["page"])))
                out += [""] * (pg["line_from"] - 1 - len(out)) + [str(u) for u in units[:pg["lines"]]]
            self._lines[key] = out
        return self._lines[key]

    def text(self, proj, module, nums):
        ls = self.lines(proj, module)
        return "\n".join(ls[n - 1] for n in sorted(set(nums)) if 0 < n <= len(ls))

    def whole(self, proj):
        if proj not in self._all:
            self._all[proj] = "\n".join("\n".join(self.lines(proj, m["name"])) for m in (self.structs.get(proj) or {}).get("modules") or [])
        return self._all[proj]

    def lang(self, proj):
        return (self.structs.get(proj) or {}).get("language", "")

    def known(self, proj):
        """機械が知っている、プログラムの中の名前: 手続き・モジュールの名前と、用語の対応に載せた名前（英数字を含むものだけ）"""
        st = self.structs.get(proj) or {}
        names = {p["name"] for p in st.get("procedures") or []} | {m["name"] for m in st.get("modules") or []}
        names |= {re.sub(r"\.[A-Za-z0-9]+$", "", n) for n in names} | {n for (p, n), t in self.by.items() if p == proj and t.get("basis") != "unknown"}
        return {n for n in names if n and re.search(r"[A-Za-z]", n) and len(n) >= MIN}

    def unknown(self, proj):
        return {n for (p, n), t in self.by.items() if p == proj and t.get("basis") == "unknown"}

    def evidence(self, proj, name):
        """名前の定義の行から、言葉の根拠を機械が読む（code-guide の用語の対応）: (根拠, モジュール, 行) か None。
        定義の行の行末か、すぐ上の行にコメントがある → comment。定数の定義（Const・static final・VALUE・#define）→ definition。場所は、定義の行"""
        at = self.defined(proj, name) or self.declared(proj, name)
        if not at:
            return None
        mod, n, _text = at
        lang, ls = self.lang(proj), self.lines(proj, mod)
        mark, raw = COMMENT.get(lang), ls[n - 1]
        code = literal_re(lang).sub(" ", raw)
        above = ls[n - 2] if n >= 2 else ""
        if mark and mark in code and code.split(mark, 1)[1].strip():
            return "comment", mod, n
        if (mark and above.strip().startswith(mark) and above.strip()[len(mark):].strip()) or (lang == "cobol" and above[6:7] == "*" and above[7:].strip()):
            return "comment", mod, n
        if CONSTANT.search(code):
            return "definition", mod, n
        return None

    def fill(self):
        """根拠（basis）の無い行と、コメント・定義が根拠なのに場所（source）の無い行に、機械が読んだ根拠と場所を入れて保存する。入れた所の説明を返す"""
        done = []
        for t in self.rows:
            ev = self.evidence(t.get("project"), str(t.get("name") or ""))
            if not t.get("basis") and ev:
                t["basis"] = ev[0]
                done.append(f"「{t.get('name')}」の根拠を「{BASIS[ev[0]]}」にした（{ev[1]} {ev[2]} 行）")
            if t.get("basis") in ("comment", "definition") and not isinstance(t.get("source"), dict) and ev and ev[0] == t["basis"]:
                t["source"] = {"module": ev[1], "line": ev[2]}
                done.append(f"「{t.get('name')}」の根拠の場所を入れた（{ev[1]} {ev[2]} 行）")
        if done:
            raw = load(self.work / FILE, {}) or {}
            raw["terms"] = self.rows
            save(self.work / FILE, raw)
        return done

    def warnings(self):
        """誤りではないが、見直す所: 定義の行にコメントがあるのに、根拠が「つづり」"""
        out = []
        for k, t in enumerate(self.rows):
            ev = self.evidence(t.get("project"), str(t.get("name") or "")) if t.get("basis") == "spelling" else None
            if ev and ev[0] == "comment":
                out.append(f"terms[{k}]「{t.get('name')}」: 定義の行にコメントがある（{ev[1]} {ev[2]} 行）のに、根拠が「つづり」。コメントを読んで言葉を決める"
                           "（コメントから読んだなら basis を消す: 機械が「コメント」と場所を入れる）")
        return out

    def problems(self):
        """用語の対応そのものの誤り"""
        out, seen = [], set()
        for k, t in enumerate(self.rows):
            at, proj, name = f"terms[{k}]", t.get("project"), str(t.get("name") or "")
            if proj not in self.structs:
                out.append(f"{at}: project「{proj}」が、プログラムのまとまりに無い")
                continue
            if (proj, name) in seen:
                out.append(f"{at}: 「{name}」が、2回載っている（1つの名前に、言葉は1つ）")
            seen.add((proj, name))
            if not t.get("basis"):
                out.append(f"{at}「{name}」: 根拠（basis）が無く、定義の行にコメント・定数の定義も無い（つづりから読んだなら spelling、読めないなら unknown を書く）")
            if t.get("basis") != "unknown" and not str(t.get("word") or "").strip():
                out.append(f"{at}「{name}」: 業務の言葉（word）が空（意味が読めないなら、basis を unknown にする。一覧に「不明」として出る）")
            if name and name not in self.whole(proj):
                out.append(f"{at}: 名前「{name}」が、{proj} のソースに無い（もとの名前は、ソースの文字のまま書く）")
            src = t.get("source") if isinstance(t.get("source"), dict) else {}
            if t.get("basis") in ("comment", "definition") and not (src.get("module") and src.get("line")):
                out.append(f"{at}「{name}」: 根拠が「{BASIS[t['basis']]}」なのに、場所（source の module・line）が無い（コメント・定義のある行を書く）")
            if src.get("line") and name not in "\n".join(self.lines(proj, str(src.get("module")))[src["line"] - 1:src["line"]]):
                out.append(f"{at}「{name}」: 定義の場所（{src.get('module')} {src.get('line')} 行）に、その名前が無い")
        return out

    def declared(self, proj, name):
        """名前（日本語の名前も）が、実行文でない行に、文字列・コメントの外で出てくる、はじめの所: (モジュール, 行, 文字)。無ければ None"""
        st, lang = self.structs.get(proj) or {}, self.lang(proj)
        ran = {(p["module"], n) for p in st.get("procedures") or [] for n in p.get("exec") or []}
        pat = re.compile(r"(?<![\w\-])" + re.escape(name) + r"(?![\w\-])")
        for m in st.get("modules") or []:
            for n, line in enumerate(self.lines(proj, m["name"]), 1):
                if (m["name"], n) not in ran and pat.search(bare(line, lang)):
                    return m["name"], n, line.strip()
        return None

    def defined(self, proj, name):
        """名前が、実行文でない行（宣言・定義の行）に出てくる、はじめの所: (モジュール, 行, 文字)。無ければ None（部品・枠組みの名前か、読んでいない定義）"""
        if proj not in self._defs:
            st, lang, found = self.structs.get(proj) or {}, self.lang(proj), {}
            ran = {(p["module"], n) for p in st.get("procedures") or [] for n in p.get("exec") or []}
            for m in st.get("modules") or []:
                for n, line in enumerate(self.lines(proj, m["name"]), 1):
                    if (m["name"], n) not in ran and line.strip():
                        for tok in code_names(line, lang):
                            found.setdefault(tok, (m["name"], n, line.strip()))
            self._defs[proj] = found
        return self._defs[proj].get(name)



def candidates(terms, draft):
    """図にする単位1つの、名前の候補: [(名前, 出てくる行の数, 登録した行か None, 定義の行か None)]。箱になる項目の、実行文の行から拾う。
    このソースに定義のある名前が先（定義の行つき）、定義の無い名前（部品・枠組みの名前か、読んでいない定義）があと"""
    proj = draft.unit["project"]
    known, lang, count = terms.known(proj), terms.lang(proj), {}
    for it in draft.walk():
        if it["state"] == "excluded" or not it.get("exec"):
            continue
        for n in it["exec"]:
            for tok in code_names(terms.text(proj, it["module"], [n]), lang):
                if len(tok) >= MIN and (shaped(tok) or tok in known):
                    count[tok] = count.get(tok, 0) + 1
    whole = {t for t in count if "-" in t}                           # - でつないだ名前の部分（ACCT と ID）は、つないだ形（ACCT-ID）があれば出さない
    parts = {p for t in whole for p in t.split("-")}
    rows = [(tok, count[tok], terms.by.get((proj, tok)), terms.defined(proj, tok)) for tok in count if tok in whole or tok not in parts]
    return sorted(rows, key=lambda r: (r[3] is None and r[2] is None, -r[1], r[0]))
