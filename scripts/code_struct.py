#!/usr/bin/env python3
"""code_struct.py — プログラムのソースから、構造を機械で取り出す（コードは実行しない。文字として読むだけ）。

  python3 scripts/code_struct.py -w <作業フォルダ>              # コードのある資料ごとに code/<資料ID>.struct.json を書く
  python3 scripts/code_struct.py -w <作業フォルダ> --show M05    # 取り出した結果（手続き・起動点・木）を表示する
  python3 scripts/code_struct.py --file a.bas b.cls              # ファイルを直接解析して表示する（作業フォルダなし）

取り出すもの（形は references/schemas.md の「コードの構造」）:
  modules      モジュール（ファイル）と、そのページ・行
  procedures   手続きごとに、ブロックの木（続く文のまとまり seq／分岐 if・select／繰り返し loop／呼び出し call／操作者とのやり取り interact／
               抜ける exit／エラーの設定 onerror）、呼び出し、入出力（シート・ファイル・テーブル・メール・外部プログラム）、実行文の行
  entries      起動点（コードの外から始まる所）: ボタン・イベント・自動実行・定時実行・スクリプトの本体
  units        アクティビティ図にする単位: 起動点ごとに1つ。呼ばれる手続きは、呼ぶ所が1か所だけで箱が 8 個以下なら呼ぶ側に入れ、ほかは別の単位
  unreached    どの起動点からもたどれない手続き
構造の段（tier）: parsed 機械が構文から／heuristic 機械が行の形から／llm LLM が木を書く（機械は、範囲と入れ子を確かめる）。
行の番号は、モジュールの先頭からの番号（1 から）。ページの中の番号への読み替えは cite() が行う。
"""
import argparse
import pathlib
import re
import sys

import extract_access
from common import MASK, load, manifest, mask_secrets, page_json, load_page, save, work_dir  # noqa: F401（MASK は、ほかから cs.MASK で引く）

INLINE_MAX = 8                                                       # 呼ぶ所が1か所で、箱がこの数以下の手続きは、呼ぶ側の図に入れる
CODE_KINDS = ("vba", "code")                                         # ページの種類（manifest の page_list[].kind）
LANG_OF_EXT = {".bas": "vba", ".cls": "vba", ".frm": "vba", ".vbs": "vbscript", ".wsf": "vbscript", ".vb": "vbnet", ".py": "python", ".sql": "sql",
               ".bat": "batch", ".cmd": "batch", ".ps1": "powershell", ".psm1": "powershell", ".sh": "shell", ".js": "javascript", ".gs": "javascript",
               ".ts": "typescript", ".java": "java", ".cs": "csharp", ".cbl": "cobol", ".cob": "cobol", ".cpy": "cobol", ".php": "php", ".rb": "ruby",
               ".pl": "perl", ".kt": "kotlin", ".go": "go", ".c": "c", ".cpp": "cpp", ".h": "c", ".jcl": "jcl"}
LANG_LABEL = {"vba": "VBA", "vbscript": "VBScript", "python": "Python", "sql": "SQL", "batch": "バッチ", "powershell": "PowerShell", "shell": "シェル",
              "javascript": "JavaScript", "typescript": "TypeScript", "java": "Java", "csharp": "C#", "cobol": "COBOL", "vbnet": "VB.NET",
              "accessmacro": "Access マクロ", "accessdef": "Access の定義"}
VBA_FAMILY = ("vba", "vbscript", "accessmacro", "accessdef")         # 1つのまとまりとして読む言語（Access を書き出したフォルダは、全部で1つ）
TIER_LABEL = {"parsed": "機械（構文から）", "heuristic": "機械（行の形から）", "llm": "LLM の読取"}
TRIGGER_LABEL = {"button": "ボタン", "event": "イベント", "auto": "自動実行", "timer": "定時実行", "shortcut": "ショートカット", "ribbon": "リボン",
                 "manual": "マクロの一覧から実行", "form": "フォームの操作", "script": "スクリプトの本体", "query": "クエリ", "entry": "入口"}

# ---------- VBA・VBScript ----------

IDENT = r"[A-Za-z_\u0080-￿][\w\u0080-￿]*"
PROC_START = re.compile(rf"(?i)^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+({IDENT})\s*(\(.*)?$")
PROC_END = re.compile(r"(?i)^End\s+(Sub|Function|Property)$")
DECL = re.compile(r"(?i)^(Dim|Const|Static|Private|Public|Global|Friend|Declare|Option|Implements|Event|Attribute|Def\w+|#\w+)\b")
BLOCK_SKIP = re.compile(r"(?i)^(?:(?:Public|Private)\s+)?(Type|Enum)\s+\w+$")
VBA_EVENT = re.compile(r"(?i)^(Auto_?Open|Auto_?Close|Workbook_\w+|Worksheet_\w+|Document_\w+|UserForm_\w+|Form_\w+|Report_\w+|Class_\w+|\w+_(Click|DblClick|Change|"
                       r"AfterUpdate|BeforeUpdate|Activate|Deactivate|Open|Close|Load|Unload|Current|Initialize|Terminate|Calculate|SelectionChange|"
                       r"BeforeSave|BeforeClose|BeforeDoubleClick|BeforeRightClick|Enter|Exit|GotFocus|LostFocus|KeyDown|KeyPress|KeyUp|NotInList|Timer|"
                       r"MouseDown|MouseUp|MouseMove|NoData|Print|Format|Resize|Error|Dirty|Delete|BeforeInsert|AfterInsert|BeforeDelConfirm|AfterDelConfirm))$")
EVENT_JA = {"workbook_open": "ブックを開いたとき", "auto_open": "ブックを開いたとき", "workbook_beforeclose": "ブックを閉じるとき", "auto_close": "ブックを閉じるとき",
            "workbook_beforesave": "ブックを保存するとき", "workbook_sheetchange": "どれかのシートの値が変わったとき", "worksheet_change": "値が変わったとき",
            "worksheet_selectionchange": "選ぶセルが変わったとき", "worksheet_activate": "シートを表示したとき", "worksheet_beforedoubleclick": "セルをダブルクリックしたとき",
            "worksheet_calculate": "再計算されたとき", "userform_initialize": "フォームを開いたとき", "form_load": "フォームを開いたとき", "form_open": "フォームを開いたとき",
            "form_current": "表示するレコードが変わったとき", "form_beforeupdate": "レコードを保存する前", "form_afterupdate": "レコードを保存したあと",
            "form_close": "フォームを閉じるとき", "report_open": "レポートを開いたとき",
            "worksheet_deactivate": "ほかのシートに移ったとき", "chart_activate": "グラフのシートを表示したとき", "chart_deactivate": "ほかのシートに移ったとき",
            "workbook_activate": "ブックの画面に切り替えたとき", "workbook_deactivate": "ほかのブックに切り替えたとき",
            "workbook_windowactivate": "ブックの画面に切り替えたとき", "workbook_windowdeactivate": "ほかの画面に切り替えたとき",
            "workbook_sheetactivate": "どれかのシートを表示したとき", "workbook_newsheet": "シートを足したとき", "workbook_beforeprint": "印刷するとき",
            "workbook_aftersave": "ブックを保存したあと", "userform_activate": "フォームを表示したとき", "userform_queryclose": "フォームを閉じるとき",
            "userform_terminate": "フォームを閉じたとき"}
CTL_EVENT_JA = {"click": "をクリック", "dblclick": "をダブルクリック", "change": "の値を変えたとき", "afterupdate": "の値を変えたあと", "beforeupdate": "の値を変える前",
                "enter": "に入ったとき", "exit": "から出たとき", "notinlist": "に一覧に無い値を入れたとき", "gotfocus": "に入ったとき", "lostfocus": "から出たとき",
                "keypress": "でキーを押したとき", "keydown": "でキーを押したとき", "keyup": "でキーを離したとき"}
# 画面の部品（フォームのリスト・テキスト・ラベル）に書く文。機械は入出力と見ないが、画面に見える動きなので、補助の候補（図にしない手続き）から外す（決まり 6.32）
UI_WRITE = re.compile(r"(?i)\.(AddItem|RemoveItem|RowSource|ListIndex|Caption|ControlSource|Requery|SetFocus)\b|^Me\.\w+\.(Value|Text|List|Visible|Enabled|Clear)\b|^Me\.Repaint\b")
INTERACT = [("msgbox", re.compile(r"(?i)\bMsgBox\b")), ("inputbox", re.compile(r"(?i)\bInputBox\b")),
            ("filedialog", re.compile(r"(?i)\b(GetOpenFilename|GetSaveAsFilename|FileDialog)\b")),
            ("form", re.compile(r"(?i)\bDoCmd\.OpenForm\b"))]
SHEET = re.compile(r'(?i)\b(?:Work)?sheets\s*\(\s*("([^"]*)"|[^()]*(?:\([^()]*\))?[^()]*)\s*\)')
WRITE_METHOD = re.compile(r"(?i)\.(Clear\w*|Delete|Insert|PasteSpecial|Paste|Sort|AutoFilter|RemoveDuplicates|AutoFill|FillDown|Merge|UnMerge|Protect|Unprotect|"
                          r"Replace|TextToColumns|AddNew|Update|Value2?\s*=|Formula\w*\s*=|NumberFormat\w*\s*=|Name\s*=|Visible\s*=)")
TECH = [("画面の更新の制御", re.compile(r"(?i)^Application\.(ScreenUpdating|DisplayAlerts|EnableEvents|Calculation|StatusBar|Cursor|CutCopyMode|Interactive)\s*=")),
        ("画面の更新の制御", re.compile(r"(?i)^(DoEvents|DoCmd\.SetWarnings\b.*|DoCmd\.Hourglass\b.*|DoCmd\.Echo\b.*)$")),
        ("ログ", re.compile(r"(?i)^Debug\.(Print|Assert)\b")),
        ("後片付け", re.compile(r"(?i)^(Set\s+\S+\s*=\s*Nothing|Erase\s+.+|[\w.\u0080-\uffff]+\.(Close|Quit)\b.*|Close(\s+#?\w+)?)$"))]
FORMAT_PROP = re.compile(r"(?i)\.(Font|Interior|Borders?|NumberFormat\w*|HorizontalAlignment|VerticalAlignment|ColumnWidth|RowHeight|WrapText|Orientation|IndentLevel|"
                         r"ShrinkToFit|FormatConditions|Characters)\b")       # 書式だけの項目（字・塗り・罫線・表示の形・列の幅・行の高さ）
FORMAT_CALL = re.compile(r"(?i)\.(AutoFit|BorderAround|ClearFormats)\b")
FETCH_ONLY = re.compile(rf"(?i)^(?:Let\s+)?({IDENT})\s*=\s*(?:[^=]*\.End\s*\(\s*xl(?:Up|Down|ToLeft|ToRight)\s*\)\s*\.\s*(?:Row|Column)|[^=]*\.(?:Rows|Columns|ListRows|Cells)\s*\.\s*Count"
                        rf"|(?:Len|UBound|LBound)\s*\([^=]*\)|(?:Date|Now|Time)(?:\s*\(\s*\))?)$")   # 最後の行・件数・長さ・今日の日付を取るだけ
STEP_ON = re.compile(rf"(?i)^(?:Let\s+)?({IDENT})\s*=\s*({IDENT})\s*\+\s*1$")    # 1 つ進める（位置か、件数か）
MOVE_ON = re.compile(rf"(?i)^{IDENT}\s*\.\s*Move(?:Next|Previous|First|Last)\s*$")  # レコードの位置を進める
INDEX_USE = re.compile(rf"(?i)\b(?:Cells|Offset)\s*\(\s*({IDENT})\s*[,)]|\b(?:Cells|Offset)\s*\(\s*[^,()]+,\s*({IDENT})\s*\)|\bRange\s*\(\s*\"[A-Za-z]+\"\s*&\s*({IDENT})\b")
COPY_ONLY = re.compile(rf"(?i)^(?:Set\s+|Let\s+)?({IDENT})\s*=\s*(?:({IDENT})|[+-]?\d+(?:\.\d+)?[#&!%@]?|\"\s*\")$")
LITERALS = {"true", "false", "nothing", "empty", "null"}
DECL_VARS = re.compile(r"(?i)^(?:(?:Dim|Private|Public|Global|Static)\s+(?:Const\s+|WithEvents\s+)?|Const\s+)(?!(?:Sub|Function|Property|Declare|Enum|Type|Event|Static)\b)(.+)$")
ASSIGNED = re.compile(rf"(?i)^(?:(?:Set\s+|Let\s+)?({IDENT})\s*=|For\s+(?:Each\s+)?({IDENT})\b)")
NAVIGATE = re.compile(r"(?i)^[\w.()\"\u0080-￿ ,]*\.(Activate|Select)\s*$")
DYNAMIC = "（実行時に決まる）"
ACTIVE = "（表示中のシート）"


def declared(text):
    """宣言の文（Dim a, b(1 To 3) As Long・Private Const X = 1）→ 宣言している名前（小文字）"""
    mt = DECL_VARS.match(blank_strings(text))
    depth, piece, out = 0, "", []
    for ch in (mt.group(1) + ",") if mt else "":
        if ch == "," and depth == 0:
            m2 = re.match(rf"\s*({IDENT})", piece)
            out += [m2.group(1).lower()] if m2 else []
            piece = ""
        else:
            depth += (ch == "(") - (ch == ")")
            piece += ch
    return out


def declared_in(p):
    """手続きの中で宣言した名前（モジュールの変数と同じ名前を、手続きの中で宣言し直しているとき）"""
    return {x for text, _a, _b in p["mod"].stmts[p["s0"]:p["s1"]] for x in declared(text)}


def strip_comment(line):
    """行のコメントを外した文字と、コメントの文字。文字列の中の ' は数えない"""
    quoted = False
    for i, ch in enumerate(line):
        if ch == '"':
            quoted = not quoted
        elif ch == "'" and not quoted:
            return line[:i].rstrip(), line[i + 1:].strip()
    mt = re.match(r"(?i)^\s*Rem(\s+(.*))?$", line)
    return ("", (mt.group(2) or "").strip()) if mt else (line.rstrip(), "")


def format_only(text):
    """書式だけを変える文か（字・塗り・罫線・表示の形・列の幅・行の高さ）。データの読み書きではない（決まり 6.10。機械が「書式・表示の調整」で対象外にする）"""
    parts = split_top(text, "=")
    if len(parts) > 1 and not re.match(r"(?i)^(If|ElseIf|Do|Loop|While|Case|Select|For|Call|Set)\b", text):
        return bool(FORMAT_PROP.search(blank_strings(parts[0])))
    return bool(FORMAT_CALL.search(blank_strings(text)))


def blank_strings(text):
    """文字列の中身を空白にする（語を探すとき、文字列の中の語に当てない）。長さは変えない"""
    out, quoted = [], False
    for ch in text:
        if ch == '"':
            quoted = not quoted
            out.append(ch)
        else:
            out.append(" " if quoted else ch)
    return "".join(out)


DATE_LITERAL = re.compile(r"#\s*(?:\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?|\d{1,2}:\d{2}(?::\d{2})?|[A-Za-z]{3,9}\.? \d{1,2},? \d{4})(?:\s*[AaPp][Mm])?\s*#")


def split_top(text, sep):
    """文字列・かっこ・日付の定数（#…#）の外にある sep で切る"""
    out, depth, quoted, cur, i = [], 0, False, "", 0
    while i < len(text):
        ch = text[i]
        if ch == "#" and not quoted:
            mt = DATE_LITERAL.match(text, i)
            if mt:
                cur, i = cur + mt.group(0), mt.end()
                continue
        if ch == '"':
            quoted = not quoted
        elif not quoted and ch in "([":
            depth += 1
        elif not quoted and ch in ")]":
            depth -= 1
        if not quoted and depth == 0 and text.startswith(sep, i) and not (sep == ":" and text[i + 1:i + 2] == "="):
            out.append(cur)
            cur, i = "", i + len(sep)
            continue
        cur += ch
        i += 1
    out.append(cur)
    return out


def literal(expr):
    """式が文字列だけなら、その中身。そうでなければ None（実行時に決まる）"""
    mt = re.fullmatch(r'\s*"((?:[^"]|"")*)"\s*', expr or "")
    return mt.group(1).replace('""', '"') if mt else None


def template(expr, consts=None):
    """文字と変数をつないだ式（"売上_" & 月 & ".csv"）→ 書かれた文字はそのまま、変数は ＜名前＞ にした名前。文字が1つも無ければ None"""
    parts, got = [], False
    for x in split_top(expr.strip().strip("()"), "&"):
        x = x.strip()
        lit = literal(x)
        if lit is None and consts and x.lower() in consts:
            lit = consts[x.lower()]
        if lit is not None:
            parts.append(lit)
            got = got or bool(lit.strip("\\/ "))
        else:
            parts.append("＜" + re.sub(r"\(.*$", "", x)[:20] + "＞")
    return "".join(parts) if got else None


def file_name(expr, consts=None):
    """ファイルを指す式 → (名前, 実行時に決まるか)。名前は、フォルダを除いた末尾"""
    nm = literal(expr)
    if nm is None and consts and expr.strip().lower() in consts:
        nm = consts[expr.strip().lower()]
    if nm is not None:
        return pathlib.PureWindowsPath(nm).name or nm, False
    t = template(expr, consts)
    if t:
        tail = re.split(r"[\\/]", t)[-1]
        return (tail or t), not tail.strip("＜＞") or tail.startswith("＜") and tail.endswith("＞") and tail.count("＜") == 1
    return DYNAMIC, True


def first_string(text):
    mt = re.search(r'"((?:[^"]|"")+)"', text)
    return mt.group(1).replace('""', '"') if mt else ""


class VbaModule:
    """モジュール1つを、文（論理行）に切って、手続きごとの木にする"""

    def __init__(self, name, kind, lines, sheet=None, start=1):
        self.name, self.kind, self.lines, self.sheet = name, kind, lines, sheet
        self.start = start                                           # コードの最初の行（Access のフォーム・レポートは、定義のあと）
        self.stmts, self.comment_at = [], {}                         # 文: (文字, 開始の行, 終了の行)／コメントだけの行 → 文字
        self.cut()
        self.consts = {}                                             # 文字の定数: 名前（小文字）→ 中身
        for text, _a, _b in self.stmts:
            mt = re.match(rf'(?i)^(?:(?:Public|Private|Global)\s+)?Const\s+({IDENT})(?:\s+As\s+\w+)?\s*=\s*("(?:[^"]|"")*")\s*$', text)
            if mt:
                self.consts[mt.group(1).lower()] = literal(mt.group(2))
        self.heads = []                                              # 手続き: {name, kind, lines, s0, s1}（s0・s1 は stmts の中の位置）
        i = 0
        while i < len(self.stmts):
            text, a, _b = self.stmts[i]
            if BLOCK_SKIP.match(text):                               # Type・Enum の定義は読み飛ばす
                while i < len(self.stmts) and not re.match(r"(?i)^End\s+(Type|Enum)$", self.stmts[i][0]):
                    i += 1
            mt = PROC_START.match(text)
            if mt:
                j = next((k for k in range(i + 1, len(self.stmts)) if PROC_END.match(self.stmts[k][0])), len(self.stmts) - 1)
                self.heads.append({"name": mt.group(2), "kind": re.sub(r"\s+", " ", mt.group(1)).lower(), "public": not re.match(r"(?i)^(Private|Friend)\b", text),
                                   "params": bool(re.sub(r"[()\s]", "", (mt.group(3) or "").split(")")[0])), "lines": [a, self.stmts[j][2]], "s0": i + 1, "s1": j,
                                   "args": [x.lower() for x in re.findall(rf"(?i)(?:ByVal|ByRef|Optional|ParamArray|\(|,)\s*(?:ByVal\s+|ByRef\s+)?({IDENT})(?=\s*(?:\(\s*\))?\s*(?:As\b|,|\)|=))", mt.group(3) or "")]})
                i = j
            i += 1

    def cut(self):
        buf, start = "", None
        for k, raw in enumerate(self.lines, 1):
            if k < self.start:
                continue
            code, note = strip_comment(raw)
            if not code.strip():
                if note and start is None:
                    self.comment_at[k] = note
                if start is None:
                    continue
            if start is None:
                start = k
            if re.search(r"(^|\s)_$", code):                         # 行の継続
                buf += code[:-1] + " "
                continue
            text = re.sub(r"\s+", " ", (buf + code).strip())
            text = re.sub(r"^\d+\s+", "", text)                      # 行番号
            buf, a, start = "", start, None
            if not text:
                continue
            lab = re.match(rf"^({IDENT}):(?!=)\s*(.*)$", text)
            if lab and lab.group(1).lower() not in ("else", "do", "loop", "next", "wend", "then"):
                self.stmts.append((f"@label {lab.group(1)}", a, k))
                text = lab.group(2)
                if not text:
                    continue
            if re.match(r"(?i)^If\b.*\bThen\s*\S", blank_strings(text)) and not re.match(r"(?i)^If\b.*\bThen\s*$", blank_strings(text)):
                self.stmts.append((text, a, k))                      # 1行の If は、切らずに1文にする
                continue
            for part in split_top(text, ":"):
                if part.strip():
                    self.stmts.append((part.strip(), a, k))


class VbaParser:
    """プロジェクト（モジュールの集まり）を解析する"""

    def __init__(self, modules, sheets=None, buttons=None, access=None):
        self.mods = modules
        self.access = access                                         # Access の定義（マクロ・クエリの名前を引く。無ければ None）
        self.sheets = sheets or {}                                   # シートのコード名 → シート名
        self.buttons = buttons or []
        self.procs = {}                                              # "モジュール.手続き" → 手続き
        self.by_name = {}                                            # 手続きの名前（小文字）→ ["モジュール.手続き"]
        self.forms = {m.name.lower() for m in modules if m.kind == "form"}
        self.consts = {k: v for m in modules for k, v in m.consts.items()}
        for m in modules:
            for h in m.heads:
                pid = f"{m.name}.{h['name']}"
                if pid in self.procs:                                # 同じ名前の手続き（Property の Get・Let・Set）は、種類を添えて別の番号にする
                    pid, k = f"{pid}（{h['kind'].split()[-1].capitalize()}）", 2
                    while pid in self.procs:
                        pid, k = f"{m.name}.{h['name']}（{k}）", k + 1
                self.procs[pid] = dict(h, id=pid, module=m.name, mod=m)
                self.by_name.setdefault(h["name"].lower(), []).append(pid)
        # オブジェクトの変数越しの呼び出し（Dim x As clsFoo … x.Run）を追うための、変数の型: モジュールの宣言と、手続きの中の宣言・引数
        self.cur, self.mtypes, self.ltypes, self.rtypes = None, {}, {}, {}
        known = {m.name.lower(): m.name for m in modules if m.kind != "standard"}

        def typed(text, into):
            for mt in re.finditer(rf"(?i)\b({IDENT})(?:\(\))?\s+As\s+(?:New\s+)?({IDENT})\b", text):
                if mt.group(2).lower() in known:
                    into[mt.group(1).lower()] = known[mt.group(2).lower()]
            mt = re.match(rf"(?i)^Set\s+({IDENT})\s*=\s*New\s+({IDENT})\s*$", text)
            if mt and mt.group(2).lower() in known:
                into[mt.group(1).lower()] = known[mt.group(2).lower()]
        if known:
            for m in modules:
                inside = set()
                for h in m.heads:
                    inside.update(range(h["s0"] - 1, h["s1"] + 1))
                for i, (text, _a, _b) in enumerate(m.stmts):
                    if i not in inside:
                        typed(text, self.mtypes.setdefault(m.name, {}))
            for pid, pr in self.procs.items():
                for text, _a, _b in pr["mod"].stmts[pr["s0"] - 1:pr["s1"]]:
                    typed(text, self.ltypes.setdefault(pid, {}))
                head = pr["mod"].stmts[pr["s0"] - 1][0]              # 関数・プロパティの戻り値の型（Options.Load のように、戻り値のオブジェクトの手続きを呼ぶ）
                mt = re.search(rf"(?i)\)\s+As\s+({IDENT})\s*$", head)
                if mt and mt.group(1).lower() in known and pr["kind"] != "sub":
                    self.rtypes.setdefault(pr["name"].lower(), set()).add(known[mt.group(1).lower()])
        # 「変数の準備」を見分けるための、モジュールの変数（手続きの外の宣言。値が手続きの外に残る）と、手続きの中で宣言した・値を入れた変数
        self.mvars, self.gvars = {}, set()
        for m in modules:
            inside = set()
            for h in m.heads:
                inside.update(range(h["s0"] - 1, h["s1"] + 1))
            for i, (text, _a, _b) in enumerate(m.stmts):
                if i not in inside:
                    names = declared(text)
                    self.mvars.setdefault(m.name, set()).update(names)
                    if re.match(r"(?i)^(Public|Global)\b", text):
                        self.gvars.update(names)
        for pr in self.procs.values():
            pr["vars"] = set(pr.get("args") or [])
            index, shown = set(), set()                              # セルの行・列の番号に使う変数／値として見せる・書く・ほかに写す変数
            for text, _a, _b in pr["mod"].stmts[pr["s0"]:pr["s1"]]:
                pr["vars"].update(declared(text))
                mt = ASSIGNED.match(blank_strings(text))
                if mt:
                    pr["vars"].add((mt.group(1) or mt.group(2)).lower())
                index.update(x.lower() for m2 in INDEX_USE.finditer(text) for x in m2.groups() if x)
                rest = INDEX_USE.sub(" ", blank_strings(text))
                parts = split_top(rest, "=")
                own = re.match(rf"(?i)^(?:Let\s+)?({IDENT})\s*$", parts[0].strip()) if len(parts) > 1 else None
                if re.search(r"(?i)\bMsgBox\b", rest) or (len(parts) > 1 and (re.search(r"(?i)\.(Value2?|Formula\w*|Caption|Text)\s*$", parts[0]) or own)):
                    seen = {w.lower() for w in re.findall(IDENT, "=".join(parts[1:]) if len(parts) > 1 else rest)}
                    shown.update(seen - ({own.group(1).lower()} if own and STEP_ON.match(blank_strings(text).strip()) else set()))
            pr["index"] = index - shown                              # 位置だけに使う変数（1 つ進める文は、位置を進めるだけの文）
        for p in self.procs.values():
            self.parse_proc(p)

    def working(self, left, right):
        """手続きの中だけの変数に、定数か、ほかの変数の値を、そのまま入れるだけの文か（left = right。right が None なら定数）"""
        p, low = self.cur, left.lower()
        outer = self.mvars.get(p["module"], set()) | self.gvars
        if low == p["name"].lower() or low in (p.get("args") or []):  # 戻り値と、引数（呼んだ側に返す値）は、残す
            return False
        if low in self.by_name or (low in outer and low not in declared_in(p)):   # プロパティに入れる文と、モジュールの変数（手続きの外に残る値）も、残す
            return False
        if right is None or right.lower() in LITERALS or re.match(r"vb[A-Z]", right):
            return True
        r = right.lower()                                            # 右が、引数なしの関数（Now・自作の関数）かもしれないときは、残す
        return r not in self.by_name and (r in p["vars"] or r in outer)

    def fetch_only(self, text):
        """最後の行・件数・長さ・今日の日付を取るだけの文／繰り返しの位置を進めるだけの文（決まり 6.10 の「変数の準備」）。
        入れる先は、手続きの中だけの変数。1 つ進める文は、その変数をセルの行・列の番号にだけ使うとき（見せる・書く・ほかに写す変数は、業務の件数かもしれない）"""
        bare = blank_strings(text).strip()
        if MOVE_ON.match(bare):
            return True
        p = self.cur
        step = STEP_ON.match(bare)
        mt = step or FETCH_ONLY.match(bare)
        if not mt:
            return False
        low = mt.group(1).lower()
        clock = re.search(r"(?i)=\s*(Date|Now|Time)\b", bare)
        if clock and clock.group(1).lower() in self.by_name:          # 同じ名前の自作の手続き
            return False
        outer = self.mvars.get(p["module"], set()) | self.gvars
        if low == p["name"].lower() or low in (p.get("args") or []) or low in self.by_name or (low in outer and low not in declared_in(p)):
            return False
        return low in p.get("index", set()) and step.group(2).lower() == low if step else True

    # ---- 名前の引き当て ----

    def resolve(self, name, module):
        """手続きの名前 → "モジュール.手続き"。同じモジュールを先に、次に公開の手続き。無ければ None"""
        if "." in name:
            mod, _, nm = name.rpartition(".")
            if mod.lower() == "me":                                  # 自分のモジュールの手続き
                mod = module
            tp = dict(self.mtypes.get(module, {}))                   # クラスの変数（Dim x As clsFoo）越しの呼び出し
            if self.cur is not None and self.cur.get("module") == module:
                tp.update(self.ltypes.get(self.cur["id"], {}))
            mod = tp.get(mod.lower(), mod)
            back = self.rtypes.get(mod.lower()) or ()                # 変数でなければ、関数・プロパティの戻り値の型（1つに決まるときだけ）
            if len(back) == 1 and mod.lower() not in tp:
                mod = next(iter(back))
            hit = [x for x in self.by_name.get(nm.lower(), []) if x.lower() == f"{mod}.{nm}".lower()]
            return hit[0] if hit else None
        cands = self.by_name.get(name.lower(), [])
        own = [x for x in cands if x.startswith(module + ".")]
        pub = [x for x in cands if self.procs[x]["public"] and self.procs[x]["mod"].kind == "standard"]
        return (own or pub or [None])[0]

    # ---- 手続き1つ ----

    def parse_proc(self, p):
        m = p["mod"]
        self.cur, self.binds, self.withs = p, {}, []
        self.active = None                                           # 直前の「シートを選ぶ」文で表示したシート（分岐・繰り返し・呼び出しをまたぐと、分からなくなる）
        self.calls, self.io, self.inter, self.dynamic, self.ui = [], [], [], [], []
        stmts = m.stmts[p["s0"]:p["s1"]]
        self.S, self.i = stmts, 0
        body = self.block(())
        handlers, main = [], []
        labels = {n["name"].lower(): idx for idx, n in enumerate(body) if n["kind"] == "label"}
        targets = {n["label"].lower() for n in self.walk(body) if n["kind"] == "onerror" and n.get("label")}
        cut = min((labels[t] for t in targets if t in labels), default=None)
        if cut is not None:                                          # エラー処理の節: ラベルから手続きの終わりまで
            handlers.append({"label": body[cut]["name"], "lines": [body[cut]["lines"][0], p["lines"][1]], "body": self.group(body[cut + 1:], m)})
            main = body[:cut]
        else:
            main = body
        p["body"], p["handlers"] = self.group(main, m), handlers
        p["calls"], p["io"], p["interactions"], p["dynamic"], p["ui"] = self.calls, self.io, self.inter, self.dynamic, self.ui
        nodes = list(self.walk(p["body"])) + [n for h in handlers for n in self.walk(h["body"])]
        owned = set()                                                # 1行に文が幾つもある書き方（a = 1: If … Then …／Case x: …: Exit For）: 行は、先に出てくるノードが持つ
        for n in nodes:
            if n.get("exec"):
                n["exec"] = [ln for ln in dict.fromkeys(n["exec"]) if ln not in owned]
                owned.update(n["exec"])
        p["exec"] = sorted({ln for n in nodes for ln in n.get("exec", [])})
        p["boxes"] = sum(1 for n in nodes if n["kind"] in ("call", "interact", "if", "select", "loop") or (n["kind"] == "seq" and not n.get("tech")))

    def walk(self, nodes):
        for n in nodes:
            yield n
            for b in n.get("branches") or []:
                yield from self.walk(b["body"])
            yield from self.walk(n.get("body") or [])

    def peek(self):
        return self.S[self.i][0] if self.i < len(self.S) else None

    def block(self, stop):
        """stop の語で始まる文の手前までを、ノードの並びにする"""
        out = []
        while self.i < len(self.S):
            text, a, b = self.S[self.i]
            low = text.lower()
            if any(re.match(s, low) for s in stop):
                break
            self.i += 1
            out += self.stmt(text, a, b)
        return out

    def stmt(self, text, a, b):
        shown, self.active, self.simple = self.active, None, False
        out = self.stmt_of(text, a, b, shown)
        if not self.simple:                                          # 分岐・繰り返し・飛び先などをまたいだ: 表示中のシートは、分からない
            self.active = None
        return out

    def stmt_of(self, text, a, b, shown):
        bare, low = blank_strings(text), text.lower()
        if text.startswith("@label "):
            return [{"kind": "label", "name": text[7:], "lines": [a, b]}]
        if re.match(r"(?i)^#If\b", text):                             # 条件つきコンパイル（どの枝が動くかは、コンパイルの設定で決まる）
            return [self.compile_block(text, a, b)]
        if re.match(r"(?i)^#(ElseIf|Else|End\s+If|Const)\b", text):    # 対になる #If が、手続きの外にある
            return []
        if DECL.match(text) and not re.match(r"(?i)^(Private|Public|Friend|Global|Static)\s+(Sub|Function|Property)\b", text):
            return []
        mt = re.match(r"(?i)^If\b(.*)\bThen\s*$", bare)
        if mt:
            return [self.if_block(text, a, b)]
        mt = re.match(r"(?i)^If\b(.*?)\bThen\s+(.*)$", bare)
        if mt:                                                       # 1行の If
            cond = text[2:mt.end(1)].strip()
            rest = text[mt.start(2):]
            parts = re.split(r"(?i)\bElse\b", blank_strings(rest), maxsplit=1)
            then_t = rest[:len(parts[0])].strip()
            else_t = rest[len(parts[0]) + 4:].strip() if len(parts) > 1 else ""
            node = {"kind": "if", "single": True, "head": [a, b], "lines": [a, b], "exec": list(range(a, b + 1)), "branches": []}
            self.note_head(node, cond, a)
            for label, t in (("then", then_t), ("else", else_t)):
                if t:
                    inner = [x for part in split_top(t, ":") if part.strip() for x in self.stmt(part.strip(), a, b)]
                    for x in inner:
                        x["exec"] = []                               # 行は、判断の行として数える
                    node["branches"].append({"label": label, "cond": cond if label == "then" else "", "head": [a, b], "body": inner})
            return [node]
        if re.match(r"(?i)^Select\s+Case\b", text):
            return [self.select_block(text, a, b)]
        mt = re.match(r"(?i)^(For\s+Each|For|Do\s+While|Do\s+Until|Do|While)\b(.*)$", text)
        if mt and not low.startswith("do events"):
            return [self.loop_block(a, b, mt)]
        if re.match(r"(?i)^With\b", text):
            self.withs.append(self.target_of(text[4:].strip()))
            body = self.block((r"end with$",))
            self.i += 1
            self.withs.pop()
            return body
        mt = re.match(r"(?i)^On\s+Error\s+(GoTo\s+(\S+)|Resume\s+Next)$", text)
        if mt:
            lab = mt.group(2)
            mode = "resume_next" if not lab else "off" if lab in ("0", "-1") else "goto"
            return [{"kind": "onerror", "mode": mode, "label": lab if mode == "goto" else None, "lines": [a, b], "exec": list(range(a, b + 1))}]
        mt = re.match(r"(?i)^(Exit\s+(Sub|Function|Property|For|Do)|End|Stop|GoTo\s+(\S+)|Resume(\s+\S+)?|Return)$", text)
        if mt:
            what = "goto" if low.startswith("goto") else "resume" if low.startswith("resume") else \
                "loop" if re.match(r"(?i)^Exit\s+(For|Do)", text) else "program" if low == "end" else "proc"   # End: プログラムごと終える（呼んだ側に戻らない）
            return [{"kind": "exit", "what": what, "label": mt.group(3) or (mt.group(4) or "").strip() or None, "lines": [a, b], "exec": list(range(a, b + 1))}]
        self.active, self.simple = shown, True
        return self.plain(text, a, b)

    def note_head(self, node, cond, line):
        """条件の式の中の、呼び出し・やり取り・入出力"""
        info = self.analyse(cond, line, head=True)
        node["cond"] = cond
        if info["interact"]:
            node["interact"] = info["interact"]
        if info["calls"]:
            node["calls"] = info["calls"]
        if info["io"]:
            node["io"] = info["io"]

    def if_block(self, text, a, b):
        cond = re.sub(r"(?i)\s*Then\s*$", "", text[2:]).strip()
        node = {"kind": "if", "head": [a, b], "exec": list(range(a, b + 1)), "branches": []}
        self.note_head(node, cond, a)
        label, head, bcond = "then", [a, b], cond
        while True:
            body = self.block((r"elseif\b", r"else$", r"end if$"))
            node["branches"].append({"label": label, "cond": bcond, "head": head, "body": body})
            nxt = self.peek()
            if nxt is None:
                break
            t, x, y = self.S[self.i]
            self.i += 1
            if re.match(r"(?i)^End If$", t):
                node["lines"] = [a, y]
                break
            if re.match(r"(?i)^ElseIf\b", t):
                label, head, bcond = "elseif", [x, y], re.sub(r"(?i)\s*Then\s*$", "", t[6:]).strip()
                node["exec"] += list(range(x, y + 1))
                self.analyse(bcond, x, head=True)
            else:
                label, head, bcond = "else", [x, y], ""
        node.setdefault("lines", [a, self.S[self.i - 1][2] if self.i else b])
        return node

    def compile_block(self, text, a, b):
        """条件つきコンパイル（#If … #ElseIf … #Else … #End If）を、分岐として読む。中の Exit などを、無条件のものにしない"""
        clean = lambda t: "コンパイルの条件: " + re.sub(r"(?i)\s*Then\s*$", "", t).strip()
        node = {"kind": "if", "head": [a, b], "exec": list(range(a, b + 1)), "branches": [], "cond": clean(text[3:])}
        label, head, bcond = "then", [a, b], node["cond"]
        while True:
            body = self.block((r"#elseif\b", r"#else$", r"#end if$"))
            node["branches"].append({"label": label, "cond": bcond, "head": head, "body": body})
            if self.i >= len(self.S):
                break
            t, x, y = self.S[self.i]
            self.i += 1
            if re.match(r"(?i)^#End\s+If$", t):
                node["lines"] = [a, y]
                break
            if re.match(r"(?i)^#ElseIf\b", t):
                label, head, bcond = "elseif", [x, y], clean(t[7:])
                node["exec"] += list(range(x, y + 1))
            else:
                label, head, bcond = "else", [x, y], ""
        node.setdefault("lines", [a, self.S[self.i - 1][2] if self.i else b])
        return node

    def select_block(self, text, a, b):
        expr = re.sub(r"(?i)^Select\s+Case\s*", "", text).strip()
        node = {"kind": "select", "head": [a, b], "exec": list(range(a, b + 1)), "branches": []}
        self.note_head(node, expr, a)
        self.block((r"case\b", r"end select$"))                      # Select と最初の Case の間の文は無い
        while self.peek() is not None:
            t, x, y = self.S[self.i]
            self.i += 1
            if re.match(r"(?i)^End Select$", t):
                node["lines"] = [a, y]
                break
            label = re.sub(r"(?i)^Case\s*", "", t).strip()
            body = self.block((r"case\b", r"end select$"))
            node["branches"].append({"label": "else" if label.lower() == "else" else "case", "cond": "" if label.lower() == "else" else label, "head": [x, y], "body": body})
        node.setdefault("lines", [a, b])
        return node

    def loop_block(self, a, b, mt):
        form = re.sub(r"\s+", " ", mt.group(1)).lower()
        cond = mt.group(2).strip()
        node = {"kind": "loop", "form": form, "head": [a, b], "exec": list(range(a, b + 1))}
        self.note_head(node, cond, a)
        stop = (r"next\b",) if form.startswith("for") else (r"wend$",) if form == "while" else (r"loop\b",)
        node["body"] = self.block(stop)
        if self.peek() is not None:
            t, x, y = self.S[self.i]
            self.i += 1
            node["lines"], node["tail"] = [a, y], [x, y]
            tail = re.sub(r"(?i)^Loop\s*", "", t).strip() if form.startswith("do") else ""
            if tail:                                                 # Do … Loop While 条件
                node["cond"] = tail
                node["exec"] += list(range(x, y + 1))
                self.analyse(tail, x, head=True)
        node.setdefault("lines", [a, b])
        return node

    # ---- 文の中身 ----

    def target_of(self, expr):
        """式が指す先（シート・ブックなど）。分からなければ None"""
        low = expr.strip().lower()
        mt = SHEET.search(expr)
        if mt:
            return ("シート", mt.group(2) if mt.group(2) is not None else DYNAMIC)
        head = re.match(rf"^({IDENT})", expr.strip())
        if head:
            h = head.group(1).lower()
            if h in self.binds:
                return self.binds[h]
            code = next((v for k, v in self.sheets.items() if k.lower() == h), None)
            if code:
                return ("シート", code)
            if h in ("activesheet", "selection", "activecell"):
                return ("シート", self.active or ACTIVE)
            if h == "me" and self.cur["mod"].kind == "sheet":
                return ("シート", self.cur["mod"].sheet or self.cur["mod"].name)
        if low.startswith(".") and self.withs and self.withs[-1]:
            return self.withs[-1]
        return None

    def add_io(self, out, kind, name, way, line, dynamic=None):
        item = {"kind": kind, "name": name, "way": way, "line": line}
        if (kind, name, way) not in [(x["kind"], x["name"], x["way"]) for x in out]:
            out.append(item)
        self.io.append(dict(item, procedure=self.cur["id"]))
        if dynamic:
            self.dynamic.append({"line": line, "what": kind, "expr": dynamic[:80]})

    def add_query(self, out, name, line):
        """保存してあるクエリを動かす文。クエリの定義（SQL）が手元にあれば、その SQL が読む・書くテーブルも付ける（出典に、定義の行を足す）"""
        items, via = self.access.query_io(name) if self.access else ([("クエリ", name, "out")], None)
        for kind, nm, way in items:
            self.add_io(out, kind, nm, way, line)
            hit = next(x for x in out if (x["kind"], x["name"], x["way"]) == (kind, nm, way))
            if via and (kind, nm) != items[0][:2]:
                hit["via"] = via

    def lead_name(self, bare):
        """文そのものが呼び出しであるときの、呼び先の書き方（Call F／F 引数／x = F(…) の F）。ほかの文の式に埋まった呼び出しだけなら None"""
        t = re.sub(r"(?i)^Call\s+", "", bare.strip())
        mt = re.match(rf"(?i)^(?:(?:Set|Let)\s+)?{IDENT}\s*=\s*(.+)$", t)
        if mt:                                                       # 変数に入れる文: 右辺が、呼び出し1つだけのとき
            m2 = re.match(rf"(?i)^(?:New\s+)?((?:{IDENT}\.)*{IDENT})\s*(\(.*\))?$", mt.group(1).strip())
            if not m2:
                return None
            arg = m2.group(2)
            if arg:
                depth = 0
                for k, ch in enumerate(arg):
                    depth += (ch == "(") - (ch == ")")
                    if depth == 0 and k < len(arg) - 1:
                        return None                                  # かっこのあとに、式が続く（F(1) + G(2)）
            return m2.group(1)
        m2 = re.match(rf"(?i)^((?:{IDENT}\.)*{IDENT})(?=\s|\(|$)", t)
        return m2.group(1) if m2 else None

    def analyse(self, text, line, head=False):
        """文（か条件の式）の中の、呼び出し・操作者とのやり取り・入出力。{"calls", "lead", "interact", "io"}。lead: 文そのものが呼び出しであるときの呼び先"""
        bare, m, out = blank_strings(text), self.cur["mod"], {"calls": [], "lead": None, "interact": None, "io": []}
        lead = None
        if not head:
            ln = self.lead_name(bare)
            lead = self.resolve(ln, m.name) if ln else None
            if lead is None and ln and re.match(r"(?i)^(?:Call\s+)?(?:Work)?sheets$", ln):   # シートのモジュールの手続きを、シート名で呼ぶ文
                x = re.match(rf'(?i)^(?:Call\s+)?(?:Work)?sheets\s*\(\s*"([^"]+)"\s*\)\.({IDENT})', text)
                mod = next((mm.name for mm in self.mods if x and (mm.sheet or "").lower() == x.group(1).lower()), None)
                lead = self.resolve(f"{mod}.{x.group(2)}", m.name) if mod else None
            if lead is None and re.match(r"(?i)^(?:Set\s+)?" + IDENT + r"\s*=\s*New\s+(" + IDENT + r")\s*$", bare):
                lead = self.resolve(re.search(rf"(?i)New\s+({IDENT})\s*$", bare).group(1) + ".Class_Initialize", m.name)
        # 呼び出し
        mt = re.match(r"(?i)^(?:Application\.)?Run\s*\(?\s*\"([^\"]+)\"", text)
        names = [(mt.group(1).split("!")[-1], True)] if mt else []
        if mt and not head:
            lead = self.resolve(mt.group(1).split("!")[-1], m.name) or lead
        for x in re.finditer(rf"(?<![\w.\u0080-￿])((?:{IDENT}\.)?{IDENT})(?=\s*(\(|$|\s))", bare):
            names.append((x.group(1), x.start() == 0 or bare[x.end():x.end() + 1] == "(" or bool(re.match(r"(?i)^Call\s+$", bare[:x.start()]))))
        for x in re.finditer(rf'(?i)\b(?:Work)?sheets\s*\(\s*"([^"]+)"\s*\)\.({IDENT})', text):   # シートのモジュールの手続きを、シート名で呼ぶ
            mod = next((mm.name for mm in self.mods if (mm.sheet or "").lower() == x.group(1).lower()), None)
            if mod:
                names.append((f"{mod}.{x.group(2)}", True))
        for x in re.finditer(rf"(?i)\bNew\s+({IDENT})\b", bare):       # オブジェクトを作ると、クラスの初期化（Class_Initialize）が動く
            names.append((f"{x.group(1)}.Class_Initialize", True))
        for nm, sure in names:
            pid = self.resolve(nm, m.name)
            if pid and (pid != self.cur["id"] or sure and not re.match(rf"(?i)^{re.escape(nm)}\s*=", bare)) and not (
                    pid == self.cur["id"] and re.match(rf"(?i)^{re.escape(nm)}\s*=", bare)):
                if pid == self.cur["id"] and not sure:
                    continue
                if pid not in out["calls"]:
                    out["calls"].append(pid)
                    getter = pid != lead and str(self.procs[pid].get("kind", "")).lower().startswith("property")   # 式の中の、値を取るだけのプロパティ（図には出さない）
                    self.calls.append({"to": pid, "line": line, "head": head or getter})
                    if getter:
                        out.setdefault("quiet", []).append(pid)
        mt = re.search(r'(?i)\bDoCmd\.RunMacro\s+"([^"]+)"', text)
        if mt and self.access:                                       # Access のマクロを呼ぶ
            pid = self.access.macro_pid(mt.group(1))
            if pid and pid not in out["calls"]:
                out["calls"].append(pid)
                self.calls.append({"to": pid, "line": line, "head": head})
            if pid and not head and re.match(r"(?i)^DoCmd\.RunMacro\b", bare):
                lead = pid
        out["lead"] = lead if lead in out["calls"] else None
        mt = re.match(rf"(?i)^(?:Call\s+)?((?:{IDENT}\.)?{IDENT})\b", bare)
        if mt and not out["calls"] and not head and re.match(r"(?i)^Call\s", bare):
            self.calls.append({"to": None, "name": mt.group(1), "line": line, "head": False})   # コードの中に見つからない呼び先
        # 操作者とのやり取り
        for how, rx in INTERACT:
            if rx.search(bare):
                out["interact"] = {"how": how, "text": first_string(text)}
                break
        show = re.match(rf"(?i)^({IDENT})\.Show\b", bare)
        if show and (show.group(1).lower() in self.forms or show.group(1).lower() == "me"):
            out["interact"] = {"how": "form", "text": show.group(1)}
        if out["interact"]:
            self.inter.append(dict(out["interact"], line=line))
        if not head and UI_WRITE.search(bare):
            self.ui.append(line)
        # 束ねる（Set 変数 = 対象）
        mt = re.match(rf"(?i)^Set\s+({IDENT})\s*=\s*(.+)$", text)
        if mt:
            tg = self.target_of(mt.group(2))
            if mt.group(2).strip().lower().startswith("workbooks.open"):
                tg = None
            if tg:
                self.binds[mt.group(1).lower()] = tg
                out["bound"] = True                                  # 変数に対象を入れただけの文（読み書きは、その変数を使う文で数える）
                return out
        self.find_io(text, bare, line, out["io"], head)
        return out

    def find_io(self, text, bare, line, out, head=False):
        low = bare.lower()
        parts = split_top(text, "=") if not head and not re.match(r"(?i)^(If|ElseIf|Do|Loop|While|Case|Select|For|Call|Set)\b", text) else [text]   # 条件の式の = は、比べているだけ
        lhs = parts[0] if len(parts) > 1 and not re.search(r"[<>]\s*$", parts[0]) and ":" not in parts[0][-1:] else ""
        # ファイル
        fname = lambda expr: file_name(expr, self.consts)
        mt = re.search(r"(?i)\bWorkbooks\.Open(?:Text)?\s*\(?\s*(?:Filename\s*:=\s*)?([^,)]+)", text)
        if mt:
            nm, dyn = fname(mt.group(1))
            self.add_io(out, "ファイル", nm, "in", line, mt.group(1) if dyn else None)
            st = re.match(rf"(?i)^Set\s+({IDENT})\s*=", text)
            if st:
                self.binds[st.group(1).lower()] = ("ファイル", nm)
        mt = re.search(r"(?i)\.(SaveAs|SaveCopyAs|ExportAsFixedFormat)\b(.*)$", text)
        if mt:
            arg = re.search(r'(?i)Filename\s*:=\s*([^,)]+)', mt.group(2)) or re.match(r"\s*\(?\s*([^,)]+)", mt.group(2))
            nm, dyn = fname(arg.group(1)) if arg else (DYNAMIC, True)
            self.add_io(out, "ファイル", nm, "out", line, (arg.group(1) if arg else "") if dyn else None)
        elif re.search(r"(?i)\b(ThisWorkbook|ActiveWorkbook)\.Save\b", bare):
            self.add_io(out, "ファイル", "（このブック）", "out", line)
        mt = re.match(r"(?i)^Open\s+(.+?)\s+For\s+(Input|Output|Append|Binary|Random)\b", text)
        if mt:
            nm, dyn = fname(mt.group(1))
            self.add_io(out, "ファイル", nm, "in" if mt.group(2).lower() == "input" else "out", line, mt.group(1) if dyn else None)
        mt = re.match(r"(?i)^(Kill|MkDir|RmDir)\s+(.+)$", text)
        if mt:
            nm, dyn = fname(mt.group(2))
            self.add_io(out, "ファイル", nm, "out", line, mt.group(2) if dyn else None)
        mt = re.match(r"(?i)^(FileCopy|Name)\s+(.+?)(?:,|\s+As\s+)(.+)$", text)
        if mt:
            for expr, way in ((mt.group(2), "in"), (mt.group(3), "out")):
                nm, dyn = fname(expr)
                self.add_io(out, "ファイル", nm, way, line, expr if dyn else None)
        mt = re.search(r"(?i)\.(CreateTextFile|OpenTextFile|CopyFile|MoveFile|DeleteFile)\s*\(?\s*([^,)]+)", text)
        if mt:
            nm, dyn = fname(mt.group(2))
            self.add_io(out, "ファイル", nm, "in" if mt.group(1).lower() == "opentextfile" else "out", line, mt.group(2) if dyn else None)
        # メール・外部プログラム・印刷
        if re.search(r"(?i)(Outlook\.Application|\.CreateItem\s*\(|CDO\.Message|\.SendMail\b|\bSendObject\b)", text) or re.search(r"(?i)\.Send\s*$", bare):
            self.add_io(out, "メール", "Outlook" if "outlook" in text.lower() or re.search(r"(?i)\.(Send|CreateItem)", bare) else "メール", "out", line)
        mt = re.search(r'(?i)\bShell\s*\(?\s*([^,)]+)', text) if re.search(r"(?i)(^|[\s=(])Shell\b", bare) else None
        if mt:
            nm = literal(mt.group(1))
            self.add_io(out, "外部プログラム", (nm or DYNAMIC).split(" ")[0].split("\\")[-1] or DYNAMIC, "out", line, None if nm else mt.group(1))
        if re.search(r"(?i)\.(PrintOut|PrintPreview)\b", bare) or re.search(r"(?i)\bDoCmd\.OpenReport\b", bare):
            rep = re.search(r'(?i)DoCmd\.OpenReport\s+"([^"]+)"', text)
            self.add_io(out, "レポート" if rep else "印刷", rep.group(1) if rep else "印刷", "out", line)
        # データベース（接続の文字・SQL の文字・Access の操作）
        relink = bool(re.search(r"(?i)\.Connect\s*$", lhs))             # リンクしたテーブル・クエリの接続先を決めるだけの文（読み書きは、そのテーブルを使う文で数える）
        for sql in re.findall(r'"((?:[^"]|"")*)"', text):
            if not relink and re.search(r"(?i)\b(Provider|Data Source|DSN|Driver|Dbq)\s*=", sql):
                keys = {k.strip().lower(): v.strip() for k, _, v in (x.partition("=") for x in sql.split(";")) if v.strip()}
                first = sql.split(";")[0].strip().lower()
                if first and "=" not in first and first != "odbc":   # Access のリンクの文字（Text;… Excel 12.0;…）の DSN は、取り込みの定義の名前。データベースではない
                    keys.pop("dsn", None)
                nm = next((keys[k] for k in ("initial catalog", "database", "dbq", "dsn", "data source", "server") if k in keys), "")
                if nm:
                    self.add_io(out, "データベース", pathlib.PureWindowsPath(nm).name, "in", line)
            if re.match(r"(?i)\s*(SELECT|INSERT|UPDATE|DELETE)\b", sql):
                writes, reads = extract_access.sql_tables(sql.replace('""', '"'))   # SQL の読み方は 1 つ（extract_access.sql_tables）
                for t in writes:
                    self.add_io(out, "テーブル", t, "out", line)
                for t in reads:
                    self.add_io(out, "テーブル", t, "in", line)
        if re.search(r"(?i)\b(DoCmd\.RunSQL|\.Execute|OpenRecordset)\b", bare) and not re.search(r'"\s*(SELECT|INSERT|UPDATE|DELETE)\b', text, re.I):
            arg = re.search(r'(?i)(?:RunSQL|Execute|OpenRecordset)\s*\(?\s*([^,)]+)', text)
            nm = literal(arg.group(1)) if arg else None
            if nm and self.access and nm.lower() in self.access.queries:
                self.add_query(out, nm, line)
            elif nm:
                self.add_io(out, "テーブル", nm, "in" if "openrecordset" in low else "out", line)
            else:
                self.add_io(out, "テーブル", DYNAMIC, "in" if "openrecordset" in low else "out", line, arg.group(1) if arg else "")
        mt = re.search(r'(?i)\bDoCmd\.OpenQuery\s+"([^"]+)"', text)
        if mt:
            self.add_query(out, mt.group(1), line)
        mt = re.search(r'(?i)\bDoCmd\.(TransferSpreadsheet|TransferText|OutputTo)\s+(.*)$', text)
        if mt:                                                       # 引数の位置: 種類, （定義・形式）, テーブル, ファイル（OutputTo は 種類, 名前, 形式, ファイル）
            how = mt.group(1).lower()
            args = [x.strip() for x in split_top(mt.group(2), ",")]
            named = {k.strip().lower(): v.strip() for k, _, v in (x.partition(":=") for x in args) if v}
            pos = lambda i: args[i] if i < len(args) and ":=" not in args[i] else ""
            kind = named.get("transfertype") or pos(0)
            way = "out" if how == "outputto" or "export" in kind.lower() else "in"
            table = named.get("tablename") or named.get("objectname") or pos(1 if how == "outputto" else 2)
            file = named.get("filename") or named.get("outputfile") or pos(3)
            if literal(table):
                what = "クエリ" if self.access and literal(table).lower() in self.access.queries else "テーブル"
                self.add_io(out, what, literal(table), "in" if way == "out" else "out", line)
            elif table:
                self.add_io(out, "テーブル", DYNAMIC, "in" if way == "out" else "out", line, table)
            if file:
                nm, dyn = fname(file)
                self.add_io(out, "ファイル", nm, way, line, file if dyn else None)
        for _fn, t in re.findall(r'(?i)\b(DLookup|DCount|DSum|DMax|DMin|DAvg)\s*\(\s*"[^"]*"\s*,\s*"([^"]+)"', text):
            t = t.strip().strip("[]")
            if self.access and t.lower() in self.access.queries:     # 相手が、保存してあるクエリ（テーブルと呼ばない。そのクエリが読むテーブルも付く）
                self.add_query(out, t, line)
            else:
                self.add_io(out, "テーブル", t, "in", line)
        # シート
        if NAVIGATE.match(bare) or format_only(text):                 # 選ぶだけ・書式だけの文: データの読み書きではない
            return
        seen = []
        for mt in SHEET.finditer(text):
            owner = re.search(rf"({IDENT})\s*\.\s*$", text[:mt.start()])
            book = self.binds.get(owner.group(1).lower()) if owner else None
            if book and book[0] == "ファイル":                       # 開いたファイルのシート（wb.Sheets(1)）は、そのファイルを触っている
                seen.append((mt.start(), book, None))
                continue
            seen.append((mt.start(), ("シート", mt.group(2) if mt.group(2) is not None else DYNAMIC), None if mt.group(2) is not None else mt.group(1)))
        for mt in re.finditer(rf"(?<![\w.\u0080-￿\")])({IDENT})(?=\s*\.)", bare):
            tg = self.target_of(mt.group(1))
            if tg and tg[0] == "シート" and not re.match(r"(?i)\s*\.\s*(?:Work)?sheets\s*\(", bare[mt.end():]):
                seen.append((mt.start(), tg, None))
        for mt in re.finditer(r"(?<![\w.\u0080-￿\")])\.(Range|Cells|Rows|Columns|UsedRange|Name|Copy|Paste\w*|Clear\w*|Sort|AutoFilter|ListObjects|PivotTables)\b", bare):
            if self.withs and self.withs[-1] and self.withs[-1][0] == "シート":
                seen.append((mt.start(), self.withs[-1], None))
        for mt in re.finditer(r"(?i)(?<![\w.\u0080-￿])(Range|Cells|Rows|Columns)\s*\(", bare):
            own = self.cur["mod"]
            seen.append((mt.start(), ("シート", (own.sheet or own.name) if own.kind == "sheet" else self.active or ACTIVE), None))
        dest = re.search(r"(?i)\.(Copy|Cut)\s+(?=\S)", bare)                # 写す先を、引数にそのまま書く形（Range(…).Copy 写す先）
        if lhs and re.search(r"(?i)\.(Visible|ScrollArea|DisplayPageBreaks|EnableSelection)\s*$", lhs):
            return                                                   # 表示・非表示などの切り替え（データの読み書きではない）
        to = min((at for at, _tg, _dyn in seen if dest and at >= dest.end()), default=None)   # 写す先は、引数の先頭の1つ（その引数の中で読むセルは、読む）
        for at, tg, dyn in sorted(seen, key=lambda x: x[0]):
            way = "out" if (lhs and at < len(lhs) and ("." in lhs or "(" in lhs)) or (not lhs and WRITE_METHOD.search(bare[at:])) else "in"
            if re.search(r"(?i)\bDestination\s*:=", bare[:at]) or at == to:
                way = "out"
            elif (dest and not lhs) or head:                         # 条件の式は、読むだけ
                way = "in"
            self.add_io(out, tg[0], tg[1], way, line, dyn)

    def plain(self, text, a, b):
        """分岐でも繰り返しでもない文1つ → ノードの並び。式の中で呼んでいる手続きは、先に呼び出しのノードにする（行は、文のノードが持つ）。
        こうしないと、引数の中で関数を呼ぶ文（シートに書く・メニューを出す）が、呼び先の中身に置き換わって、文そのものが図から消える"""
        info = self.analyse(text, a)
        lead = info["lead"] if not info["interact"] else None
        node = {"kind": "interact" if info["interact"] else "call" if lead else "stmt", "lines": [a, b], "exec": list(range(a, b + 1)), "text": text[:120]}
        tech = "変数の準備" if info.get("bound") else next((why for why, rx in TECH if rx.match(text)), None)
        if tech == "後片付け" and re.match(rf"(?i)^Set\s+{re.escape(self.cur['name'])}\s*=", text):
            tech = None                                              # 戻り値を「無し」にする文（Set 関数名 = Nothing）は、後片付けではない
        mt = COPY_ONLY.match(blank_strings(text))                    # 数を足す文は、業務の件数のこともあるので、機械では決めない
        if mt and not tech and self.working(mt.group(1), mt.group(2)):
            tech = "変数の準備"
        if not tech and node["kind"] == "stmt" and format_only(text):
            tech = "書式・表示の調整"
        fetch = not tech and node["kind"] == "stmt" and self.fetch_only(text)   # 取るだけ・進めるだけの文は、読む先があっても「変数の準備」（決まり 6.10）
        if fetch:
            tech = "変数の準備"
        if tech and node["kind"] == "stmt" and (not info["io"] or (fetch and all(x["way"] == "in" for x in info["io"]))):
            node["tech"] = tech
        if info["interact"]:
            node["interact"] = info["interact"]
        if lead:
            node["to"] = lead
        if info["io"] and not (lead and re.match(rf'(?i)^(?:Call\s+)?(?:Work)?sheets\s*\(\s*"[^"]+"\s*\)\.{IDENT}(\s|\(|$)', text)):   # シート名で手続きを呼ぶ文は、シートを読んでいない
            node["io"] = info["io"]
        inner = [x for x in info["calls"] if x != lead and x not in info.get("quiet", [])]
        if inner and node["kind"] == "stmt" and not node.get("io") and re.match(rf"(?i)^(?:Set\s+|Let\s+)?{IDENT}\s*=", blank_strings(text)):
            node["kind"], node["to"] = "call", inner.pop()           # 変数に、呼んだ結果を入れる・足すだけの文（x = x + F(1)）: 最後に呼ぶ手続きの呼び出しが、文を持つ（箱を増やさない）
            node.pop("tech", None)
        if info["calls"]:                                            # 呼んだ先が、表示するシートを変えているかもしれない
            self.active = None
        mt = re.match(r"(?i)^(.+?)\.(Activate|Select|Visible\s*=.*)$", blank_strings(text))
        if mt and node["kind"] == "stmt" and not node.get("io"):     # シートを選ぶ・出す・隠すだけの文: 読み書きは無いが、まとまりを切る境になる（別のシートの作業が、1つの箱に混ざらない）
            tg = self.target_of(text[:mt.end(1)])
            if tg and tg[0] == "シート" and tg[1] != DYNAMIC:
                node["at"] = tg
        mt = re.match(r"(?i)^(.+)\.(Activate|Select)\s*$", blank_strings(text))
        if mt:                                                       # シートを選ぶ文: 次の文からの「表示中のシート」が決まる（セルを選ぶ文では、変わらない）
            tg = self.target_of(text[:mt.end(1)]) if not re.search(r"(?i)\.(Range|Cells|Rows|Columns|Shapes)\s*\(|^\s*(Range|Cells|Rows|Columns|Selection|ActiveCell)\b", text) else ("keep",)
            if tg != ("keep",):
                self.active = tg[1] if tg and tg[0] == "シート" and tg[1] not in (DYNAMIC, ACTIVE) else None
        return [{"kind": "call", "lines": [a, b], "exec": [], "to": x, "text": text[:120]} for x in inner] + [node]

    # ---- 続く文をまとめる ----

    def group(self, nodes, m):
        """文（stmt）の並びを、まとまり（seq）にする。区切り: 分岐・繰り返し・呼び出し・やり取り、触る先の切り替わり、コメントだけの行"""
        out, cur = [], None

        def close():
            nonlocal cur
            if cur:
                out.append(cur)
            cur = None
        for n in nodes:
            for key in ("body",):
                if n.get(key) is not None:
                    n[key] = self.group(n[key], m)
            for br in n.get("branches") or []:
                br["body"] = self.group(br["body"], m)
            if n["kind"] == "label":
                close()
                out.append(n)
                continue
            if n["kind"] != "stmt":
                close()
                out.append(n)
                continue
            noted = m.comment_at.get(n["lines"][0] - 1)              # すぐ上の、コメントだけの行（名付けの手がかり）
            tg = next(((x["kind"], x["name"]) for x in n.get("io") or []), None) or (tuple(n["at"]) if n.get("at") else None)
            if cur and (noted or (tg and cur["target"] and tg != cur["target"]) or cur.get("tech") != n.get("tech")):
                close()
            if cur is None:
                cur = {"kind": "seq", "lines": list(n["lines"]), "exec": [], "io": [], "target": None, "hint": noted or ""}
                if n.get("tech"):
                    cur["tech"] = n["tech"]                          # 技術的な文だけのまとまり（図にしない候補。理由つき）
            cur["lines"][1] = n["lines"][1]
            cur["exec"] += n["exec"]
            cur["target"] = cur["target"] or tg
            for x in n.get("io") or []:
                same = next((y for y in cur["io"] if (y["kind"], y["name"], y["way"]) == (x["kind"], x["name"], x["way"])), None)
                if same is None:
                    cur["io"].append(x)
                elif x["line"] != same["line"] and x["line"] not in same.get("also", []):
                    same.setdefault("also", []).append(x["line"])    # 同じ先を、ほかの行でも読む・書く（まとまりを分けたときに、その行の側にも付ける）
        close()
        for n in out:
            n.pop("target", None)
        return out


def vba_entries(parser, shortcuts=(), access=None, embedded=None, ribbon=(), unread=()):
    """起動点: ボタン・イベント・自動実行・定時実行・ショートカット・引数の無い公開の Sub（マクロの一覧から実行できる）。
    Access のフォーム・レポートは、定義でイベントに割り当ててあるものだけを起動点にする（コード・マクロ・関数の式・埋め込みマクロ）"""
    import extract_access
    out, seen = [], set()
    layouts = access.layouts if access else {}

    def add(pid, trigger, label, why):
        if pid in parser.procs and (pid, trigger, label) not in seen:
            seen.add((pid, trigger, label))
            out.append({"procedure": pid, "trigger": trigger, "label": label, "evidence": why})
    shapes = {}                                                      # 同じマクロを、いくつもの図形に割り当ててあることがある（起動点は1つにまとめる）
    for b in parser.buttons:
        pid = parser.resolve(b["macro"], "")
        if pid:
            shapes.setdefault(pid, []).append(b)
    for pid, bs in shapes.items():
        b = bs[0]
        more = f"（ほか {len(bs) - 1} 個の図形にも割り当て）" if len(bs) > 1 else ""
        add(pid, "button", f"シート「{b['sheet']}」の「{b['shape']}」ボタン{more}", f"ボタンに割り当てられたマクロ {b['macro']}")
    for b in ribbon:                                                 # リボンの定義（customUI）
        pid = parser.resolve(b["macro"], "")
        if pid:
            add(pid, "auto" if b.get("what") == "load" else "button",
                "ブックを開いたとき（リボンの読み込み）" if b.get("what") == "load" else f"リボンの「{b['label']}」", f"リボンの定義（customUI）の割り当て {b['macro']}")
    called = {c["to"] for p in parser.procs.values() for c in p["calls"] if c["to"]}
    for m in parser.mods:
        lay = layouts.get(m.name.lower())
        if lay is None or lay.get("unread"):
            continue
        import extract_access
        labels = [extract_access.event_label(lay, lay["name"], ev) for ev in lay["events"]]
        for ev in lay["events"]:
            label = extract_access.event_label(lay, lay["name"], ev)
            if labels.count(label) > 1 and ev["owner_kind"] == "control":   # 同じ文字のボタンが幾つもある
                label += f"（{ev['owner']}）"
            if access.startup and access.startup == lay["name"] and ev["owner_kind"] == "top" and ev["event"].lower() in ("open", "load"):
                label += "（データベースを開くと、最初に開く）"
            trig = "timer" if ev["event"].lower() == "timer" else "button" if "ボタン" in label else "form" if lay["object"] == "form" else "event"
            why = f"{'フォーム' if lay['object'] == 'form' else 'レポート'}の定義 {ev['line']} 行（{ev['prop']} ＝ {ev['value'][:40]}）"
            if ev.get("guess"):                                      # 直接読み: 割り当てはあるが、イベントの種類は、手続きの名前から決めた
                why += "。イベントの種類は、手続きの名前から"
            if ev["kind"] == "code":
                pid = f"{m.name}.{ev['handler']}"
                pid = pid if pid in parser.procs else next((x for x in parser.procs if x.lower() == pid.lower()), None)
            elif ev["kind"] == "macro":
                pid = access.macro_pid(ev["value"])
            elif ev["kind"] == "expr":
                pid = parser.resolve(re.sub(r"^=\s*|\s*\(.*$", "", ev["value"]), m.name)
            else:
                pid = (embedded or {}).get((m.name.lower(), f"{ev['owner']}.{ev['prop']}"))
            if pid:
                add(pid, trig, label, why)
            elif ev["kind"] != "embedded":
                access.unread.append({"what": "イベントの割り当て先", "where": f"{m.name} {ev['line']} 行", "why": f"{label} → {ev['value']} が、書き出しの中に見つからない"})
    for m in parser.mods:
        lay = layouts.get(m.name.lower())
        if lay is not None and not lay.get("unread"):
            continue
        if lay is not None:                                          # 直接読みで、定義（イベントの割り当て）を読めていないフォーム・レポート: 手続きの名前から、候補を出す
            word = "フォーム" if lay["object"] == "form" else "レポート"
            found = 0
            for h in m.heads:
                if VBA_EVENT.match(h["name"]) and h.get("kind", "sub") != "function":
                    ctl, _, ev = h["name"].rpartition("_")
                    top = ctl.lower() in ("form", "report")
                    what = extract_access.TOP_EVENT_JA.get(ev.lower(), f"のイベント {ev}") if top else f"の「{ctl}」{CTL_EVENT_JA.get(ev.lower(), f'のイベント {ev}')}"
                    add(f"{m.name}.{h['name']}", "form" if lay["object"] == "form" else "event", f"{word}「{lay['name']}」{what}（割り当ては未確認）",
                        f"手続きの名前 {h['name']}（定義を読めていないので、イベントに割り当ててあるかは未確認）")
                    found += 1
            access.unread.append({"what": f"{word}の定義", "where": lay["name"], "why": lay["unread"] + (f"。手続きの名前から、起動点の候補を {found} 件出した" if found else "")})
            continue
        for h in m.heads:
            pid, low = f"{m.name}.{h['name']}", h["name"].lower()
            if VBA_EVENT.match(h["name"]) and (m.kind != "standard" or low.startswith("auto")) and not (m.kind == "class" and low.startswith("class_")):
                ctl, _, ev = h["name"].rpartition("_")
                if low in EVENT_JA:
                    what = EVENT_JA[low]
                elif ctl.lower() in ("workbook", "worksheet", "chart", "userform", "form", "report", "class", "document"):
                    what = f"イベント {h['name']}"
                else:
                    what = f"「{ctl}」{CTL_EVENT_JA.get(ev.lower(), f'のイベント {ev}')}"
                where = {"sheet": f"シート「{m.sheet or m.name}」", "form": f"フォーム「{m.name}」", "book": "ブック", "report": f"レポート「{m.name}」"}.get(m.kind, f"「{m.name}」")
                add(pid, "auto" if low.startswith(("auto", "workbook_open")) else "form" if m.kind == "form" else "event",
                    what if m.kind in ("book", "standard") else f"{where}の{what}" if what.startswith("「") else f"{where}: {what}", f"手続きの名前 {h['name']}")
    for p in parser.procs.values():                                  # 定時実行・ショートカット・コードでのボタンの割り当て
        if "mod" not in p:                                           # Access のマクロ（VBA の文は無い）
            continue
        m = p["mod"]
        for text, a, _b in m.stmts[p["s0"]:p["s1"]]:
            for rx, trig, word in ((r'(?i)\bOnTime\b.*?"([^"]+)"\s*(?:,|\)|$)', "timer", "時刻を決めて実行（Application.OnTime）"),
                                   (r'(?i)\bOnKey\s*\(?\s*"([^"]+)"\s*,\s*"([^"]+)"', "shortcut", "キーの割り当て（Application.OnKey）"),
                                   (r'(?i)\.OnAction\s*=.*"([^"]+)"\s*$', "button", "コードで割り当てたボタン・メニュー（OnAction）")):
                mt = re.search(rx, text)
                if mt:
                    name = mt.group(mt.lastindex).split("!")[-1].strip("'")
                    if name.startswith("."):                         # 式で組み立てた割り当て（シート名 & ".手続き"）は、手続きの名前だけで引く
                        cands = parser.by_name.get(name[1:].lower(), [])
                        pid = cands[0] if len(cands) == 1 else None
                    else:
                        pid = parser.resolve(name, m.name)
                    if pid and trig == "button":                     # どのボタン・メニューかを、同じ文の名前から添える（無ければ、手続きの名前）
                        ctl = re.search(r'"([^"]+)"\s*\)\.OnAction', text)
                        add(pid, trig, word.replace("（OnAction）", f"「{ctl.group(1)}」（OnAction）" if ctl else f"（OnAction → {pid.split('.')[-1]}）"), f"{m.name} {a} 行")
                    elif pid:
                        add(pid, trig, word + (f"「{mt.group(1)}」" if trig == "shortcut" else ""), f"{m.name} {a} 行")
    key_label = lambda k: f"ショートカット Ctrl+{'Shift+' if k.isupper() else ''}{k.upper()}"   # 大文字で割り当ててあれば、Shift も押す
    for s in shortcuts:                                              # Office のファイルから取り出した、ショートカットの割り当て
        if s["key"].strip():
            add(f"{s['module']}.{s['procedure']}", "shortcut", key_label(s["key"]), "モジュールの属性（VB_Invoke_Func）")
    for m in parser.mods:                                            # 書き出したモジュールの、ショートカットの割り当て（割り当てが無いと、キーの所は空白）
        for ln in m.lines:
            mt = re.match(rf'(?i)^Attribute\s+({IDENT})\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^\s"\\])', ln)
            if mt and f"{m.name}.{mt.group(1)}" in parser.procs:
                add(f"{m.name}.{mt.group(1)}", "shortcut", key_label(mt.group(2)), "モジュールの属性（VB_Invoke_Func）")
    has = {e["procedure"] for e in out}
    for m in parser.mods:
        for h in m.heads:
            pid = f"{m.name}.{h['name']}"
            if access:                                               # Access には、VBA の Sub を選んで実行する一覧が無い（呼ばれていなければ「たどれない手続き」）
                break
            if (m.kind in ("standard", "book", "sheet") and h["kind"] == "sub" and h["public"] and not h["params"] and pid not in has and pid not in called
                    and (m.kind == "standard" or not VBA_EVENT.match(h["name"]))):
                add(pid, "manual", f"マクロ「{h['name']}」を実行（ボタンの割り当ては{'読めていない' if unread else '見つからない'}）", "引数の無い公開の Sub で、ほかから呼ばれていない")
    return out


def module_name(lines, fallback):
    """書き出したモジュールの名前（Attribute VB_Name）。無ければファイル名"""
    for ln in lines[:40]:
        mt = re.match(r'(?i)^Attribute\s+VB_Name\s*=\s*"([^"]+)"', ln)
        if mt:
            return mt.group(1)
    return fallback


def module_kind(name, ext, sheets=None, book=""):
    """モジュールの種類: standard 標準モジュール／book ブック／sheet シート／form フォーム／report レポート／class クラス。
    book: ブックのコード名（言語によって ThisWorkbook でないことがある）"""
    low = name.lower()
    if low in ("thisworkbook", "thisdocument") or (book and low == book.lower()):
        return "book"
    if sheets and any(k.lower() == low for k in sheets):
        return "sheet"
    if low.startswith("form_"):
        return "form"
    if low.startswith("report_"):
        return "report"
    return {".frm": "form", ".cls": "class"}.get(ext, "standard")


# ---------- まとめる（言語によらない） ----------

def outline(language, lines):
    """手続きの範囲 [(開始の行, 終了の行)]（ページを切る所を決めるのに使う）"""
    if language in ("vba", "vbscript"):
        return [tuple(h["lines"]) for h in VbaModule("m", "standard", lines).heads]
    if language == "python":
        import ast
        try:
            return [(n.lineno, n.end_lineno) for n in ast.parse("\n".join(lines)).body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))]
        except (SyntaxError, ValueError):
            return []
    return []


def page_cuts(language, lines, limit):
    """ページの先頭の行の並び（1 から）。limit 行を超えるソースを、手続きの切れ目で切る。1つの手続きが limit を超えるときだけ、途中で切る"""
    n, starts = len(lines), []
    for a, _b in outline(language, lines):
        while a > 1 and lines[a - 2].strip().startswith(("'", "#", "//", "--", "*")):   # 手続きのすぐ上のコメントは、同じページに入れる
            a -= 1
        starts.append(a)
    cuts, at = [1], 1
    while n - at + 1 > limit:
        cand = [x for x in starts if at < x <= at + limit]
        at = max(cand) if cand else at + limit
        cuts.append(at)
    return cuts


def reach(procs, start):
    seen, todo = [], [start]
    while todo:
        u = todo.pop(0)
        if u in seen or u not in procs:
            continue
        seen.append(u)
        todo += [c["to"] for c in procs[u]["calls"] if c.get("to")]
    return seen


def mark_ends(procs):
    """いつもプログラムごと終える手続き（最後の文が、プログラムの終了か、そういう手続きの呼び出し）に印を付ける。呼んだ側の図で、そのあとを「終了」につなぐ"""
    ends, more = set(), True
    while more:
        more = False
        for pid, p in procs.items():
            last = p["body"][-1] if p.get("body") else None
            if pid not in ends and last and ((last["kind"] == "exit" and last.get("what") == "program") or (last["kind"] == "call" and last.get("to") in ends)):
                ends.add(pid)
                more = True
    for pid in ends:
        procs[pid]["ends"] = True


def plan_units(procs, entries):
    """アクティビティ図にする単位と、呼ぶ側の図に入れる手続き。(units, inline, unreached, blank)。
    blank: 起動点からたどれるが、図にする文が1つも無い手続き（異常終了の手続きなど）。図にせず、呼ぶ側では呼び出しを対象外にする"""
    roots = list(dict.fromkeys(e["procedure"] for e in entries))
    reached = list(dict.fromkeys(p for r in roots for p in reach(procs, r)))
    sites = {}
    for u in reached:
        for c in procs[u]["calls"]:
            if c.get("to") in procs:
                sites.setdefault(c["to"], []).append((u, c["line"]))
    inline, shared, blank = [], [], []
    for p in reached:
        if p in roots:
            continue
        cyc = p in reach(procs, p)[1:]
        if procs[p]["boxes"] == 0 and not procs[p].get("handlers") and len(set(sites.get(p, []))) != 1:
            blank.append(p)
            continue
        if len(set(sites.get(p, []))) == 1 and procs[p]["boxes"] <= INLINE_MAX and not cyc and not any(c.get("head") for u in reached for c in procs[u]["calls"] if c.get("to") == p):
            inline.append(p)
        else:
            shared.append(p)
    units = [{"kind": "entry", "entry": e["id"], "procedure": e["procedure"]} for e in entries]
    units += [{"kind": "shared", "procedure": p} for p in shared if p not in roots]
    return units, inline, [p for p in procs if p not in reached], blank


def ends_program(nodes):
    """そのブロックの中に、プログラムごと終える文があるか"""
    for n in nodes or []:
        if n.get("kind") == "exit" and n.get("what") == "program":
            return True
        if ends_program(n.get("body")) or any(ends_program(b.get("body")) for b in n.get("branches") or []):
            return True
    return False


def touches(procs, pid):
    """その手続きと、そこからたどれる手続きが触るもの: 入出力・操作者とのやり取り・プログラムの終了・実行時に決まる名前。無ければ空（＝補助の候補）"""
    out = []
    for q in reach(procs, pid):
        p = procs[q]
        out += [f"{x['kind']}「{x['name']}」" for x in p.get("io") or []]
        out += ["操作者とのやり取り" for _x in p.get("interactions") or []]
        out += ["実行時に決まる名前" for _x in p.get("dynamic") or []]
        out += ["画面の部品への書き込み" for _x in p.get("ui") or []]
        if p.get("ends") or ends_program(p.get("body")) or any(ends_program(h.get("body")) for h in p.get("handlers") or []):
            out.append("プログラムの終了")
    return list(dict.fromkeys(out))


def helper_candidates(st):
    """補助の候補: 共通の処理（図にする単位）のうち、入出力・やり取り・プログラムの終了が、たどれる範囲に無いもの。
    図にしないと決めるのは、台帳（structure/systems.json の folded_procedures）。機械は、候補を出すだけ（画面の部品に書く手続きは、候補に入れない）"""
    procs = {p["id"]: p for p in st.get("procedures") or []}
    return [u["procedure"] for u in st.get("units") or [] if u.get("kind") == "shared" and u["procedure"] in procs and not touches(procs, u["procedure"])]


def strip_node(n):
    out = {k: v for k, v in n.items() if k not in ("text",) and v not in (None, [], "")}
    if "branches" in n:
        out["branches"] = [dict({k: v for k, v in b.items() if k != "body"}, body=[strip_node(x) for x in b["body"]]) for b in n["branches"]]
    if "body" in n:
        out["body"] = [strip_node(x) for x in n["body"]]
    return out


def access_entries(access, procs, entries):
    """Access のマクロの起動点: AutoExec（開いたとき）・AutoKeys（キーの割り当て）・どこからも呼ばれていないマクロ（ナビゲーションから実行できる）"""
    called = {c["to"] for p in procs.values() for c in p["calls"] if c.get("to")} | {e["procedure"] for e in entries}
    out = []
    for mac in access.macros.values():
        low = mac["module"].lower()
        for k, (sub, _rows) in enumerate(mac["parts"]):
            pid = f"{mac['module']}.{sub or '（マクロ）'}"
            if low == "autoexec" and k == 0:
                out.append({"procedure": pid, "trigger": "auto", "label": "データベースを開いたとき（AutoExec マクロ）", "evidence": "マクロの名前 AutoExec"})
            elif low == "autokeys" and sub:
                out.append({"procedure": pid, "trigger": "shortcut", "label": f"キー「{sub}」（AutoKeys マクロ）", "evidence": f"AutoKeys マクロのサブマクロ {sub}"})
            elif pid not in called and (k == 0 or sub):
                name = mac["module"] + (f".{sub}" if sub and k else "")
                out.append({"procedure": pid, "trigger": "manual", "label": f"マクロ「{name}」を実行（割り当ては見つからない）", "evidence": "ほかのマクロ・コード・フォームから呼ばれていない"})
    return out


def build_struct(doc_id, language, modules, sheets=None, buttons=None, shortcuts=(), docs=None, tree=None, ribbon=(), unread=()):
    """モジュールの並び [{"name", "kind", "lines", "pages", "sheet", "doc_id"}] → 構造。tree: 機械で読めない言語で、LLM が書いたブロックの木（code/<まとまり>.tree.json）"""
    access, read = None, {}
    if language in VBA_FAMILY:
        import extract_access
        access = extract_access.Access(modules)
        code = [m for m in modules if m.get("kind") not in extract_access.DEFS]
        mods = [VbaModule(m["name"], m.get("kind", "standard"), m["lines"], m.get("sheet"), access.code_from(m)) for m in code]
        ps = VbaParser(mods, sheets, buttons, access)
        procs, tier, embedded = ps.procs, "parsed", {}
        if access:
            resolve = lambda nm: ps.resolve(nm, "")
            extra, embedded = access.embedded_procs(code, resolve)
            extra.update(access.macro_procs(resolve))
            if extra and not procs and not all(x.get("read") == "axl" for x in extra.values()):
                tier = "heuristic"                                   # マクロだけ（アクションの行の形から読む。XML の定義のとおりに読めたものは、構文から）
            procs.update(extra)
        entries = vba_entries(ps, shortcuts, access, embedded, ribbon, unread)
        if access:
            entries += access_entries(access, procs, entries)
    else:
        import code_lang
        res = code_lang.parse(language, modules)                     # 機械で読めない言語・書き方は None（構造は LLM の読取）
        procs, entries, tier = res if res else ({}, [], "llm")
        if res is None:                                              # LLM が書いた木を確かめて、構造にする。誤りがあれば、手続きは空のまま（点検が FAIL にする）
            plain = {m["name"]: code_lang.plain_lines(language, m["lines"])[0] for m in modules}   # 空の行・注釈・かっこだけの行を除いた行
            read = {"tree": tree is not None, "total_lines": sum(len(v) for v in plain.values())}
            if tree is not None:
                errs = [tree["bad"]] if tree.get("bad") else code_lang.tree_problems(tree, modules, language)
                if errs:
                    read["errors"] = errs
                else:
                    procs, entries = code_lang.from_tree(tree, modules, language)
                    read["lines"] = sum(1 for p in procs.values() for ln in plain[p["module"]] if p["lines"][0] <= ln <= p["lines"][1])
                if str(tree.get("skip") or "").strip():
                    read["skip"] = str(tree["skip"]).strip()
    for k, e in enumerate(entries, 1):
        e["id"] = f"e{k}"
    mark_ends(procs)
    units, inline, unreached, blank = plan_units(procs, entries)
    refs = [r for r in access.references(procs) if r["procedure"] in unreached] if access else []   # クエリ・式の中から呼ばれる関数（起動点からは、たどれない）
    return {"doc_id": doc_id, "docs": docs or [doc_id], "language": language, "tier": tier,
            "modules": [{"name": m["name"], "kind": m.get("kind", "standard"), "doc_id": m.get("doc_id", doc_id), "pages": m.get("pages") or [], "lines": len(m["lines"]),
                         **({"sheet": m["sheet"]} if m.get("sheet") else {})} for m in modules],
            "procedures": [{"id": p["id"], "name": p["name"], "module": p["module"], "kind": p["kind"], "lines": p["lines"], "boxes": p["boxes"], "exec": p["exec"],
                            "body": [strip_node(n) for n in p["body"]],
                            "handlers": [{"label": h["label"], "lines": h["lines"], "body": [strip_node(n) for n in h["body"]]} for h in p["handlers"]],
                            "calls": p["calls"], "io": p["io"], "interactions": p["interactions"], "dynamic": p["dynamic"], **({"ui": p["ui"]} if p.get("ui") else {}),
                            **({"read": p["read"]} if p.get("read") else {}), **({"ends": True} if p.get("ends") else {})} for p in procs.values()],
            "entries": entries, "units": units, "inline": inline, "unreached": unreached, **({"blank": blank} if blank else {}),
            **({"referenced": refs} if refs else {}),
            **({"access": access.summary()} if access else {}), **({"read": read} if read else {})}


# ---------- 作業フォルダから ----------

def projects(m):
    """プログラムのまとまり（1つの構造として読む単位）: [{"id", "docs": [資料], "language"}]。
    Office のファイルのマクロは、そのファイルで1つ。ソースのファイルは、同じフォルダ・同じ言語のものを1つにまとめる（モジュールをまたぐ呼び出しを追うため）。
    id は、まとまりの先頭の資料の ID"""
    out, by, seen = [], {}, {}
    for d in m["docs"]:
        if d.get("duplicate_of"):
            continue
        ps = [p for p in d.get("page_list") or [] if p.get("kind") in CODE_KINDS and p.get("code")]
        if not ps:
            continue
        lang = ps[0]["code"].get("language", "vba")
        if d["kind"] != "code":
            out.append({"id": d["doc_id"], "docs": [d], "language": lang})
            continue
        fam = lang in VBA_FAMILY or bool(d.get("access"))            # Access のクエリの書き出し（SQL）も、同じフォルダのコードと1つにまとめる
        root = (d.get("access") or {}).get("root")                   # 版管理アドインの書き出しは、下のフォルダに分かれている（根のフォルダで1つにまとめる）
        key = (("root", root) if root else tuple(d.get("folder") or []), "vba" if fam else lang)
        if key not in by:
            by[key] = {"id": d["doc_id"], "docs": [], "language": lang}
            out.append(by[key])
            seen[key] = []
        by[key]["docs"].append(d)
        seen[key].append(lang)
        if fam:                                                      # まとまりの言語: コードがあれば VBA、マクロだけなら Access マクロ、定義だけなら Access の定義
            by[key]["language"] = next((x for x in ("vba", "vbscript", "accessmacro") if x in seen[key]), "accessdef")
    return out


def project_modules(work, proj):
    """まとまりのモジュール: [{"name", "kind", "doc_id", "lines", "pages": [{"page", "line_from", "lines", "doc_id"}], "sheet"}]。同じ名前が重なれば、資料 ID を添える。
    フォーム・レポートの定義と、コードが別のファイルのとき（版管理アドインの書き出し）は、定義のあとにコードをつないで、1つのモジュールにする"""
    mods = []
    for d in proj["docs"]:
        own = {}
        for p in d.get("page_list") or []:
            if p.get("kind") not in CODE_KINDS or not p.get("code"):   # 前の版でページ化したマクロのページ（カードにして読んだもの）は対象にしない
                continue
            pj = load_page(page_json(work, d["doc_id"], p["page"])) or {}
            c = pj.get("code") or {}
            key = c.get("module") or d["name"]
            mod = own.setdefault(key, {"name": key, "kind": c.get("module_kind", "standard"), "doc_id": d["doc_id"], "lines": [], "pages": [], "sheet": c.get("sheet"),
                                       "part": c.get("part")})
            mod["pages"].append({"page": p["page"], "line_from": len(mod["lines"]) + 1, "lines": len(pj.get("text") or []), "doc_id": d["doc_id"]})
            mod["lines"] += list(pj.get("text") or [])
        mods += list(own.values())
    for name in {x["name"] for x in mods if x.get("part") == "code"}:
        code = [x for x in mods if x["name"] == name and x.get("part") == "code"]
        lay = [x for x in mods if x["name"] == name and x.get("part") != "code" and x["kind"] in ("form", "report")]
        if len(code) == 1 and len(lay) == 1:
            shift = len(lay[0]["lines"])
            lay[0]["pages"] += [dict(pg, line_from=pg["line_from"] + shift) for pg in code[0]["pages"]]
            lay[0]["lines"] += code[0]["lines"]
            mods = [x for x in mods if x is not code[0]]
    names = [x["name"] for x in mods]
    for x in mods:
        if names.count(x["name"]) > 1:
            x["name"] = f"{x['name']}（{x['doc_id']}）"
    return mods


def cite(struct, module, a, b):
    """モジュールの行 [a, b] → 出典の並び [{"doc_id", "page", "lines": [ページの中の行]}]（ページをまたげば2つ以上）"""
    out = []
    m = next((x for x in struct["modules"] if x["name"] == module), None)
    for pg in (m or {}).get("pages") or []:
        lo, hi = max(a, pg["line_from"]), min(b, pg["line_from"] + pg["lines"] - 1)
        if lo <= hi:
            out.append({"doc_id": pg.get("doc_id") or m["doc_id"], "page": pg["page"], "lines": [lo - pg["line_from"] + 1, hi - pg["line_from"] + 1]})
    return out


def struct_of(work, proj):
    mods = project_modules(work, proj)
    d0 = proj["docs"][0]
    pick = lambda key, empty: next((d[key] for d in proj["docs"] if d.get(key)), empty)
    try:
        tree = load(pathlib.Path(work) / "code" / f"{proj['id']}.tree.json")
    except Exception as e:                                           # 書きかけ・閉じ忘れ（構造は作らず、誤りとして出す）
        tree = {"bad": str(e)}
    if isinstance(tree, dict) and not tree.get("bad"):
        from validate import validate
        errs = validate(tree, "code_tree")
        if errs:                                                     # 形の崩れた木からは、構造を作らない（references/schemas.md の「ブロックの木」）
            tree = {"bad": f"code/{proj['id']}.tree.json の形が崩れている: {'／'.join(errs[:5])}"}
    return build_struct(proj["id"], proj["language"], mods, sheets=pick("vba_sheets", {}), buttons=pick("vba_buttons", []), shortcuts=pick("vba_shortcuts", []),
                        docs=[d["doc_id"] for d in proj["docs"]], tree=tree if isinstance(tree, dict) else None,
                        ribbon=pick("vba_ribbon", []), unread=pick("vba_unread", [])) if mods and d0 else None


def tree_errors(st):
    """木の形の誤り: 行の範囲が手続きの外・親の外・前のノードと重なる・逆順。実行文が、どのノードにも入っていない・2つに入っている"""
    errs = []
    for p in st.get("procedures") or []:
        lo, hi = p["lines"]
        seen = {}

        def visit(nodes, a, b, where):
            last = a - 1
            for n in nodes:
                x, y = n.get("lines") or [0, 0]
                if not (a <= x <= y <= b):
                    errs.append(f"{p['id']}: {where}の {n.get('kind')}（{x}〜{y} 行）が、親の範囲（{a}〜{b} 行）の外")
                    continue
                if x < last and not n.get("single"):
                    errs.append(f"{p['id']}: {n.get('kind')}（{x}〜{y} 行）が、前のノードと重なっているか、順が逆")
                last = max(last, x)
                for ln in n.get("exec") or []:
                    seen[ln] = seen.get(ln, 0) + 1
                for br in n.get("branches") or []:
                    visit(br.get("body") or [], x, y, "分岐")
                visit(n.get("body") or [], x, y, "繰り返し")
        visit(p.get("body") or [], lo, hi, "手続き")
        for h in p.get("handlers") or []:
            visit(h.get("body") or [], lo, hi, "エラー処理")
        lost = [ln for ln in p.get("exec") or [] if ln not in seen]
        twice = [ln for ln, k in seen.items() if k > 1]
        if lost:
            errs.append(f"{p['id']}: 実行文の行が、木のどのノードにも入っていない（{lost[:6]}）")
        if twice:
            errs.append(f"{p['id']}: 同じ行が、木の2つのノードに入っている（{twice[:6]}）")
    return errs


def load_structs(work):
    """作業フォルダの構造（code/*.struct.json）: {まとまりの id: 構造}"""
    out = {}
    for f in sorted((pathlib.Path(work) / "code").glob("*.struct.json")):
        st = load(f)
        if isinstance(st, dict) and st.get("doc_id"):
            out[st["doc_id"]] = st
    return out


def show(st, out=print):
    out(f"== {st['doc_id']}（資料 {'・'.join(st.get('docs') or [st['doc_id']])}）  言語 {LANG_LABEL.get(st['language'], st['language'])}・構造の取り出し {TIER_LABEL[st['tier']]}・モジュール {len(st['modules'])}・"
        f"手続き {len(st['procedures'])}・起動点 {len(st['entries'])}・図にする単位 {len(st['units'])}")
    procs = {p["id"]: p for p in st["procedures"]}
    for e in st["entries"]:
        p = procs[e["procedure"]]
        out(f"  起動点 {e['id']}  {e['label']}  → {e['procedure']}（{p['lines'][0]}〜{p['lines'][1]} 行・箱 {p['boxes']}）  〔{TRIGGER_LABEL.get(e['trigger'], e['trigger'])}: {e['evidence']}〕")
    for u in st["units"]:
        if u["kind"] == "shared":
            out(f"  別ページにする手続き  {u['procedure']}（箱 {procs[u['procedure']]['boxes']}）")
    if st["inline"]:
        out("  呼ぶ側の図に入れる手続き: " + "・".join(st["inline"]))
    if st.get("blank"):
        out("  図にする文が無い手続き（図にしない。呼ぶ側では、呼び出しを対象外にする）: " + "・".join(st["blank"]))
    if st["unreached"]:
        out("  どの起動点からもたどれない手続き: " + "・".join(st["unreached"]))
    rd = st.get("read")
    if rd:
        out(f"  LLM の読取: 木 {'あり' if rd['tree'] else 'なし'}" + (f"・読んだ行 {rd.get('lines', 0)}/{rd['total_lines']}" if rd["tree"] else "")
            + (f"・読まない理由「{rd['skip']}」" if rd.get("skip") else "") + (f"・誤り {len(rd['errors'])} 件" if rd.get("errors") else ""))

    def tree(nodes, depth):
        for n in nodes:
            pad = "    " + "  " * depth
            ln = f"{n['lines'][0]}〜{n['lines'][1]}"
            if n["kind"] == "seq":
                io = "、".join(f"{x['kind']}「{x['name']}」{'に書く' if x['way'] == 'out' else 'を読む'}" for x in n.get("io") or [])
                out(f"{pad}[まとまり {ln}] {('（' + n['hint'] + '）') if n.get('hint') else ''}{io}" + (f"  ※{n['tech']}" if n.get("tech") else ""))
            elif n["kind"] in ("if", "select"):
                out(f"{pad}[分岐 {ln}] {n.get('cond', '')}" + ("  ※操作者に聞く" if n.get("interact") else ""))
                for b in n["branches"]:
                    out(f"{pad}  ・{b['label']} {b.get('cond', '')}")
                    tree(b["body"], depth + 2)
            elif n["kind"] == "loop":
                out(f"{pad}[繰り返し {ln}] {n.get('form', '')} {n.get('cond', '')}")
                tree(n["body"], depth + 1)
            elif n["kind"] == "call":
                out(f"{pad}[呼び出し {ln}] → {n.get('to')}")
            elif n["kind"] == "interact":
                out(f"{pad}[やり取り {ln}] {n['interact']['how']}「{n['interact'].get('text', '')}」")
            elif n["kind"] == "exit":
                out(f"{pad}[抜ける {ln}] {n['what']} {n.get('label') or ''}")
            elif n["kind"] == "onerror":
                out(f"{pad}[エラーの設定 {ln}] {n['mode']} {n.get('label') or ''}")
            elif n["kind"] == "label":
                out(f"{pad}[ラベル {ln}] {n['name']}")
    for p in st["procedures"]:
        out(f"  -- {p['id']}（{p['kind']}・{p['lines'][0]}〜{p['lines'][1]} 行・実行文 {len(p['exec'])} 行）")
        tree(p["body"], 0)
        for h in p["handlers"]:
            out(f"    エラーが起きたとき（{h['label']}）")
            tree(h["body"], 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work")
    ap.add_argument("--only", help="資料 ID（カンマ区切り）")
    ap.add_argument("--show", help="この資料の構造を表示する（資料 ID）")
    ap.add_argument("--file", nargs="+", help="ファイルを直接解析して表示する")
    ap.add_argument("--hints", action="store_true", help="機械で構文を読めない言語のまとまりについて、手続きの見出しらしい行を出す（LLM が読む範囲を決める手がかり）")
    ap.add_argument("--agree", nargs=2, metavar=("木1", "木2"), help="同じソースを2回読んだ木（tree.json）の一致を出す")
    a = ap.parse_args()
    if a.agree:
        import code_lang
        r = code_lang.tree_agreement(load(a.agree[0]), load(a.agree[1]))
        print(f"手続き {r['procedures'][0]}/{r['procedures'][1]}・ノード（種類と行の範囲。隣り合うまとまりは、つないで比べる）{r['nodes'][0]}/{r['nodes'][1]}＝{r['rate']:.1%}"
              f"・うち、分岐・繰り返し・呼び出し・やり取り・抜ける {r['flow'][0]}/{r['flow'][1]}・書いたままのノード {r['raw'][0]}/{r['raw'][1]}")
        for key, name in (("only_first", "1つ目だけ"), ("only_second", "2つ目だけ")):
            for pid, kind, span in r[key]:
                print(f"  {name}: {pid} {kind} {'〜'.join(map(str, span))} 行")
        return
    if a.file:
        mods = []
        langs = {LANG_OF_EXT.get(pathlib.Path(f).suffix.lower(), "vba") for f in a.file}
        if len(langs) > 1:
            sys.exit(f"--file: 言語の違うファイルは、別々に渡す（{'・'.join(sorted(langs))}）")
        for f in a.file:
            p = pathlib.Path(f)
            lines = p.read_text(encoding="utf-8-sig", errors="replace").splitlines()
            mods.append({"name": module_name(lines, p.stem), "kind": module_kind(module_name(lines, p.stem), p.suffix.lower()), "lines": lines})
        show(build_struct("FILE", LANG_OF_EXT.get(pathlib.Path(a.file[0]).suffix.lower(), "vba"), mods))
        return
    if not a.work:
        ap.error("-w か --file を指定する")
    work = work_dir(a.work)
    m = manifest(work)
    only = set((a.show or a.only or "").split(",")) - {""}
    n = 0
    keep = set()
    for proj in projects(m):
        keep.add(f"{proj['id']}.struct.json")
        if only and not only & {d["doc_id"] for d in proj["docs"]}:
            continue
        st = struct_of(work, proj)
        if st is None:
            continue
        n += 1
        if a.hints:
            if st["tier"] == "llm":
                import code_lang
                for mod in project_modules(work, proj):
                    print(f"== {proj['id']} {mod['name']}（{mod['doc_id']}・{len(mod['lines'])} 行）  手続きの見出しらしい行（手がかり。構造には使わない）")
                    for k, text in code_lang.hints(proj["language"], mod["lines"]):
                        print(f"  {k:>5}| {text}")
            continue
        if a.show:
            show(st)
            continue
        (work / "code").mkdir(exist_ok=True)
        save(work / "code" / f"{proj['id']}.struct.json", st)
        names = "・".join(d["rel_path"] for d in proj["docs"][:3]) + (f" ほか {len(proj['docs']) - 3} 件" if len(proj["docs"]) > 3 else "")
        print(f"{proj['id']} {names}: {LANG_LABEL.get(st['language'], st['language'])}・構造 {TIER_LABEL[st['tier']]}・手続き {len(st['procedures'])}・"
              f"起動点 {len(st['entries'])}・図にする単位 {len(st['units'])}" + (f"・たどれない手続き {len(st['unreached'])}" if st["unreached"] else ""))
        rd = st.get("read")
        if rd and not rd["tree"]:
            print(f"  ※ 機械で構文を読めない言語。LLM がソースを読んで、ブロックの木 code/{proj['id']}.tree.json を書く（references/code-guide.md。手がかりは --hints）")
        for e in (rd or {}).get("errors") or []:
            print(f"  ※ 木の誤り: {e}")
    for f in sorted((work / "code").glob("*.struct.json")) if not a.show and not only else []:   # まとまりが変わって残った、古い構造
        if f.name not in keep:
            f.unlink()
    if not n:
        print("コードのある資料が無い")


if __name__ == "__main__":
    main()
