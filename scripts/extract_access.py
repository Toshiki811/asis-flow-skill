#!/usr/bin/env python3
"""extract_access.py — Access を書き出したテキストを読む。Access のファイルそのものは開かない。
読むのは2つの並び: (1) assets/access_export/ の道具の出力（下の表）
(2) 版管理アドイン（msaccess-vcs-addin）の書き出し（<ファイル名>.src/ の forms・reports・macros・modules・queries・tbldefs）

  フォーム・レポート（F_*.form.txt・R_*.report.txt）  画面の定義（コントロールと、イベントに割り当てたもの）＋ コード（CodeBehindForm のあと）
  マクロ（X_*.macro.txt）                              アクションの並び（条件・呼び出し・クエリ・フォーム・ファイル）
  モジュール（M_*.bas）                                VBA（code_struct.py が読む）
  クエリ（access_queries.sql）                         名前・種類・SQL（読むテーブル・書くテーブル）
  テーブル（access_tables.txt）                        項目の名前と型・リンク先・リレーション（行の中身は無い）
  書き出しの一覧（access_export.txt）                  書き出したもの・書き出せなかったもの

  python3 scripts/extract_access.py <書き出したフォルダ>      # 読めた中身を表示する

フォーム・レポート・マクロの形は、Access の SaveAsText の出力（ブロックは Begin … End、項目は「名前 =値」）。
マクロは、条件のブロックを含む定義が「_AXL:」で始まる行に XML で入っているので、あればそれを読む。
形は、版管理アドインの実物の書き出しで確かめた（references/access-export.md）。読めなかった所は、読めなかったと出す。
"""
import functools
import json
import pathlib
import re
import urllib.parse
import xml.etree.ElementTree as ET

MARK = re.compile(r"^\s*CodeBehind(Form|Report)\s*$")
MARK_LAYOUT, MARK_GUESS = "AsisFlowLayout", "AsisFlowGuess"          # 直接読み（access_direct.py）が作ったテキストだけにある印
SUFFIX = ((".form.txt", "form"), (".report.txt", "report"), (".macro.txt", "macro"), (".form", "form"), (".report", "report"), (".mac", "macro"), (".macro", "macro"))
NAMES = {"access_export.txt": "log", "access_tables.txt": "tables", "access_queries.sql": "queries"}
PREFIX = {"form": "F_", "report": "R_", "macro": "X_", "module": "M_"}
LANGUAGE = {"form": "vba", "report": "vba", "module": "vba", "macro": "accessmacro", "queries": "sql", "tables": "accessdef", "log": "accessdef",
            "query": "sql", "tabledef": "accessdef", "props": "accessdef"}
DEFS = ("macro", "queries", "tables", "log", "query", "tabledef", "props")   # VBA のモジュールとしては読まないもの
VCS_MARKS = ("vcs-options.json", "vbe-project.json")                 # 版管理アドインで書き出したフォルダの印
SQL_TYPE = {"autoincrement": "autonumber", "counter": "autonumber", "varchar": "text", "char": "text", "text": "text", "longtext": "memo", "memo": "memo",
            "long": "long", "integer": "long", "int": "long", "short": "integer", "smallint": "integer", "byte": "byte", "bit": "yesno", "yesno": "yesno",
            "datetime": "datetime", "date": "datetime", "currency": "currency", "money": "currency", "double": "double", "float": "double", "single": "single",
            "real": "single", "longbinary": "ole", "binary": "binary", "varbinary": "binary", "guid": "guid", "uniqueidentifier": "guid", "decimal": "decimal",
            "numeric": "decimal", "bigint": "bigint"}
TMP_NAME = "_access_export_tmp.accdb"                                # 書き出しの道具が作って消す、仮のデータベース（残っていても読まない）
SECTIONS = {"section", "formheader", "formfooter", "pageheader", "pagefooter", "breakheader", "breakfooter"}
EVENT_PROP = re.compile(r"^(On[A-Z]\w*|Before[A-Z]\w*|After[A-Z]\w*)$")
PROP = re.compile(r"^\s*([A-Za-z_]\w*)\s*=\s?(.*)$")
CONTROL_JA = {"commandbutton": "ボタン", "togglebutton": "ボタン", "textbox": "", "combobox": "", "listbox": "", "checkbox": "", "optionbutton": "", "optiongroup": "",
              "label": "", "subform": "サブフォーム", "tab": "タブ", "page": "ページ", "image": "画像"}
TOP_EVENT_JA = {"open": "を開いたとき", "load": "を開いたとき", "current": "で、表示するレコードが変わったとき", "close": "を閉じるとき", "unload": "を閉じるとき",
                "beforeupdate": "で、レコードを保存する前", "afterupdate": "で、レコードを保存したあと", "beforeinsert": "で、新しいレコードを入れはじめたとき",
                "afterinsert": "で、新しいレコードを保存したあと", "delete": "で、レコードを削除するとき", "timer": "を開いているあいだ、一定の時間ごと",
                "activate": "が前面に来たとき", "nodata": "で、データが無いとき", "dirty": "で、レコードを書き換えはじめたとき", "error": "で、エラーが起きたとき",
                "page": "で、ページを作るとき"}
SECTION_EVENT_JA = {"format": "の書式を整えるとき", "print": "を印刷するとき", "click": "をクリック", "dblclick": "をダブルクリック"}
CTL_EVENT_JA = {"click": "をクリック", "dblclick": "をダブルクリック", "change": "の値を変えたとき", "afterupdate": "の値を変えたあと", "beforeupdate": "の値を変える前",
                "enter": "に入ったとき", "exit": "から出たとき", "notinlist": "に、一覧に無い値を入れたとき", "gotfocus": "に入ったとき", "lostfocus": "から出たとき"}
QUERY_JA = {"select": "選択", "crosstab": "クロス集計", "delete": "削除", "update": "更新", "append": "追加", "maketable": "テーブル作成", "ddl": "データ定義",
            "passthrough": "パススルー", "union": "ユニオン"}
ACTION_QUERY = ("delete", "update", "append", "maketable", "ddl")
ID = r"(?:\[([^\]]+)\]|([\w\u0080-\uffff]+(?:\.[\w\u0080-\uffff]+)*))"
# マクロのアクション（前の名前と、いまの名前）→ 扱い
ACTIONS = {"runcode": "code", "runmacro": "macro", "openquery": "query", "runsql": "sql", "openform": "form", "openreport": "report", "opentable": "table",
           "transferspreadsheet": "sheet", "importexportspreadsheet": "sheet", "transfertext": "text", "importexporttext": "text", "outputto": "output",
           "exportwithformatting": "output", "sendobject": "mail", "emaildatabaseobject": "mail", "msgbox": "msgbox", "messagebox": "msgbox",
           "stopmacro": "stop", "stopallmacros": "stop", "setwarnings": "tech", "echo": "tech", "hourglass": "tech",
           "displayhourglasspointer": "tech", "runapp": "app", "runapplication": "app"}


# ---------- ファイルの種類 ----------

def read_head(path, size=4000):
    raw = pathlib.Path(path).open("rb").read(size)
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw[: len(raw) // 2 * 2].decode("utf-16", "replace")
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def sniff(lines):
    """先頭の数行から: form／report／macro／tables／queries／log。どれでもなければ None"""
    head = [ln.strip() for ln in lines[:12] if ln.strip()]
    if not head:
        return None
    first = head[0].lstrip("﻿")
    if first.startswith("# access_export "):
        return "log"
    if first.startswith("# access_tables "):
        return "tables"
    if first.startswith("-- access_queries "):
        return "queries"
    if re.match(r"^Version\s*=\s*\d+$", first):
        for ln in head[1:]:
            mt = re.match(r"(?i)^Begin\s+(Form|Report)$", ln)
            if mt:
                return mt.group(1).lower()
        if any(re.match(r"^ColumnsShown\s*=", ln) for ln in head[1:]) or any(re.match(r"^(Action|MacroName|Condition)\s*=", ln) for ln in head[1:]):
            return "macro"
    return None


@functools.lru_cache(maxsize=1024)
def _vcs_root(folder):
    p = pathlib.Path(folder)
    for q in [p, *p.parents][:5]:
        if any((q / x).is_file() for x in VCS_MARKS):
            return str(q)
    return None


def vcs_root(path):
    """版管理アドインで書き出したフォルダの根（vcs-options.json か vbe-project.json のあるフォルダ）。その中のファイルでなければ None"""
    r = _vcs_root(str(pathlib.Path(path).parent))
    return pathlib.Path(r) if r else None


def vcs_role(path, root):
    """アドインの書き出しの中での役割: form／formcode／report／reportcode／macro／module／query／tabledef／log／props。
    読まないファイル（画像・テーマ・メニュー・テーブルの行の中身・設定 など）は None"""
    p = pathlib.Path(path)
    parts = [x.lower() for x in p.relative_to(root).parts]
    suf = p.suffix.lower()
    if len(parts) == 1:
        return {"vbe-project.json": "log", "dbs-properties.json": "props"}.get(parts[0])
    top = parts[0]
    if top in ("forms", "reports"):
        kind = top[:-1]
        return kind if suf in (f".{kind}", ".bas") else f"{kind}code" if suf == ".cls" else None
    if top == "macros":
        return "macro" if suf in (".macro", ".bas") else None
    if top == "modules":
        return "module" if suf in (".bas", ".cls") else None
    if top == "vbeforms":
        return "module" if suf == ".frm" else None
    if top == "queries":
        return "query" if suf == ".sql" else None
    if top == "tbldefs":
        return "tabledef" if suf in (".sql", ".json") else None
    return None


def skip_in_export(path):
    """アドインの書き出しのフォルダの中の、読まないファイルか（棚卸しで、資料に数えない）"""
    root = vcs_root(path)
    return root is not None and vcs_role(path, root) is None


def layout_file(path):
    """フォーム・レポートのコードのファイル（アドインの書き出しの <名前>.cls）→ 同じ名前の定義のファイル。無ければ None"""
    p = pathlib.Path(path)
    for suf in (".form", ".report", ".bas"):
        q = p.with_suffix(suf)
        if q.is_file() and q != p:
            return q
    return None


def by_name(name):
    low = name.lower()
    if low in NAMES:
        return NAMES[low]
    return next((k for suf, k in SUFFIX if low.endswith(suf)), None)


def is_export(path):
    """Access を書き出したテキストか（棚卸しで、プログラムの資料に数えるかを決める）。名前で分からなければ、.txt の先頭を見る"""
    p = pathlib.Path(path)
    root = vcs_root(p)
    if root is not None:
        return vcs_role(p, root) is not None
    if by_name(p.name):
        return True
    if p.suffix.lower() != ".txt":
        return False
    try:
        return sniff(read_head(p).splitlines()) is not None
    except OSError:
        return False


@functools.lru_cache(maxsize=64)
def log_index(folder):
    """書き出しの一覧（access_export.txt）: {"source", "access", "files": {ファイル名（小文字）: {"kind", "name", "status"}}, "failed": [{"kind", "name", "why"}]}。無ければ None"""
    f = pathlib.Path(folder) / "access_export.txt"
    if not f.is_file():
        return None
    try:
        return parse_log(read_head(f, 4_000_000).splitlines())
    except OSError:
        return None


def load_json(lines):
    try:
        v = json.loads("\n".join(lines).lstrip("\ufeff"))
        return v if isinstance(v, dict) else {}
    except ValueError:
        return {}


def parse_log(lines):
    out = {"source": "", "access": "", "files": {}, "failed": [], "error": "", "via": "", "notes": []}
    if lines and lines[0].lstrip("\ufeff").lstrip().startswith("{"):   # アドインの書き出し（vbe-project.json）: 元のファイルの名前だけが分かる
        out["source"] = str((load_json(lines).get("Items") or {}).get("FileName") or "")
        return out
    for ln in lines:
        c = ln.rstrip("\r\n").split("\t")
        if c[0] in ("source", "access", "via") and len(c) > 1:       # via direct: Access のファイルを直接読んで作ったテキスト（access_direct.py）
            out[c[0]] = c[1]
        elif c[0] == "note" and len(c) > 1:
            out["notes"].append(c[1])
        elif c[0] == "error":
            out["error"] = c[-1]
        elif c[0] in ("form", "report", "macro", "module") and len(c) >= 4:
            out["files"][c[2].lower()] = {"kind": c[0], "name": c[1], "status": c[3]}
            if not c[3].startswith("ok"):
                out["failed"].append({"kind": c[0], "name": c[1], "why": c[3]})
    return out


def describe(path, lines):
    """書き出したファイルの素性: {"access", "name", "language", "module", "module_kind"}。Access の書き出しでなければ None"""
    p = pathlib.Path(path)
    root = vcs_root(p)
    if root is not None:
        return describe_vcs(p, lines, root)
    idx = log_index(str(p.parent))
    rec = (idx or {"files": {}})["files"].get(p.name.lower())
    kind = by_name(p.name) or sniff(lines) or (rec or {}).get("kind")
    if kind is None or (kind == "module" and p.suffix.lower() != ".bas"):
        return None
    stem = p.name
    for suf, k in SUFFIX:
        if k == kind and stem.lower().endswith(suf):
            stem = stem[: -len(suf)]
            break
    else:
        stem = p.stem
    if rec:
        name = rec["name"]
    else:
        name = stem[2:] if kind in PREFIX and stem.startswith(PREFIX[kind]) and idx is not None else stem
    out = {"access": kind, "name": name, "language": LANGUAGE[kind], "module": name, "module_kind": kind}
    if kind in ("form", "report"):
        lay = parse_layout(lines)
        tail = lines[lay["code_from"] - 1:lay["code_from"] + 20] if lay and lay["code_from"] else []
        given = next((mt.group(1) for mt in (re.match(r'(?i)^Attribute\s+VB_Name\s*=\s*"([^"]+)"', x) for x in tail) if mt), None)
        out["module"] = given or f"{'Form' if kind == 'form' else 'Report'}_{name}"
    elif kind == "module":
        klass = any(re.match(r"(?i)^Attribute\s+VB_(Creatable|PredeclaredId|GlobalNameSpace)\b", x) for x in lines[:12])
        out["module_kind"] = "class" if klass else "standard"
    elif kind in ("queries", "tables", "log"):
        out["module"] = p.stem
    return out


def describe_vcs(p, lines, root):
    """アドインの書き出しのファイルの素性（describe と同じ形に、root 書き出しの根・part "code" コードだけのファイル、を足す）"""
    role = vcs_role(p, root)
    if role is None:
        return None
    stem = urllib.parse.unquote(p.stem)                              # 名前に使えない字は、%XX で書いてある
    given = next((mt.group(1) for mt in (re.match(r'(?i)^Attribute\s+VB_Name\s*=\s*"([^"]+)"', x) for x in lines[:14]) if mt), None)
    if role in ("form", "report", "formcode", "reportcode"):
        kind = role[:-4] if role.endswith("code") else role
        if not role.endswith("code") and sniff(lines) != kind:
            return None
        out = {"access": kind, "name": stem, "language": "vba", "module": f"{'Form' if kind == 'form' else 'Report'}_{stem}", "module_kind": kind}
        if role.endswith("code"):
            out["part"] = "code"
    elif role == "macro":
        if sniff(lines) != "macro":
            return None
        out = {"access": "macro", "name": stem, "language": "accessmacro", "module": stem, "module_kind": "macro"}
    elif role == "module":
        klass = any(re.match(r"(?i)^Attribute\s+VB_(Creatable|PredeclaredId|GlobalNameSpace)\b", x) for x in lines[:14])
        out = {"access": "module", "name": given or stem, "language": "vba", "module": given or stem,
               "module_kind": "form" if p.suffix.lower() == ".frm" else "class" if klass else "standard"}
    elif role == "query":
        out = {"access": "query", "name": stem, "language": "sql", "module": f"クエリ {stem}", "module_kind": "query"}
    elif role == "tabledef":
        out = {"access": "tabledef", "name": stem, "language": "accessdef", "module": f"テーブル {stem}", "module_kind": "tabledef"}
    else:
        out = {"access": role, "name": p.stem, "language": "accessdef", "module": p.stem, "module_kind": role}
    out["root"] = str(root)
    return out


# ---------- フォーム・レポートの定義 ----------

def unquote(val):
    """「"…"」→ 中身（\\" と、改行の \\015\\012 を戻す）。文字でなければ、そのまま"""
    v = val.strip()
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        v = v[1:-1].replace("\\015\\012", "\n")
        v = re.sub(r'\\(\d{3}|.)', lambda mt: chr(int(mt.group(1), 8)) if mt.group(1).isdigit() else mt.group(1), v)   # \" \\ と、8 進の字（\011 など）
    return v


def vb_ident(name):
    return re.sub(r"[^\w\u0080-￿]", "_", name)


def parse_layout(lines):
    """フォーム・レポートの定義 → {"object", "caption", "record_source", "code_from", "controls", "events", "embedded"}。その形でなければ None。
    controls: [{"name", "type", "caption", "label", "line"}]／events: [{"owner", "owner_kind", "prop", "event", "kind", "value", "line", "handler"}]
    （kind: code コードの手続き／macro マクロ／expr 関数の式／embedded 埋め込みマクロ）／embedded: {"所有者.項目": [開始の行, 終了の行]}／code_from: コードの最初の行（無ければ None）"""
    obj = next((mt.group(1).lower() for mt in (re.match(r"(?i)^\s*Begin\s+(Form|Report)\s*$", ln) for ln in lines[:12]) if mt), None)
    if obj is None:
        return None
    out = {"object": obj, "caption": "", "record_source": "", "code_from": None, "controls": [], "events": [], "embedded": {}, "exprs": [], "unread": ""}
    stack, deep, raw_from, raw_prop, last = [], 0, 0, "", None

    def owner_of(skip=0):
        typed = [b for b in stack if b["type"]]
        return typed[-1 - skip] if len(typed) > skip else None

    def close(blk):
        low = blk["type"].lower()
        props = blk["props"]
        name = props.get("Name", ("", 0))[0]
        guess = {x.lower() for x in props.get(MARK_GUESS, ("", 0))[0].split(";") if x}   # 直接読み: 種類の番号が分からず、手続きの名前から決めたイベント
        if low in ("form", "report"):
            kind, who = "top", obj.capitalize()
            out["caption"], out["record_source"] = props.get("Caption", ("", 0))[0], props.get("RecordSource", ("", 0))[0]
            out["unread"] = re.sub(r"^unread:\s*", "", props.get(MARK_LAYOUT, ("", 0))[0])   # 直接読み: 定義を読めていない（理由）
        elif low in SECTIONS:
            kind, who = "section", name
        elif name:
            kind, who = "control", name
            ctl = {"name": name, "type": blk["type"], "caption": props.get("Caption", ("", 0))[0], "label": blk.get("label", ""), "line": blk["line"]}
            out["controls"].append(ctl)
            par = owner_of()
            if low == "label" and par is not None and par["type"].lower() not in SECTIONS | {"form", "report"}:
                par["label"] = ctl["caption"]                        # コントロールに付いたラベル（項目の見出し）
        else:
            return
        for prop, (val, at) in props.items():
            if val.startswith("=") and not EVENT_PROP.match(prop):   # コントロールの元・既定値などの式（VBA の関数を呼んでいることがある）
                out["exprs"].append(val)
            if prop == "RecordSource" and val:
                out["exprs"].append(val)
            if not EVENT_PROP.match(prop) or not val:
                continue
            ev = prop[2:] if prop.startswith("On") else prop
            how = "code" if val == "[Event Procedure]" else "embedded" if val == "[Embedded Macro]" else "expr" if val.startswith("=") else "macro"
            out["events"].append({"owner": who, "owner_kind": kind, "prop": prop, "event": ev, "kind": how, "value": val, "line": at,
                                  "handler": f"{vb_ident(who)}_{ev}", **({"guess": True} if ev.lower() in guess else {})})
        for prop, span in blk["raw"].items():
            if prop.endswith("EmMacro"):
                out["embedded"][f"{who}.{prop[:-7]}"] = span

    for k, ln in enumerate(lines, 1):
        s = ln.strip()
        if deep:                                                     # 16 進の値・埋め込みマクロの中身（Begin … End が入れ子になる）
            if re.match(r"^Begin(\s|$)", s) or re.search(r"=\s*Begin$", s):
                deep += 1
            elif s == "End":
                deep -= 1
                if not deep and stack:
                    stack[-1]["raw"][raw_prop] = [raw_from, k]
            continue
        if MARK.match(ln):
            out["code_from"] = k + 1 if k < len(lines) else None
            break
        mt = re.match(r"^([A-Za-z_]\w*)\s*=\s*Begin$", s)
        if mt:
            deep, raw_from, raw_prop, last = 1, k, mt.group(1), None
            continue
        mt = re.match(r"^Begin(?:\s+(\w+))?$", s)
        if mt:
            stack.append({"type": mt.group(1) or "", "line": k, "props": {}, "raw": {}})
            last = None
            continue
        if s == "End":
            if stack:
                close(stack.pop())
            last = None
            continue
        if s.startswith('"') and last and stack:                     # 長い文字の続き
            val, at = stack[-1]["props"][last]
            stack[-1]["props"][last] = ((val + unquote(s)).replace("\r\n", "\n"), at)   # 改行の2字（\015\012）の間で、行が切れていることがある
            continue
        mt = PROP.match(ln)
        if mt and stack:
            last = mt.group(1) if mt.group(2).strip().startswith('"') else None
            stack[-1]["props"][mt.group(1)] = (unquote(mt.group(2)), k)
    while stack:                                                     # 閉じていないブロック（途中で切れた書き出し）
        close(stack.pop())
    out["events"].sort(key=lambda e: e["line"])
    out["controls"].sort(key=lambda c: c["line"])
    return out


def event_label(lay, name, ev):
    """起動点の名前: フォーム「受注入力」の「登録」ボタンをクリック"""
    who = f"{'フォーム' if lay['object'] == 'form' else 'レポート'}「{name}」"
    low = ev["event"].lower()
    if low.startswith("unknown"):                                    # 直接読み: 何かのイベントに割り当ててあるが、種類は読めていない
        part = "" if ev["owner_kind"] == "top" else f"の「{ev['owner']}」"
        return f"{who}{part}のイベント（種類は読めていない）"
    if ev["owner_kind"] == "top":
        return who + TOP_EVENT_JA.get(low, f"のイベント {ev['event']}")
    if ev["owner_kind"] == "section":
        return f"{who}の「{ev['owner']}」" + SECTION_EVENT_JA.get(low, f"のイベント {ev['event']}")
    ctl = next((c for c in lay["controls"] if c["name"] == ev["owner"]), {"type": "", "caption": "", "label": ""})
    shown = (ctl["caption"] or "").strip() or (ctl["label"] or "").strip() or ev["owner"]
    return f"{who}の「{shown}」{CONTROL_JA.get(ctl['type'].lower(), '')}" + CTL_EVENT_JA.get(low, f"のイベント {ev['event']}")


# ---------- マクロ ----------

def macro_rows(lines, a=1, b=None):
    """マクロの行（Begin … End の1組）: ([{"lines": [開始, 終了], "props": [(名前, 値, 行)]}], _AXL の行があるか)"""
    rows, cur, axl, last = [], None, False, None
    for k in range(a, (b or len(lines)) + 1):
        s = lines[k - 1].strip()
        if s == "Begin":
            cur, last = {"lines": [k, k], "props": []}, None
        elif s == "End" and cur is not None:
            cur["lines"][1] = k
            rows.append(cur)
            cur = None
        elif cur is not None:
            if s.startswith('"') and last is not None:
                nm, val, at = cur["props"][last]
                cur["props"][last] = (nm, val + unquote(s), at)
                continue
            mt = PROP.match(s)
            if mt:
                cur["props"].append((mt.group(1), unquote(mt.group(2)), k))
                last = len(cur["props"]) - 1
    for r in rows:
        if any(nm == "Comment" and val.startswith("_AXL:") for nm, val, _ in r["props"]):
            r["axl"], axl = True, True
    return rows, axl


def local(el):
    return el.tag.rsplit("}", 1)[-1]


def axl_root(rows):
    """マクロの行の中の「_AXL:」の行（XML を、いくつかの行に切って入れてある）→ XML の根。無い・壊れていれば None"""
    text = "".join(val[5:] for r in rows if r.get("axl") for nm, val, _k in r["props"] if nm == "Comment" and val.startswith("_AXL:"))
    body = re.sub(r"^\s*<\?xml[^>]*\?>\s*", "", text)
    if not body.strip():
        return None
    try:
        return ET.fromstring(body)
    except ET.ParseError:
        return None


def axl_statements(root, sub=""):
    """XML の根 → サブマクロ sub（無ければ本体）の文の並びの要素。無ければ None"""
    if root is None:
        return None
    if local(root) != "UserInterfaceMacro":
        root = next((el for el in root.iter() if local(el) == "UserInterfaceMacro"), root)
    if sub:
        holder = next((el for el in root if local(el) == "Sub" and (el.get("Name") or "").lower() == sub.lower()), None)
    else:
        holder = root
    return next((el for el in (holder if holder is not None else []) if local(el) == "Statements"), None)


AXL_FIRST = {"runcode": "FunctionName", "runmacro": "MacroName", "openquery": "QueryName", "runsql": "SQLStatement", "opentable": "TableName",
             "openform": "FormName", "openreport": "ReportName", "messagebox": "Message", "msgbox": "Message"}   # 最初の引数の名前（XML では、引数に名前が付く）


def axl_where(root, rows):
    """XML の定義だけのマクロ（直接読みのテキスト。アクション・条件が、1行に1つずつ入っている）: 文の要素 → その行。その形でなければ None"""
    marks = []
    for r in rows:
        text = next((val[5:] for nm, val, _k in r["props"] if nm == "Comment" and val.startswith("_AXL:")), None)
        mt = re.match(r"\s*<(?:\w+:)?(Action|ElseIf|If|Else)\b", text or "")
        if mt:
            marks.append((mt.group(1), r["lines"]))
    els = [el for el in root.iter() if local(el) in ("Action", "If", "ElseIf", "Else")]
    if not els or len(els) != len(marks) or any(local(el) != tag for el, (tag, _ln) in zip(els, marks)):
        return None
    return {id(el): lines for el, (_tag, lines) in zip(els, marks)}


def axl_parts(root):
    """XML の定義 → サブマクロの名前の並び（本体は ""）"""
    top = root if local(root) == "UserInterfaceMacro" else next((el for el in root.iter() if local(el) == "UserInterfaceMacro"), root)
    main = next((el for el in top if local(el) == "Statements"), None)
    names = [""] if main is not None and len(main) else []
    return names + [el.get("Name") or "" for el in top if local(el) == "Sub" and el.get("Name")]


def is_helper(r):
    """条件のブロックを、前の形（条件の列）で表すために、Access が足した行（SetLocalVar __*L0_ など）"""
    get = [val for nm, val, _k in r["props"] if nm in ("Action", "Argument")]
    return len(get) >= 2 and get[0].lower() == "setlocalvar" and get[1].startswith("__*L")


def split_macro(rows):
    """マクロの行を、サブマクロごとに分ける: [(サブマクロの名前（無ければ ""）, [行])]"""
    parts = []
    for r in rows:
        if r.get("axl"):
            continue
        sub = next((val for nm, val, _ in r["props"] if nm == "MacroName" and val), None)
        if sub is not None or not parts:
            parts.append((sub or "", []))
        if any(nm == "Action" for nm, _v, _k in r["props"]):
            parts[-1][1].append(r)
    return [(nm, rs) for nm, rs in parts if rs]


# ---------- クエリ・テーブル ----------

SQL_ID = r"(?:\[([^\]]+)\]|[\"`]([^\"`]+)[\"`]|([\w#\u0080-\uffff]+(?:\.[\w#\u0080-\uffff]+)*))"   # [角かっこ]・"引用符"・`逆引用符`・素の名前（schema.table・#一時表）
SQL_WRITE = (r"(?:INSERT\s+(?:INTO\s+)?|UPDATE\s+|DELETE\s+FROM\s+|MERGE\s+INTO\s+|TRUNCATE\s+TABLE\s+|CREATE\s+(?:TEMP(?:ORARY)?\s+)?TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?|"
             r"DROP\s+TABLE\s+(?:IF\s+EXISTS\s+)?|ALTER\s+TABLE\s+|INTO\s+)")
SQL_WORDS = {"select", "set", "values", "where", "table", "dual"}


def sql_tables(sql):
    """SQL の文字 → (書くテーブル, 読むテーブル・クエリ)。角かっこ・引用符の名前（空白を含む名前）、カンマで並べた表も拾う。
    Access のクエリ・マクロ（ここ）、VBA の文字の中の SQL（code_struct.py）、SQL のファイル・Python の中の SQL（code_lang.py）が、同じこれで読む"""
    name = lambda mt, i=1: mt.group(i) or mt.group(i + 1) or mt.group(i + 2)
    outs = [name(mt) for mt in re.finditer(rf"(?i)\b{SQL_WRITE}{SQL_ID}", sql)]
    wipe = re.match(r"(?i)\s*(DELETE|TRUNCATE|DROP)\b", sql)
    if re.match(r"(?i)\s*DELETE\b", sql):                            # DELETE 表.* FROM 表（Access）: 最初の FROM の表が消す先
        outs += [name(mt) for mt in re.finditer(rf"(?i)\bFROM\s+\(*\s*{SQL_ID}", sql)][:1]
    ins = []
    for mt in re.finditer(rf"(?i)\b(?:FROM|JOIN)\s+\(*\s*{SQL_ID}((?:\s*(?:AS\s+[\w\u0080-\uffff]+)?\s*,\s*{SQL_ID})*)", sql):
        ins.append(name(mt))
        ins += [name(x) for x in re.finditer(rf",\s*{SQL_ID}", mt.group(4) or "")]
    if re.match(r"(?i)\s*UPDATE\b", sql):                            # UPDATE 表 INNER JOIN 表 …: 先頭の表が書く先
        ins = [t for t in ins if t not in outs]
    outs = [t for t in dict.fromkeys(outs) if t.lower() not in SQL_WORDS]
    ins = [t for t in dict.fromkeys(ins) if t.lower() not in SQL_WORDS and (t not in outs or not wipe)]
    return outs, ins


def parse_queries(lines):
    """access_queries.sql → [{"name", "type", "connect", "lines": [見出しの行, 終わりの行], "sql", "writes", "reads"}]"""
    out, cur = [], None

    def close():
        if cur is None:
            return
        while cur["body"] and not cur["body"][-1][1].strip():
            cur["body"].pop()
        cur["lines"] = [cur["lines"][0], cur["body"][-1][0] if cur["body"] else cur["lines"][0]]
        cur["sql"] = " ".join(t.strip() for _k, t in cur["body"] if t.strip())
        cur["writes"], cur["reads"] = sql_tables(cur["sql"])
        del cur["body"]
        out.append(cur)
    for k, ln in enumerate(lines, 1):
        mt = re.match(r"^-- query: (.+?)(?:\ttype=(\S+))?\s*$", ln)
        if mt:
            close()
            cur = {"name": mt.group(1).strip(), "type": mt.group(2) or "", "connect": "", "lines": [k, k], "body": []}
        elif cur is not None:
            mc = re.match(r"^-- connect: (.*)$", ln)
            if mc:
                cur["connect"] = mc.group(1).strip()
            else:
                cur["body"].append((k, ln))
    close()
    return out


def parse_tables(lines):
    """access_tables.txt → {"startup_form", "tables": [{"name", "linked", "source", "connect", "rows", "fields": [{"name", "type", "pk"}], "line"}], "relations": [[元, 先]]}"""
    out, cur = {"startup_form": "", "tables": [], "relations": []}, None
    for k, ln in enumerate(lines, 1):
        c = ln.rstrip("\r\n").split("\t")
        if c[0] == "[startup]":
            out["startup_form"] = next((x[5:] for x in c[1:] if x.startswith("form=")), "")
        elif c[0] in ("[table]", "[linked]") and len(c) > 1:
            kv = dict(x.split("=", 1) for x in c[2:] if "=" in x)
            cur = {"name": c[1], "linked": c[0] == "[linked]", "source": kv.get("source", ""), "connect": kv.get("connect", ""),
                   "rows": int(kv["rows"]) if kv.get("rows", "").lstrip("-").isdigit() else None, "fields": [], "line": k}
            out["tables"].append(cur)
        elif c[0] == "[relation]" and len(c) > 2:
            out["relations"].append([c[1], c[2]])
        elif ln.startswith("  ") and cur is not None and c[0].strip():
            cur["fields"].append({"name": c[0].strip(), "type": c[1] if len(c) > 1 else "", "pk": "pk" in c[2:]})
    return out


def sql_type(sql):
    """SQL の先頭の語から、クエリの種類（アドインの書き出しには、種類の行が無い）"""
    head = re.sub(r"(?is)^\s*PARAMETERS\b.*?;\s*", "", sql).lstrip("( \t")
    mt = re.match(r"(?i)(SELECT|TRANSFORM|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP)\b", head)
    word = mt.group(1).lower() if mt else ""
    if word == "select":
        before = re.split(r"(?i)\bFROM\b", head, 1)[0]
        return "maketable" if re.search(r"(?i)\bINTO\b", before) else "union" if re.search(r"(?i)\bUNION\b", head) else "select"
    return {"transform": "crosstab", "insert": "append", "update": "update", "delete": "delete", "create": "ddl", "alter": "ddl", "drop": "ddl"}.get(word, "")


def parse_tabledef(lines):
    """アドインの書き出しのテーブルの定義（CREATE TABLE の SQL か、リンク テーブルの JSON）→ parse_tables の tables の1件。読めなければ None"""
    first = next((ln.lstrip("\ufeff").strip() for ln in lines if ln.strip()), "")
    if first.startswith("{"):
        it = load_json(lines).get("Items") or {}
        if not it.get("Name"):
            return None
        return {"name": str(it["Name"]), "linked": True, "source": str(it.get("SourceTableName") or ""), "connect": str(it.get("Connect") or ""),
                "rows": None, "fields": [], "line": 1}
    text = "\n".join(lines)
    mt = re.search(rf"(?i)CREATE\s+TABLE\s+{ID}", text)
    if not mt:
        return None
    pk = re.search(r"(?i)PRIMARY\s+KEY\s*\(([^)]*)\)", text)
    keys = [x.strip().strip("[]").lower() for x in pk.group(1).split(",")] if pk else []
    fields = []
    for k, ln in enumerate(lines, 1):
        f = re.match(r"^\s*\[([^\]]+)\]\s+([A-Za-z]+)(?:\s*\(\s*(\d+)\s*\))?", ln)
        if f and not re.match(r"(?i)\s*(CREATE|CONSTRAINT)\b", ln):
            typ = SQL_TYPE.get(f.group(2).lower(), f.group(2).lower())
            fields.append({"name": f.group(1), "type": f"text({f.group(3)})" if typ == "text" and f.group(3) else typ,
                           "pk": f.group(1).lower() in keys or bool(re.search(r"(?i)\bPRIMARY\s+KEY\b", ln))})
    return {"name": mt.group(1) or mt.group(2), "linked": False, "source": "", "connect": "", "rows": None, "fields": fields, "line": 1}


# ---------- まとまり（1つの Access のファイルを書き出したもの）----------

class Access:
    """まとまりのモジュールから、Access の定義を取り出したもの。code_struct.build_struct が使う"""

    def __init__(self, modules):
        self.layouts, self.macros, self.queries, self.tables = {}, {}, {}, []
        self.relations, self.startup, self.log, self.qmodule = [], "", None, None
        self.unread = []                                             # 読めなかったもの: {"what", "where"}
        for m in modules:
            kind = m.get("kind")
            if kind in ("form", "report"):
                lay = parse_layout(m["lines"])
                if lay:
                    lay["name"] = re.sub(r"(?i)^(Form|Report)_", "", m["name"])
                    self.layouts[m["name"].lower()] = lay
            elif kind == "macro":
                rows, axl = macro_rows(m["lines"])
                parts, root, where = split_macro(rows), axl_root(rows) if axl else None, None
                if not parts and root is not None:                   # アクションの行が無く、XML の定義だけ（直接読み）
                    where = axl_where(root, rows)
                    parts = [(nm, []) for nm in axl_parts(root)] if where is not None else []
                self.macros[m["name"].lower()] = {"module": m["name"], "parts": parts, "axl": axl, "root": root, "where": where}
                if not parts:
                    self.unread.append({"what": "マクロ", "where": m["name"], "why": "アクションの行が見つからない" + ("（新しい形のマクロ）" if axl else "")})
            elif kind == "queries":
                self.qmodule = m["name"]
                for q in parse_queries(m["lines"]):
                    self.queries[q["name"].lower()] = q
            elif kind == "tables":
                t = parse_tables(m["lines"])
                self.tables, self.relations, self.startup = t["tables"], t["relations"], t["startup_form"]
            elif kind == "log":
                self.log = parse_log(m["lines"])
            elif kind == "query":                                    # アドインの書き出し: クエリ1件が1ファイル（SQL だけ）
                name = re.sub(r"^クエリ ", "", m["name"])
                body = [(k, ln) for k, ln in enumerate(m["lines"], 1) if ln.strip()]
                sql = " ".join(ln.strip().lstrip("\ufeff") for _k, ln in body)
                writes, reads = sql_tables(sql)
                self.queries[name.lower()] = {"name": name, "type": sql_type(sql), "connect": "", "lines": [body[0][0], body[-1][0]] if body else [1, 1],
                                              "sql": sql, "writes": writes, "reads": reads, "module": m["name"]}
            elif kind == "tabledef":
                t = parse_tabledef(m["lines"])
                if t and not any(x["name"].lower() == t["name"].lower() for x in self.tables):
                    self.tables.append(t)
            elif kind == "props":
                v = (load_json(m["lines"]).get("Items") or {}).get("StartupForm")
                self.startup = str(v.get("Value") if isinstance(v, dict) else v or "")

    def __bool__(self):
        return bool(self.layouts or self.macros or self.queries or self.tables or self.log)

    def code_from(self, m):
        """VBA として読みはじめる行（フォーム・レポートは、定義のあとのコードから）"""
        lay = self.layouts.get(m["name"].lower())
        if lay is None:
            return 1
        return lay["code_from"] or len(m["lines"]) + 1

    def macro_pid(self, name):
        """マクロの名前（「グループ.サブ」も）→ 手続きの番号。無ければ None"""
        low = name.strip().lower()
        if low in self.macros:
            mac = self.macros[low]
            return f"{mac['module']}.{mac['parts'][0][0] or '（マクロ）'}" if mac["parts"] else None
        grp, _, sub = low.rpartition(".")
        mac = self.macros.get(grp)
        hit = next((nm for nm, _rs in (mac or {"parts": []})["parts"] if nm.lower() == sub), None)
        return f"{mac['module']}.{hit}" if hit else None

    def query_io(self, name):
        """クエリの名前 → 入出力の並び [(種類, 名前, 向き)]。クエリそのものと、その SQL が読む・書くテーブル"""
        q = self.queries.get(name.strip().lower())
        if q is None:
            return [("クエリ", name, "out")], None
        act = q["type"] in ACTION_QUERY or bool(q["writes"])
        out = [("クエリ", q["name"], "out" if act else "in")]
        out += [("クエリ" if t.lower() in self.queries else "テーブル", t, "out") for t in q["writes"]]
        out += [("クエリ" if t.lower() in self.queries else "テーブル", t, "in") for t in q["reads"]]
        return out, {"module": q.get("module") or self.qmodule, "lines": q["lines"]}

    # ---- マクロを、手続きの木にする ----

    def macro_procs(self, resolve):
        """マクロ（と、フォームの埋め込みマクロ）を、VBA と同じ形の手続きにする。resolve: VBA の関数の名前 → 手続きの番号"""
        procs = {}
        for mac in self.macros.values():
            for sub, rows in mac["parts"]:
                pid = f"{mac['module']}.{sub or '（マクロ）'}"
                procs[pid] = self.macro_proc(pid, sub or "（マクロ）", mac["module"], rows, resolve, axl_statements(mac.get("root"), sub), mac.get("where"))
        return procs

    def embedded_procs(self, modules, resolve):
        procs, where = {}, {}
        for m in modules:
            lay = self.layouts.get(m["name"].lower())
            for key, (a, b) in ((lay or {}).get("embedded") or {}).items():
                every, axl = macro_rows(m["lines"], a + 1, b - 1)
                rows = [r for _nm, rs in split_macro(every) for r in rs]
                owner, _, prop = key.rpartition(".")
                name = f"{vb_ident(owner)}_{prop[2:] if prop.startswith('On') else prop}（埋め込みマクロ）"
                if rows:
                    pid = f"{m['name']}.{name}"
                    procs[pid] = self.macro_proc(pid, name, m["name"], rows, resolve, axl_statements(axl_root(every)) if axl else None)
                    where[(m["name"].lower(), key)] = pid
                else:
                    self.unread.append({"what": "埋め込みマクロ", "where": f"{m['name']} {a}〜{b} 行", "why": "アクションの行が見つからない" + ("（新しい形のマクロ）" if axl else "")})
        return procs, where

    def macro_proc(self, pid, name, module, rows, resolve, stmts=None, where=None):
        """アクションの行の並び → 手続き。stmts: 「_AXL:」の XML の文の並び（あれば、条件のブロックを、その形に戻す。行の対応が取れなければ、行の形から読む）。
        where: XML の定義だけのマクロ（直接読み）の、文の要素 → 行"""
        import code_lang
        info = code_lang.Info(pid)
        if where is not None:
            nodes = self.axl_nodes(stmts if stmts is not None else [], where, info, resolve, module)
            ends = [ln for n in code_lang.walk(nodes) for ln in n["lines"]] or [1]
            p = code_lang.finish(pid, name, module, "macro", [min(ends), max(ends)], nodes, {}, info.d)
            p["public"], p["read"] = True, "axl"
            return p
        nodes = self.block_nodes(stmts, rows, info, resolve) if stmts is not None else None
        parsed = nodes is not None
        if nodes is None:
            info = code_lang.Info(pid)
            nodes = self.row_nodes(rows, info, resolve)
            if stmts is not None and any(local(el) == "ConditionalBlock" for el in stmts.iter()):
                self.unread.append({"what": "マクロの条件のブロック", "where": f"{module} {rows[0]['lines'][0]}〜{rows[-1]['lines'][1]} 行",
                                    "why": "XML の定義（_AXL）と、アクションの行の対応が取れない。条件は、行の形のまま読んだ"})
        p = code_lang.finish(pid, name, module, "macro", [rows[0]["lines"][0], rows[-1]["lines"][1]], nodes, {}, info.d)
        p["public"] = True
        if parsed:
            p["read"] = "axl"                                        # 条件のブロックを、XML の定義のとおりに読めた
        return p

    def axl_nodes(self, stmts, where, info, resolve, module):
        """XML の定義だけから（直接読み）。呼び出し・クエリ・フォームなど、最初の引数で決まるものは、書き出しと同じに読む。
        ファイルやシートを読み書きするアクションは、引数の並びを XML の名前から決められないので、文として置いて「読めていない」に記録する"""
        norm = lambda t: re.sub(r"\s+", " ", t or "").strip()

        def walk(parent):
            out = []
            for el in parent:
                tag = local(el)
                if tag == "Action":
                    a, b = where[id(el)]
                    name = el.get("Name") or ""
                    args = [(c.get("Name") or "", c.text or "") for c in el if local(c) == "Argument"]
                    how, first = ACTIONS.get(name.lower(), ""), AXL_FIRST.get(name.lower())
                    shown = f"{name} {'、'.join(v for _k, v in args if v)}"[:120]
                    if first or how in ("stop", "tech"):
                        r = {"lines": [a, b], "props": [("Action", name, a)] + ([("Argument", dict(args).get(first, ""), a)] if first else [])}
                        n = self.row_node(r, info, resolve)
                        if n["kind"] != "exit":
                            n["text"] = shown
                    else:
                        n = {"kind": "stmt", "lines": [a, b], "exec": list(range(a, b + 1)), "text": shown}
                        if how:
                            self.unread.append({"what": "マクロのアクションの入出力", "where": f"{module} {a} 行", "why": f"{name} が読み書きするものは、直接読みでは取っていない"})
                    out.append(n)
                elif tag == "ConditionalBlock":
                    node = {"kind": "if", "branches": [], "exec": []}
                    for part in el:
                        label = {"If": "then", "ElseIf": "elseif", "Else": "else"}.get(local(part))
                        if label is None:
                            continue
                        head = list(where[id(part)])
                        if label != "else":
                            node["exec"] += list(range(head[0], head[1] + 1))
                        body = walk(next((c for c in part if local(c) == "Statements"), []))
                        node["branches"].append({"label": label, "cond": norm(next((c.text for c in part if local(c) == "Condition"), "")), "head": head, "body": body})
                    if not node["branches"]:
                        continue
                    import code_lang
                    inner = [ln for br in node["branches"] for n in code_lang.walk(br["body"]) for ln in n["lines"]]
                    node["head"], node["cond"] = list(node["branches"][0]["head"]), node["branches"][0]["cond"]
                    node["lines"] = [node["head"][0], max(inner + [br["head"][1] for br in node["branches"]])]
                    out.append(node)
                elif tag != "Comment":                               # まとまり（Group）など: 中の文を、そのまま並べる
                    inner = next((c for c in el if local(c) == "Statements"), None)
                    if inner is not None:
                        out += walk(inner)
            return out
        return walk(stmts)

    def row_nodes(self, rows, info, resolve):
        """行の形から（条件の列。「...」は、前の行と同じ条件）"""
        nodes, open_if = [], None
        for r in rows:
            a, b = r["lines"]
            get = lambda key: [(val, k) for nm, val, k in r["props"] if nm == key]
            cond = next((v for v, _k in get("Condition") if v), "")
            n = self.row_node(r, info, resolve)
            if cond == "..." and open_if is not None:                # 前の行と同じ条件
                open_if["branches"][0]["body"].append(n)
                open_if["lines"][1] = b
                continue
            if cond and cond != "...":
                ck = next(k for v, k in get("Condition") if v)
                n["lines"], n["exec"] = [ck + 1, b], list(range(ck + 1, b + 1))
                open_if = {"kind": "if", "head": [a, ck], "lines": [a, b], "exec": list(range(a, ck + 1)), "cond": cond,
                           "branches": [{"label": "then", "cond": cond, "head": [a, ck], "body": [n]}]}
                nodes.append(open_if)
                continue
            open_if = None
            nodes.append(n)
        return nodes

    def block_nodes(self, stmts, rows, info, resolve):
        """XML の定義の形（条件のブロックの入れ子）に、アクションの行を1つずつ当てはめる。当てはまらなければ None。
        Access は、条件のブロックを前の形で表すために、SetLocalVar __*L… の行を足している。条件の式は、その行にある"""
        acts = [r for r in rows if not is_helper(r)]
        helps = [r for r in rows if is_helper(r)]
        norm = lambda t: re.sub(r"\s+", " ", t or "").strip()
        state = {"i": 0, "floor": 0}

        def cond_row(text, upper):
            """条件の式を入れた行（floor より後、upper より前）"""
            for h in helps:
                args = [val for nm, val, _k in h["props"] if nm == "Argument"]
                if state["floor"] < h["lines"][0] < upper and len(args) > 1 and norm(args[1]) == norm(text):
                    return h
            return None

        def walk(parent):
            out = []
            for el in parent:
                tag = local(el)
                if tag == "Action":
                    if state["i"] >= len(acts):
                        raise LookupError
                    r = acts[state["i"]]
                    state["i"] += 1
                    one, two = next(v for nm, v, _k in r["props"] if nm == "Action").lower(), (el.get("Name") or "").lower()
                    if one != two and ACTIONS.get(one) and ACTIONS.get(two) and ACTIONS[one] != ACTIONS[two]:
                        raise LookupError
                    n = self.row_node(r, info, resolve)
                    state["floor"] = n["lines"][1]
                    out.append(n)
                elif tag == "ConditionalBlock":
                    node = {"kind": "if", "branches": [], "exec": []}
                    for part in el:
                        label = {"If": "then", "ElseIf": "elseif", "Else": "else"}.get(local(part))
                        if label is None:
                            continue
                        cond = norm(next((c.text for c in part if local(c) == "Condition"), ""))
                        upper = acts[state["i"]]["lines"][0] if state["i"] < len(acts) else 10 ** 9
                        if label == "else":
                            near = [h for h in helps if state["floor"] < h["lines"][0] < upper]
                            head = near[-1]["lines"] if near else None
                        else:
                            h = cond_row(cond, upper)
                            if h is None:
                                raise LookupError
                            head = h["lines"]
                            node["exec"] += list(range(head[0], head[1] + 1))
                        if head:
                            state["floor"] = head[1]
                        body = walk(next((c for c in part if local(c) == "Statements"), []))
                        if head is None:
                            head = [body[0]["lines"][0]] * 2 if body else [state["floor"]] * 2
                        node["branches"].append({"label": label, "cond": cond, "head": list(head), "body": body})
                    if not node["branches"]:
                        continue
                    upper = acts[state["i"]]["lines"][0] if state["i"] < len(acts) else 10 ** 9
                    tail = [h for h in helps if state["floor"] < h["lines"][0] < upper]   # ブロックの終わりの行（SetLocalVar … False）
                    if tail:
                        state["floor"] = tail[-1]["lines"][1]
                    node["head"], node["cond"] = list(node["branches"][0]["head"]), node["branches"][0]["cond"]
                    node["lines"] = [node["head"][0], state["floor"]]
                    out.append(node)
                elif tag != "Comment":                               # まとまり（Group）など: 中の文を、そのまま並べる
                    inner = next((c for c in el if local(c) == "Statements"), None)
                    if inner is not None:
                        out += walk(inner)
            return out
        try:
            nodes = walk(stmts)
        except (LookupError, StopIteration):
            return None
        return nodes if state["i"] == len(acts) else None

    def row_node(self, r, info, resolve):
        """アクションの行1つ → ノード（条件のブロックは、呼ぶ側が XML の定義から戻す）"""
        import code_struct as cs
        a, b = r["lines"]
        get = lambda key: [(val, k) for nm, val, k in r["props"] if nm == key]
        action, _at = get("Action")[0]
        args = [v for v, _k in get("Argument")]
        arg = lambda i: args[i] if i < len(args) else ""
        how = ACTIONS.get(action.lower(), "")
        n = {"kind": "stmt", "lines": [a, b], "exec": list(range(a, b + 1)), "text": f"{action} {'、'.join(x for x in args if x)}"[:120]}
        io, lit = [], lambda v: v and not v.startswith("=")
        if how in ("code", "macro"):                                 # VBA の関数・ほかのマクロを呼ぶ。呼び先が書き出しに無ければ、ふつうの文のまま（追えなかった呼び出しに記録する）
            to = resolve(re.sub(r"\s*\(.*$", "", arg(0)).lstrip("=")) if how == "code" else self.macro_pid(arg(0))
            if to:
                n["kind"], n["to"] = "call", to
            info.d["calls"].append({"to": to, "line": a, "head": False, **({} if to else {"name": arg(0)})})
        elif how == "query":
            items, via = self.query_io(arg(0))
            for kind, nm, way in items:
                info.io(io, kind, nm, way, a)
                if via and (kind, nm) != items[0][:2]:
                    next(x for x in io if (x["kind"], x["name"], x["way"]) == (kind, nm, way))["via"] = via
        elif how == "sql":
            outs, ins = sql_tables(arg(0))
            for t in outs:
                info.io(io, "テーブル", t, "out", a)
            for t in ins:
                info.io(io, "テーブル", t, "in", a)
        elif how == "table":
            info.io(io, "テーブル", arg(0), "in", a)
        elif how in ("sheet", "text"):
            imp = arg(0) in (("0", "2") if how == "sheet" else ("0", "1", "5", "6"))
            table, file = (arg(2), arg(3))
            if table:
                info.io(io, "テーブル", table, "out" if imp else "in", a)
            if lit(file):
                info.io(io, "ファイル", pathlib.PureWindowsPath(file).name, "in" if imp else "out", a)
            else:
                info.io(io, "ファイル", cs.DYNAMIC, "in" if imp else "out", a, file or action)
        elif how == "output":
            if arg(1):
                info.io(io, "クエリ" if arg(1).lower() in self.queries else "レポート" if arg(0) == "3" else "テーブル", arg(1), "in", a)
            if lit(arg(3)):
                info.io(io, "ファイル", pathlib.PureWindowsPath(arg(3)).name, "out", a)
            else:
                info.io(io, "ファイル", cs.DYNAMIC, "out", a, arg(3) or action)
        elif how == "report":
            info.io(io, "レポート", arg(0), "out", a)
        elif how == "mail":
            info.io(io, "メール", "メール", "out", a)
        elif how == "app":
            info.io(io, "外部プログラム", (arg(0).split(" ")[0].split("\\")[-1] or cs.DYNAMIC) if lit(arg(0)) else cs.DYNAMIC, "out", a, None if lit(arg(0)) else arg(0))
        elif how in ("msgbox", "form"):
            n["kind"] = "interact"
            n["interact"] = {"how": how, "text": arg(0)}
            info.d["inter"].append(dict(n["interact"], line=a))
        elif how == "stop":
            n = {"kind": "exit", "what": "proc", "label": None, "lines": [a, b], "exec": list(range(a, b + 1))}
        elif how == "tech":
            n["tech"] = "画面の更新の制御"
        if io:
            n["io"] = io
        return n

    def references(self, procs):
        """クエリの SQL・フォームやレポートの定義の式（="…"）の中に、名前が出てくる VBA の関数: [{"procedure", "by"}]"""
        funcs = {}
        for pid, p in procs.items():
            mod = p.get("mod")
            if mod is not None and mod.kind == "standard" and p.get("public") and p["kind"] == "function":
                funcs.setdefault(p["name"].lower(), pid)
        out, seen = [], set()

        def scan(text, by):
            for mt in re.finditer(r"([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)\s*\(", text):
                pid = funcs.get(mt.group(1).lower())
                if pid and (pid, by) not in seen:
                    seen.add((pid, by))
                    out.append({"procedure": pid, "by": by})
        for q in self.queries.values():
            scan(q["sql"], f"クエリ「{q['name']}」")
        for lay in self.layouts.values():
            for expr in lay.get("exprs") or []:
                scan(expr, f"{'フォーム' if lay['object'] == 'form' else 'レポート'}「{lay['name']}」の式")
        return out

    # ---- 構造に載せる要約 ----

    def summary(self):
        if not self:
            return None
        out = {"forms": [], "reports": [], "macros": [m["module"] for m in self.macros.values()],
               "queries": [{"name": q["name"], "type": q["type"], "lines": q["lines"], "writes": q["writes"], "reads": q["reads"],
                            **({"connect": q["connect"]} if q["connect"] else {})} for q in self.queries.values()],
               "tables": [{"name": t["name"], "fields": len(t["fields"]), **({"rows": t["rows"]} if t["rows"] is not None else {}),
                           **({"linked": True, "source": t["source"], "connect": t["connect"]} if t["linked"] else {})} for t in self.tables],
               "relations": self.relations, "unread": self.unread}
        for lay in self.layouts.values():
            out["forms" if lay["object"] == "form" else "reports"].append(
                {"name": lay["name"], "caption": lay["caption"], "record_source": lay["record_source"], "controls": len(lay["controls"]),
                 "events": len(lay["events"]), "code": bool(lay["code_from"])})
        if self.startup:
            out["startup_form"] = self.startup
        if self.log:
            out["source"], out["failed"] = self.log["source"], self.log["failed"]
            if self.log.get("via"):
                out["via"], out["notes"] = self.log["via"], self.log.get("notes") or []
            if self.log["error"]:
                out["error"] = self.log["error"]
        return out


# ---------- 表示 ----------

def show_folder(folder, out=print):
    folder = pathlib.Path(folder)
    mods = []
    for f in sorted(folder.iterdir()):
        if not f.is_file() or f.name == TMP_NAME:
            continue
        lines = read_head(f, 50_000_000).splitlines()
        d = describe(f, lines)
        if d:
            mods.append({"name": d["module"], "kind": d["module_kind"], "lines": lines, "file": f.name})
    acc = Access(mods)
    if not acc:
        out("Access を書き出したテキストが見つからない")
        return
    if acc.log:
        out(f"元のファイル: {acc.log['source']}（Access {acc.log['access']}）  書き出せなかったもの: {len(acc.log['failed'])}")
        for x in acc.log["failed"]:
            out(f"  × {x['kind']} {x['name']}: {x['why']}")
    for lay in acc.layouts.values():
        out(f"{'フォーム' if lay['object'] == 'form' else 'レポート'}「{lay['name']}」  見出し「{lay['caption']}」  元のデータ「{lay['record_source']}」  "
            f"コントロール {len(lay['controls'])}・コード {'あり' if lay['code_from'] else 'なし'}")
        for ev in lay["events"]:
            out(f"    {ev['line']:>5} 行  {event_label(lay, lay['name'], ev)}  → {ev['value']}")
    for mac in acc.macros.values():
        out(f"マクロ「{mac['module']}」  " + "／".join(f"{nm or '（本体）'}: {len(rs)} 行" for nm, rs in mac["parts"]) + ("  ※_AXL あり" if mac["axl"] else ""))
    for q in acc.queries.values():
        out(f"クエリ「{q['name']}」（{QUERY_JA.get(q['type'], q['type'])}）  書く: {'・'.join(q['writes']) or 'なし'}  読む: {'・'.join(q['reads']) or 'なし'}")
    for t in acc.tables:
        out(f"テーブル「{t['name']}」" + (f"（リンク: {t['connect']}）" if t["linked"] else "") + f"  項目 {len(t['fields'])}"
            + (f"・{t['rows']} 件" if t["rows"] is not None else ""))
    for x in acc.unread:
        out(f"読めなかった: {x['what']} {x['where']}（{x['why']}）")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", help="書き出したフォルダ（読めた中身を表示する）")
    show_folder(ap.parse_args().folder)
