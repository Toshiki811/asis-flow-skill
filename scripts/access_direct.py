#!/usr/bin/env python3
"""access_direct.py — Access のファイル（.accdb・.mdb）を、Access なしで直接読む。読めて、実物との答え合わせで合った部分だけを使う。

  python3 scripts/access_direct.py <ファイル>                  # 読めた中身の一覧を表示する
  python3 scripts/access_direct.py <ファイル> --out <フォルダ>  # 書き出しの道具（assets/access_export/）と同じ形のテキストを作る

ファイルのバイトを読むだけ。Access は使わない。マクロも VBA も実行しない。元のファイルには書き込まない。
読むもの:
  オブジェクトの一覧  システム テーブル MSysObjects（テーブル・リンク テーブル・クエリ・フォーム・レポート・マクロ・モジュールの名前）
  テーブルの定義      項目の名前と型・主キー・件数（行の中身は読まない）
  クエリ              MSysQueries の部品から、SQL の文を組み立てる（文そのものは、ファイルに入っていない）
  VBA のソース        MSysAccessStorage の中の VBA プロジェクト（Excel と同じ圧縮の形）を戻す
  フォームの定義      コントロールの名前・表示の文字・イベントの割り当て（Access 2007 以降の形で保存されたものだけ）
  マクロ              XML の定義（_AXL）が入っているものだけ
読めないもの（読めていない、と記録する。書き出し references/access-export.md で読む）:
  暗号化・パスワードつきのファイル、Access 97 以前の形式
  レポートの定義、前の形で保存されたフォームの定義、XML の定義が無いマクロ、フォームに埋め込んだマクロ、リレーション、起動時に開くフォーム
形式に、公式の説明は無い。ページとレコードの形は、公開されている解析（mdbtools の文書）、VBA は MS-OVBA、
フォームの定義の形は、実物（書き出しと対になったファイル）と突き合わせて確かめた。答え合わせの結果は references/access-export.md。
"""
import argparse
import hashlib
import pathlib
import re
import shutil
import struct
import sys
import xml.etree.ElementTree as ET

from common import mask_secrets
from extract_access import vb_ident

OBJECT_TYPE = {1: "table", 4: "linked", 6: "linked", 5: "query", -32768: "form", -32764: "report", -32766: "macro", -32761: "module"}
COL_TYPE = {1: "yesno", 2: "byte", 3: "integer", 4: "long", 5: "currency", 6: "single", 7: "double", 8: "datetime", 9: "binary", 10: "text", 11: "ole",
            12: "memo", 15: "guid", 16: "decimal", 18: "attachment", 19: "bigint"}
QUERY_TYPE = {1: "select", 2: "maketable", 3: "append", 4: "update", 5: "delete", 6: "crosstab", 7: "ddl", 8: "passthrough", 9: "union"}
QUERY_FLAG = {0: "select", 16: "crosstab", 32: "delete", 48: "update", 64: "append", 80: "maketable", 96: "ddl", 144: "passthrough", 128: "union"}   # MSysObjects の Flags
JOIN = {1: "INNER JOIN", 2: "LEFT JOIN", 3: "RIGHT JOIN"}
PARAM_TYPE = {0: "Value", 1: "Bit", 2: "Byte", 3: "Short", 4: "Long", 5: "Currency", 6: "IEEESingle", 7: "IEEEDouble", 8: "DateTime", 9: "Binary", 10: "Text",
              11: "LongBinary", 15: "Guid"}
# フォームの定義の、コントロールの種類（AcControlType）と、プロパティの番号。実物の書き出しと突き合わせて確かめたもの（ほかの番号は、名前が分からない）
CONTROL = {100: "Label", 101: "Rectangle", 102: "Line", 103: "Image", 104: "CommandButton", 105: "OptionButton", 106: "CheckBox", 107: "OptionGroup",
           108: "BoundObjectFrame", 109: "TextBox", 110: "ListBox", 111: "ComboBox", 112: "Subform", 114: "UnboundObjectFrame", 118: "PageBreak",
           119: "CustomControl", 122: "ToggleButton", 123: "Tab", 124: "Page", 126: "Attachment", 127: "EmptyCell", 128: "WebBrowser",
           129: "NavigationControl", 130: "NavigationButton", 152: "Section"}
PROP_TEXT = {17: "Caption", 20: "Name", 23: "DefaultValue", 27: "ControlSource", 91: "RowSource", 93: "RowSourceType", 132: "SourceObject",
             145: "ValidationRule", 156: "RecordSource"}
PROP_SKIP = {7, 18, 34, 61, 160, 343}                                # 図に使わない文字のプロパティ（Picture・ColumnWidths・FontName・ValidationText・DatasheetFontName・HyperlinkAddress）
PROP_EVENT = {77: "OnOpen", 78: "OnClose", 85: "BeforeUpdate", 86: "AfterUpdate", 110: "OnTimer", 117: "OnResize", 126: "OnClick", 127: "OnLoad",
              128: "OnUnload", 224: "OnDblClick"}
VALUE_SIZE = {1: (0, 1, 4), 2: (1,), 3: (2, 4), 4: (4,), 5: (8,), 6: (4,), 7: (8,), 8: (8,), 9: (16,)}
MARK_LAYOUT, MARK_GUESS = "AsisFlowLayout", "AsisFlowGuess"          # 直接読みのテキストだけにある印（extract_access.parse_layout が読む）
AXL_BLOCK = ("UserInterfaceMacro", "Statements", "ConditionalBlock", "If", "ElseIf", "Else", "Sub", "Group", "Action", "Comment")


class Unreadable(Exception):
    pass


# ---------- ページとレコード ----------

class JetFile:
    """Jet4・ACE（Access 2000 以降）のファイル。ページは 4096 バイト"""

    def __init__(self, path):
        self.data = pathlib.Path(path).read_bytes()
        if self.data[4:19] not in (b"Standard Jet DB", b"Standard ACE DB"):
            raise Unreadable("Access のファイルの形をしていない")
        self.version = self.data[0x14]
        if self.version == 0:
            raise Unreadable("Access 97 以前の形式（直接読みの対象外）")
        self.psize = 4096
        if len(self.data) < 3 * self.psize or self.page(2)[0] != 2:
            raise Unreadable("システム テーブルを読めない（暗号化されているか、壊れている）")
        self._defs = {}                                              # テーブルの定義のページ → 読んだ定義

    def page(self, n):
        return self.data[n * self.psize:(n + 1) * self.psize]

    def rows_of_page(self, pg):
        """データのページのレコード: [(番号, 中身, 別のページに移してある印)]。消したレコードと、移した先のレコードは除く（移した先は、移した元からたどる）"""
        n = struct.unpack_from("<H", pg, 12)[0]
        out = []
        for i in range(n):
            off = struct.unpack_from("<H", pg, 14 + 2 * i)[0]
            end = self.psize if i == 0 else struct.unpack_from("<H", pg, 14 + 2 * (i - 1))[0] & 0x1FFF
            if off & 0x8000:
                continue
            out.append((i, pg[off & 0x1FFF:end], bool(off & 0x4000)))
        return out

    def row_at(self, pointer):
        """ページとレコードの番号（上位 3 バイトがページ）→ レコードの中身"""
        page, row = pointer >> 8, pointer & 0xFF
        pg = self.page(page)
        if not pg or pg[0] != 1:
            return b""
        n = struct.unpack_from("<H", pg, 12)[0]
        if row >= n:
            return b""
        off = struct.unpack_from("<H", pg, 14 + 2 * row)[0]
        end = self.psize if row == 0 else struct.unpack_from("<H", pg, 14 + 2 * (row - 1))[0] & 0x1FFF
        return pg[off & 0x1FFF:end]

    def data_pages(self, tdef, buf):
        """テーブルが使っているデータのページ（定義にある「使っているページの表」から）"""
        m = self.row_at(struct.unpack_from("<I", buf, 55)[0])
        if not m:
            raise Unreadable("テーブルが使っているページの表を読めない")
        out = []
        if m[0] == 0:
            start, bits = struct.unpack_from("<I", m, 1)[0], m[5:]
            out = [start + i for i in range(len(bits) * 8) if bits[i // 8] & (1 << (i % 8))]
        else:
            per = (self.psize - 4) * 8
            for k in range((len(m) - 1) // 4):
                pn = struct.unpack_from("<I", m, 1 + 4 * k)[0]
                pg = self.page(pn) if pn else b""
                if pg and pg[0] == 5:
                    out += [k * per + i for i in range(per) if pg[4 + i // 8] & (1 << (i % 8))]
        return [n for n in out if self.page(n)[:1] == b"\x01" and struct.unpack_from("<I", self.page(n), 4)[0] == tdef]

    def table_buf(self, tdef):
        buf, n, seen = b"", tdef, set()
        while n and n not in seen:
            seen.add(n)
            pg = self.page(n)
            if not pg or pg[0] != 2:
                raise Unreadable("テーブルの定義を読めない")
            buf += pg if not buf else pg[8:]
            n = struct.unpack_from("<I", pg, 4)[0]
        return buf

    def table_def(self, tdef):
        """テーブルの定義のページ → {"rows": 件数, "columns": [{"name", "type", "num", "var", "fixed", "auto", "offset", "size"}], "pk": 主キーの列の番号, "pages"}"""
        if tdef in self._defs:
            return self._defs[tdef]
        buf = self.table_buf(tdef)
        rows = struct.unpack_from("<I", buf, 16)[0]
        ncols = struct.unpack_from("<H", buf, 45)[0]
        nidx, nreal = struct.unpack_from("<II", buf, 47)
        at = 63 + nreal * 12
        cols = []
        for _k in range(ncols):
            c = buf[at:at + 25]
            cols.append({"type": c[0], "num": struct.unpack_from("<H", c, 5)[0], "var": struct.unpack_from("<H", c, 7)[0], "fixed": bool(c[15] & 1),
                         "auto": bool(c[15] & 4), "offset": struct.unpack_from("<H", c, 21)[0], "size": struct.unpack_from("<H", c, 23)[0]})
            at += 25
        for c in cols:
            ln = struct.unpack_from("<H", buf, at)[0]
            c["name"] = buf[at + 2:at + 2 + ln].decode("utf-16-le", "replace")
            at += 2 + ln
        self._defs[tdef] = {"rows": rows, "columns": cols, "pk": self.primary_key(buf, at, nidx, nreal), "pages": self.data_pages(tdef, buf)}
        return self._defs[tdef]

    def primary_key(self, buf, at, nidx, nreal):
        """列の名前のあとの、索引の定義から、主キーの列の番号。読めなければ None（主キーは分からない、とする）"""
        try:
            real = []
            for _k in range(nreal):
                keys = [struct.unpack_from("<H", buf, at + 4 + 3 * i)[0] for i in range(10)]
                real.append([k for k in keys if k != 0xFFFF])
                at += 52
            for _k in range(nidx):
                which, kind = struct.unpack_from("<I", buf, at + 8)[0], buf[at + 23]
                at += 28
                if kind == 1:
                    return real[which] if which < len(real) else None
            return []
        except (struct.error, IndexError):
            return None

    def text(self, raw):
        if raw[:2] == b"\xff\xfe":                                   # 圧縮した文字（1字1バイト。0x00 で、圧縮のあり・なしが切り替わる）
            out, packed, i, body = [], True, 0, raw[2:]
            while i < len(body):
                if body[i] == 0:
                    packed = not packed
                    i += 1
                elif packed:
                    out.append(chr(body[i]))
                    i += 1
                else:
                    out.append(body[i:i + 2].decode("utf-16-le", "replace"))
                    i += 2
            return "".join(out)
        return raw.decode("utf-16-le", "replace")

    def long_value(self, field):
        """長い値（メモ・OLE）の 12 バイトの見出し → 中身のバイト"""
        if len(field) < 12:
            return b""
        size = struct.unpack_from("<I", field, 0)[0]
        flag, size = size >> 24, size & 0x3FFFFFFF
        if flag & 0x80:
            return field[12:12 + size]
        ptr = struct.unpack_from("<I", field, 4)[0]
        if flag & 0x40:
            return self.row_at(ptr)[:size]
        out, seen = b"", set()
        while ptr and ptr not in seen and len(out) < size:
            seen.add(ptr)
            row = self.row_at(ptr)
            if len(row) < 4:
                break
            out += row[4:]
            ptr = struct.unpack_from("<I", row, 0)[0]
        if len(out) < size:
            raise Unreadable("長い値を、最後まで読めない")
        return out[:size]

    def rows(self, tdef, want=None):
        """テーブルの行: 列の名前 → 値。want: 読む列の名前（省けば全部）。長い値は、バイトのまま返す（メモは文字）。
        1行でも形が合わなければ Unreadable（半端に読んだ結果は使わない）"""
        td = self.table_def(tdef)
        cols = [c for c in td["columns"] if want is None or c["name"] in want]
        out = []
        for n in td["pages"]:
            for _i, row, moved in self.rows_of_page(self.page(n)):
                if moved:
                    row = self.row_at(struct.unpack_from("<I", row, 0)[0]) if len(row) >= 4 else b""
                try:
                    out.append(self.parse_row(row, cols))
                except (struct.error, IndexError):
                    raise Unreadable("テーブルの行を読めない")
        return out

    def parse_row(self, row, cols):
        ncols = struct.unpack_from("<H", row, 0)[0]
        mask_len = (ncols + 7) // 8
        mask = row[len(row) - mask_len:]
        nvar = struct.unpack_from("<H", row, len(row) - mask_len - 2)[0]
        var_at = [struct.unpack_from("<H", row, len(row) - mask_len - 2 - 2 * (i + 1))[0] for i in range(nvar + 1)]
        out = {}
        for c in cols:
            present = c["num"] < ncols and bool(mask[c["num"] // 8] & (1 << (c["num"] % 8)))
            if c["type"] == 1:
                out[c["name"]] = present
                continue
            if not present:
                out[c["name"]] = None
                continue
            if c["fixed"]:
                raw = row[2 + c["offset"]:2 + c["offset"] + c["size"]]
            elif c["var"] < nvar:
                raw = row[var_at[c["var"]]:var_at[c["var"] + 1]]
            else:
                raw = b""
            t = c["type"]
            if t == 2:
                out[c["name"]] = raw[0] if raw else None
            elif t == 3:
                out[c["name"]] = struct.unpack("<h", raw[:2])[0] if len(raw) >= 2 else None
            elif t == 4:
                out[c["name"]] = struct.unpack("<i", raw[:4])[0] if len(raw) >= 4 else None
            elif t == 10:
                out[c["name"]] = self.text(raw)
            elif t == 12:
                out[c["name"]] = self.text(self.long_value(raw))
            elif t == 11:
                out[c["name"]] = self.long_value(raw)
            else:
                out[c["name"]] = raw
        return out


# ---------- 一覧・テーブル ----------

def catalog(db):
    """MSysObjects → [{"id", "name", "kind", "type", "flags", "database", "connect", "foreign"}]"""
    out = []
    for r in db.rows(2, {"Id", "Name", "Type", "Flags", "Database", "Connect", "ForeignName"}):
        if r.get("Name") is None or r.get("Type") is None:
            continue
        out.append({"id": r["Id"], "name": r["Name"], "type": r["Type"], "kind": OBJECT_TYPE.get(r["Type"], ""), "flags": r.get("Flags") or 0,
                    "database": r.get("Database") or "", "connect": r.get("Connect") or "", "foreign": r.get("ForeignName") or ""})
    return out


def is_system(o):
    return o["name"].lower().startswith(("msys", "usys", "~", "f_")) or bool(o["flags"] & 0x80000002)


def fingerprint(src):
    """書き出しの記録（access_export.txt）に書く、元のファイルの印の行。inventory.py が、前に読んだときから変わっていないかを見る"""
    return f"sha256\t{hashlib.sha256(pathlib.Path(src).read_bytes()).hexdigest()}"


def mask_secret(text):
    return mask_secrets([text or ""])[0][0]


def tables_of(db, cat):
    """利用者のテーブルとリンク テーブル: [{"name", "linked", "source", "connect", "rows", "fields": [{"name", "type", "pk"}]}]"""
    out = []
    for o in cat:
        if o["kind"] == "table" and not is_system(o):
            try:
                td = db.table_def(o["id"] & 0x00FFFFFF)
            except (Unreadable, struct.error, IndexError):
                out.append({"name": o["name"], "linked": False, "source": "", "connect": "", "rows": None, "fields": [], "unread": True})
                continue
            fields = []
            for c in sorted(td["columns"], key=lambda c: c["num"]):
                typ = f"text({c['size'] // 2})" if c["type"] == 10 else "autonumber" if c["auto"] and c["type"] == 4 else COL_TYPE.get(c["type"], str(c["type"]))
                fields.append({"name": c["name"], "type": typ, "pk": bool(td["pk"]) and c["num"] in td["pk"]})
            out.append({"name": o["name"], "linked": False, "source": "", "connect": "", "rows": td["rows"], "fields": fields, "pk_known": td["pk"] is not None})
        elif o["kind"] == "linked" and not is_system(o):
            connect = (o["connect"] or "") + (f";DATABASE={o['database']}" if o["database"] else "")
            out.append({"name": o["name"], "linked": True, "source": o["foreign"], "connect": mask_secret(connect), "rows": None, "fields": []})
    return out


# ---------- クエリ（部品から、SQL の文を組み立てる）----------

def quoted(expr):
    return len(expr) >= 2 and expr[0] == "[" and expr[-1] == "]"


def ident(name, whole=False):
    """名前を、要るときだけ角かっこで囲む（空白や記号を含む名前）。whole: 点で区切らない（別名）"""
    parts = [name] if whole else name.split(".")
    return ".".join(f"[{p}]" if re.search(r"\W", p) and not quoted(p) else p for p in parts)


class Source:
    """FROM の1つ分（テーブルか、結合）"""

    def __init__(self, key=None, text="", left=None, right=None, kind=0, on=""):
        self.key, self.text, self.left, self.right, self.kind, self.on = key, text, left, right, kind, on

    def has(self, name):
        return self.key == name if self.left is None else self.left.has(name) or self.right.has(name)

    def sql(self, nested=False):
        if self.left is None:
            return self.text
        body = f"{self.left.sql(True)} {JOIN.get(self.kind, 'JOIN')} {self.right.sql(True)} ON {self.on}"
        return f"({body})" if nested else body


def from_clause(tables, joins):
    """テーブルの行と、結合の行 → FROM に並べる文字の並び。組み立てられなければ Unreadable"""
    srcs = []
    for t in tables:
        text = (f"[{t['Expression']}]." if t.get("Expression") else "") + (ident(t["Name1"]) if t.get("Name1") else "")
        if t.get("Name2"):
            text += " AS " + ident(t["Name2"], True)
        srcs.append(Source(key=t.get("Name2") or t.get("Name1"), text=text))
    combo = {}
    for j in joins:                                                  # 同じ2つのテーブルの結合の条件は、AND でまとめる
        combo.setdefault((j.get("Name1"), j.get("Name2")), []).append(j)
    for (a, b), js in combo.items():
        if len({j.get("Flag") for j in js}) > 1:
            raise Unreadable("結合の種類が揃っていない")
        on = js[0]["Expression"] if len(js) == 1 else "(" + ") AND (".join(j["Expression"] or "" for j in js) + ")"
        left = next((s for s in srcs if s.has(a)), None)
        right = next((s for s in srcs if s.has(b)), None)
        if left is not None and left is right:
            raise Unreadable("結合の組み立てができない")
        for s in (left, right):
            if s is not None:
                srcs.remove(s)
        srcs.append(Source(left=left or Source(key=a, text=ident(a or "")), right=right or Source(key=b, text=ident(b or "")), kind=js[0].get("Flag") or 0, on=on or ""))
    return [s.sql() for s in srcs]


def query_sql(rows, kind):
    """MSysQueries の行（1つのクエリの分）と種類 → (SQL の文, 接続文字列)"""
    order = lambda r: (r.get("Order") or b"")
    by = lambda a: sorted((r for r in rows if r.get("Attribute") == a), key=order)
    one = lambda a: next(iter(by(a)), {})
    flag = one(3).get("Flag") or 0
    typ = one(1)
    clean = lambda s: (s or "").replace("\x00", "").strip()
    if kind in ("ddl", "passthrough"):
        return clean(typ.get("Expression")), mask_secret(typ.get("Name1") or "") if kind == "passthrough" else ""
    lines = []
    params = [f"{p['Name1']} {PARAM_TYPE.get(p.get('Flag') or 0, 'Value')}" for p in by(2) if p.get("Name1")]
    if params:
        lines.append(f"PARAMETERS {', '.join(params)};")
    cols, tables, joins = by(6), by(5), by(7)
    remote = one(4)
    into = lambda: (f" IN '{remote.get('Name1') or ''}'" + (f" [{remote['Expression']}]" if remote.get("Expression") else "")
                    if remote.get("Name1") or remote.get("Expression") else "")          # ほかのデータベースのテーブル
    orders = [clean(r.get("Expression")) + (" DESC" if (r.get("Name1") or "").upper() == "D" else "") for r in by(11)]
    if kind == "union":
        parts = [re.sub(r"[\r\n]+", "\n", clean(t.get("Expression"))) for t in tables]
        if len(parts) != 2:
            raise Unreadable("ユニオンの部品の数が合わない")
        lines += [parts[0], "UNION " + ("" if flag & 0x02 else "ALL ") + parts[1]]
        if orders:
            lines.append("ORDER BY " + ", ".join(orders))
        return "\n".join(lines) + ";", ""
    alias = lambda r: clean(r.get("Expression")) + (f" AS {ident(r['Name1'], True)}" if r.get("Name1") else "")
    where, having = clean(one(8).get("Expression")), clean(one(10).get("Expression"))
    groups, pivot = by(9), []
    if kind == "crosstab":
        head = [r for r in cols if not (r.get("Flag") or 0) & 3]
        pivot = [r for r in cols if (r.get("Flag") or 0) & 1]
        if len(head) != 1 or len(pivot) != 1:
            raise Unreadable("クロス集計の部品の数が合わない")
        lines.append("TRANSFORM " + alias(head[0]))
        cols = [r for r in cols if (r.get("Flag") or 0) & 2]
        groups = [r for r in groups if (r.get("Flag") or 0) & 2]
    froms = from_clause(tables, joins)
    star = ["*"] if flag & 0x01 else []
    pred = "DISTINCT " if flag & 0x02 else "DISTINCTROW " if flag & 0x08 else (f"TOP {one(3).get('Name1') or ''}{' PERCENT' if flag & 0x20 else ''} " if flag & 0x10 else "")

    def tail():
        out = []
        if froms:
            out.append("FROM " + ", ".join(froms) + (into() if kind != "maketable" and kind != "append" else ""))
        if where:
            out.append("WHERE " + where)
        if groups:
            out.append("GROUP BY " + ", ".join(clean(r.get("Expression")) for r in groups))
        if having:
            out.append("HAVING " + having)
        if orders:
            out.append("ORDER BY " + ", ".join(orders))
        return out
    if kind == "update":
        lines.append("UPDATE " + ", ".join(froms) + into())
        lines.append("SET " + ", ".join(f"{ident(r.get('Name2') or '')} = {clean(r.get('Expression'))}" for r in cols))
        if where:
            lines.append("WHERE " + where)
    elif kind == "delete":
        lines.append("DELETE " + ", ".join([alias(r) for r in cols] + star))
        lines += tail()
    elif kind == "append":
        values = [r for r in cols if (r.get("Flag") or 0) & 0x8000]
        target = [ident(r["Name2"]) for r in cols if r.get("Name2")]
        lines.append(f"INSERT INTO {ident(typ.get('Name1') or '')}{into()}" + (f" ({', '.join(target)})" if target else ""))
        if values:
            lines.append("VALUES (" + ", ".join(clean(r.get("Expression")) for r in values) + ")")
        else:
            lines.append("SELECT " + pred + ", ".join([clean(r.get("Expression")) for r in cols] + star))
            lines += tail()
    else:
        head = "SELECT " + pred + ", ".join([alias(r) for r in cols] + star)
        if kind == "maketable":
            head += f" INTO {ident(typ.get('Name1') or '')}{into()}"
        lines.append(head)
        lines += tail()
        if kind == "crosstab":
            lines.append("PIVOT " + clean(pivot[0].get("Expression")))
    if flag & 0x04:
        lines.append("WITH OWNERACCESS OPTION")
    return "\n".join(lines) + ";", ""


def queries_of(db, cat):
    """保存してあるクエリ: ([{"name", "type", "connect", "sql"}], 読めなかったクエリの名前と理由)。フォーム・レポートの中に埋めたクエリ（~sq_…）は除く"""
    sysq = next((o for o in cat if o["name"] == "MSysQueries" and o["type"] == 1), None)
    named = [o for o in cat if o["kind"] == "query" and not o["name"].startswith("~")]
    if sysq is None:
        return [], [(o["name"], "クエリの部品の表（MSysQueries）が無い") for o in named]
    parts = {}
    for r in db.rows(sysq["id"] & 0x00FFFFFF, {"Attribute", "Expression", "Flag", "Name1", "Name2", "ObjectId", "Order"}):
        parts.setdefault(r.get("ObjectId"), []).append(r)
    out, bad = [], []
    for o in named:
        rows = parts.get(o["id"], [])
        kind = next((QUERY_TYPE.get(r.get("Flag") or 0) for r in rows if r.get("Attribute") == 1), None) or QUERY_FLAG.get(o["flags"] & 0xF0)
        if not rows or kind is None:
            bad.append((o["name"], "クエリの部品が見つからない" if not rows else "クエリの種類が分からない"))
            continue
        try:
            sql, connect = query_sql(rows, kind)
        except (Unreadable, KeyError, TypeError, IndexError) as e:
            bad.append((o["name"], str(e) or "クエリの部品を組み立てられない"))
            continue
        out.append({"name": o["name"], "type": kind, "connect": connect, "sql": sql})
    return out, bad


# ---------- 入れ物（フォーム・マクロ・VBA）----------

def storage_of(db, cat):
    """MSysAccessStorage（フォーム・マクロ・VBA の入れ物）→ {パス（/ 区切り）: 中身のバイト}"""
    st = next((o for o in cat if o["name"] == "MSysAccessStorage" and o["type"] == 1), None)
    if st is None:
        return {}
    rows = db.rows(st["id"] & 0x00FFFFFF, {"Id", "ParentId", "Name", "Type", "Lv"})
    by = {r["Id"]: r for r in rows if r.get("Id") is not None}
    out = {}

    def path(r, depth=0):
        par = by.get(r.get("ParentId"))
        if par is None or depth > 12 or par is r:
            return ""
        return path(par, depth + 1) + "/" + (r.get("Name") or "")
    for r in rows:
        if r.get("Type") == 2 and r.get("Lv") is not None:
            out[path(r).lstrip("/")] = r["Lv"]
    return out


def dir_names(data):
    """入れ物の目次（DirData）→ {番号: 名前}。形が合わなければ None"""
    out, at = {}, 4
    while at < len(data):
        if data[at] != 4 or at + 2 > len(data) or data[at + 1] < 4 or at + 2 + data[at + 1] > len(data):
            return None
        ln = data[at + 1]
        out[struct.unpack_from("<I", data, at + ln - 2)[0]] = data[at + 2:at + ln - 2].decode("utf-16-le", "replace")
        at += 2 + ln
    return out


def decompress(data):
    """VBA の圧縮（MS-OVBA 2.4.1）を戻す"""
    from oletools.olevba import decompress_stream
    return bytes(decompress_stream(bytearray(data)))


def vba_of(storage):
    """VBA プロジェクト → ([{"name", "kind", "code"}], ロックが掛かっているか, 読めなかった理由)"""
    root = next((k[:-len("/dir")] for k in storage if k.lower().endswith("/vba/dir")), None)
    if root is None:
        return [], None, ""                                          # VBA の無いファイル
    try:
        d = decompress(storage[root + "/dir"])
    except Exception as e:
        return [], None, f"VBA プロジェクトの目次を戻せない（{type(e).__name__}）"
    mods, cur, at, codec = [], None, 0, "cp1252"
    while at + 6 <= len(d):
        rid, size = struct.unpack_from("<HI", d, at)
        if rid == 0x0009:                                            # PROJECTVERSION だけ、長さの書き方が違う
            at += 12
            continue
        body = d[at + 6:at + 6 + size]
        at += 6 + size
        if rid == 0x0003 and size >= 2:
            cp = struct.unpack_from("<H", body, 0)[0]
            codec = {65001: "utf-8", 10000: "mac_roman"}.get(cp, f"cp{cp}")
        elif rid == 0x0019:
            cur = {"name": body.decode(codec, "replace"), "stream": "", "offset": 0, "kind": "standard"}
            mods.append(cur)
        elif cur is not None and rid == 0x0047:
            cur["name"] = body.decode("utf-16-le", "replace")
        elif cur is not None and rid == 0x001A:
            cur["stream"] = body.decode(codec, "replace")
        elif cur is not None and rid == 0x0032:
            cur["stream"] = body.decode("utf-16-le", "replace")
        elif cur is not None and rid == 0x0031 and size >= 4:
            cur["offset"] = struct.unpack_from("<I", body, 0)[0]
        elif cur is not None and rid == 0x0022:
            cur["kind"] = "class"
    low = {k.lower(): k for k in storage}
    out, bad = [], []
    for m in mods:
        key = low.get(f"{root}/{m['stream']}".lower())
        try:
            code = decompress(storage[key][m["offset"]:]).decode(codec, "replace")
        except Exception:
            bad.append(m["name"])
            continue
        out.append({"name": m["name"], "kind": m["kind"], "code": code.replace("\r\n", "\n").replace("\r", "\n")})
    proj = next((storage[k] for k in storage if k.lower() == root.rsplit("/", 1)[0].lower() + "/project"), b"")
    locked = None
    mt = re.search(rb'(?m)^CMG="([0-9A-Fa-f]+)"', proj)
    if mt:
        import extract_vba
        locked = extract_vba.cmg_locked(mt.group(1).decode("ascii"))
    return out, locked, ("VBA のモジュールを戻せない: " + "、".join(bad[:8])) if bad else ""


# ---------- フォームの定義 ----------
# 見出し 10 バイト（版 u16・要る版 u16・0 u32・種類 u16）のあとに、オブジェクトが入れ子で並ぶ。
#   オブジェクト := 記録* 終わり    記録 := 枠 u32・プロパティの番号 u16・型 u32・幅 u32・長さ u32・値
#   終わり := FD 00 00 00 ｜ FE 00 00 00 子（付いているラベル）｜ FF 00 00 00 数 u16 子…    子 := 種類 u16 オブジェクト

def blob_record(b, at):
    if at + 18 > len(b):
        return None
    _slot, pid, typ, wide, ln = struct.unpack_from("<IHIII", b, at)
    if not 1 <= typ <= 12 or at + 18 + ln > len(b) or (wide >= 0x10000 and not (typ == 11 and wide == 0xFFFF0000)):
        return None
    if (typ in VALUE_SIZE and ln not in VALUE_SIZE[typ]) or (typ in (10, 12) and ln % 2):
        return None
    return pid, typ, b[at + 18:at + 18 + ln], at + 18 + ln


def blob_object(b, at, typ, depth=0):
    obj = {"type": typ, "props": [], "children": []}
    if depth > 16:
        raise ValueError("入れ子が深すぎる")
    while True:
        r = blob_record(b, at)
        if r is not None:
            obj["props"].append(r[:3])
            at = r[3]
            continue
        mark = b[at:at + 4]
        if mark == b"\xfd\x00\x00\x00":
            return obj, at + 4
        if mark == b"\xfe\x00\x00\x00" and at + 6 <= len(b):
            child, at = blob_object(b, at + 6, struct.unpack_from("<H", b, at + 4)[0], depth + 1)
            obj["children"].append(child)
            return obj, at
        if mark == b"\xff\x00\x00\x00" and at + 8 <= len(b):
            n = struct.unpack_from("<H", b, at + 4)[0]
            at += 6
            for _k in range(n):
                child, at = blob_object(b, at + 2, struct.unpack_from("<H", b, at)[0], depth + 1)
                obj["children"].append(child)
            return obj, at
        raise ValueError(f"位置 {at} で、形が合わない")


def layout_tree(blob):
    """フォームの定義 → (木, 版)。Access 2007 以降の形で、最後まで形が合ったときだけ。合わなければ Unreadable"""
    if len(blob) < 12:
        raise Unreadable("定義が空")
    version = struct.unpack_from("<H", blob, 0)[0]
    if version < 20:
        raise Unreadable(f"前の形式で保存されている（版 {version}）")
    try:
        obj, at = blob_object(blob, 10, 0)
    except (ValueError, struct.error, IndexError, RecursionError) as e:
        raise Unreadable(f"定義の形が合わない（{e}）")
    if at != len(blob):
        raise Unreadable("定義を、最後まで読めない")
    return obj, version


def quote(text):
    """SaveAsText の文字の書き方（\\ と " は \\ を前に付ける。改行などは 8 進）"""
    out = text.replace("\\", "\\\\").replace('"', '\\"')
    return '"' + re.sub(r"[\x00-\x1f]", lambda m: f"\\{ord(m.group(0)):03o}", out) + '"'


def layout_text(tree, version, kind, handlers, macros):
    """木 → (書き出しと同じ形のテキストの行, 気づき)。handlers: コードにある手続きの名前（小文字）→ 名前／macros: マクロの名前（小文字）の集まり。
    番号の分からないイベントの割り当て（[Event Procedure]）は、コードの手続きの名前（コントロール名_イベント）から決めて、印を付ける"""
    lines, notes = [f"Version ={version}", f"Begin {kind}"], {"guess": [], "open": [], "embedded": 0}

    def emit(obj, depth, top):
        pad = "    " * depth
        text = lambda v: v.decode("utf-16-le", "replace")
        name = next((text(v) for pid, typ, v in obj["props"] if pid == 20 and typ in (10, 12)), "")
        seen, unknown = set(), 0
        for pid, typ, v in obj["props"]:
            if typ not in (10, 12) or pid in PROP_SKIP:
                continue
            val = text(v)
            if pid in PROP_TEXT:
                lines.append(f"{pad}{PROP_TEXT[pid]} ={quote(val)}")
            elif pid in PROP_EVENT:
                lines.append(f"{pad}{PROP_EVENT[pid]} ={quote(val)}")
                seen.add(re.sub(r"^On", "", PROP_EVENT[pid]).lower())
                notes["embedded"] += val == "[Embedded Macro]"
            elif val == "[Event Procedure]":
                unknown += 1
            elif val == "[Embedded Macro]":
                notes["embedded"] += 1
            elif val.startswith("="):                                # 番号の分からないプロパティの式（イベントか、既定値などか、分からない）
                lines.append(f"{pad}Expression{pid} ={quote(val)}")
            elif val.lower() in macros or val.lower().split(".")[0] in macros:
                lines.append(f"{pad}OnUnknown{pid} ={quote(val)}")   # マクロの名前が入っている: 何かのイベントに、マクロを割り当ててある
        owner = kind if top else name
        if unknown and owner:
            prefix = vb_ident(owner).lower() + "_"
            cands = [h for low, h in handlers.items() if low.startswith(prefix) and low[len(prefix):] not in seen and "_" not in low[len(prefix):]]
            for h in cands:
                ev = h[len(prefix):]
                lines.append(f"{pad}{ev if re.match(r'(?i)(Before|After)', ev) else 'On' + ev} =\"[Event Procedure]\"")
            if cands:
                lines.append(f"{pad}{MARK_GUESS} ={quote(';'.join(h[len(prefix):] for h in cands))}")
                notes["guess"].append(f"{owner}（{len(cands)} 件）")
            if len(cands) != unknown:
                notes["open"].append(f"{owner}（割り当て {unknown} 件に、候補の手続き {len(cands)} 件）")
        if obj["children"]:
            lines.append(f"{pad}Begin")
            for c in obj["children"]:
                word = CONTROL.get(c["type"]) or ("Section" if c["type"] >= 150 else f"Control{c['type']}")
                lines.append(f"{pad}    Begin {word}")
                emit(c, depth + 2, False)
                lines.append(f"{pad}    End")
            lines.append(f"{pad}End")
    emit(tree, 1, True)
    lines.append("End")
    return lines, notes


def stub_layout(kind, why):
    return ["Version =0", f"Begin {kind}", f"    {MARK_LAYOUT} ={quote('unread: ' + why)}", "End"]


# ---------- マクロ ----------

def macro_xml(blob):
    """マクロの入れ物の中の、XML の定義（「_AXL:」で始まる文字を、順につないだもの）。無い・壊れていれば None"""
    mark = "_AXL:".encode("utf-16-le")
    out, at = [], blob.find(mark)
    while at >= 0:
        end = at + len(mark)
        while end + 1 < len(blob) and blob[end:end + 2] != b"\x00\x00":
            end += 2
        out.append(blob[at + len(mark):end].decode("utf-16-le", "replace"))
        at = blob.find(mark, end)
    text = "".join(out)
    body = re.sub(r"^\s*<\?xml[^>]*\?>\s*", "", text)
    if not body.strip():
        return None
    try:
        ET.fromstring(body)
    except ET.ParseError:
        return None
    return text


def macro_text(xml):
    """XML の定義 → 書き出しと同じ形のテキスト（_AXL: の行。アクション・条件を、1行ずつに分ける。つなげば、元の XML になる）"""
    ends = "|".join(x for x in AXL_BLOCK if x not in ("Action", "Comment"))
    cut = re.compile(rf"(?=<(?:\w+:)?(?:{'|'.join(AXL_BLOCK)})\b)|(?=</(?:\w+:)?(?:{ends})\s*>)")
    lines, inside, buf = ["Version =196611", "ColumnsShown =0"], False, ""
    for piece in (p for p in cut.split(xml) if p):
        if inside:                                                   # アクション・注釈の中は、切らない
            buf += piece
        else:
            if buf:
                lines += ["Begin", f"    Comment ={quote('_AXL:' + buf)}", "End"]
            buf = piece
            inside = bool(re.match(r"<(?:\w+:)?(Action|Comment)\b", piece)) and not re.match(r"<[^>]*/>", piece)
        if inside and re.search(r"</(?:\w+:)?(Action|Comment)\s*>", piece):
            inside = False
    if buf:
        lines += ["Begin", f"    Comment ={quote('_AXL:' + buf)}", "End"]
    return lines


# ---------- まとめて読む ----------

def read(path):
    """Access のファイル → {"version", "tables", "queries", "modules", "forms", "reports", "macros", "locked", "failed", "notes"}。読めなければ Unreadable。
    forms・reports: [{"name", "lines"（定義とコード）, "layout": 定義を読めたか, "why", "code"}]／macros: [{"name", "lines"}]／failed: [(種類, 名前, 理由)]"""
    db = JetFile(path)
    try:
        cat = catalog(db)
    except (struct.error, IndexError):
        raise Unreadable("オブジェクトの一覧（MSysObjects）を読めない（暗号化されているか、壊れている）")
    if not any(o["name"] == "MSysObjects" for o in cat):
        raise Unreadable("オブジェクトの一覧（MSysObjects）を読めない（暗号化されているか、壊れている）")
    notes, failed = [], []
    names = {k: [o["name"] for o in cat if o["kind"] == k and not o["name"].startswith("~")] for k in ("form", "report", "macro", "module")}
    try:
        storage = storage_of(db, cat)
    except (Unreadable, struct.error, IndexError) as e:
        storage = {}
        notes.append(f"フォーム・マクロ・VBA の入れ物（MSysAccessStorage）を読めない（{e}）")
    try:
        mods, locked, why = vba_of(storage)
    except ImportError:
        mods, locked, why = [], None, "oletools が無いので、VBA を読んでいない（pip3 install oletools）"
    if why:
        notes.append(why)
    by_low = {m["name"].lower(): m for m in mods}
    for nm in names["module"]:
        if nm.lower() not in by_low:
            failed.append(("module", nm, "VBA のソースを取り出せない"))
    try:
        queries, bad = queries_of(db, cat)
    except (Unreadable, struct.error, IndexError) as e:
        queries, bad = [], [(o["name"], str(e)) for o in cat if o["kind"] == "query" and not o["name"].startswith("~")]
    failed += [("query", nm, why) for nm, why in bad]
    macros = []
    spots = {"form": "Forms", "report": "Reports", "macro": "Scripts"}
    dirs = {k: dir_names(storage.get(f"{v}/\x03DirData", b"")) or {} for k, v in spots.items()}
    blob_of = lambda kind, nm: next((storage.get(f"{spots[kind]}/{n}/Blob") for n, x in dirs[kind].items() if x.lower() == nm.lower()), None)
    for nm in names["macro"]:
        xml = macro_xml(blob_of("macro", nm) or b"")
        if xml is None:
            failed.append(("macro", nm, "XML の定義が入っていない（前の形のマクロ）。直接読みでは読めない"))
        else:
            macros.append({"name": nm, "lines": macro_text(xml)})
    low_macros = {m.lower() for m in names["macro"]}
    out = {"form": [], "report": []}
    for kind in ("form", "report"):
        word = "Form" if kind == "form" else "Report"
        for nm in names[kind]:
            mod = by_low.get(f"{word}_{vb_ident(nm)}".lower()) or by_low.get(f"{word}_{nm}".lower())
            code = mod["code"].rstrip("\n").split("\n") if mod else []
            handlers = {mt.group(1).lower(): mt.group(1) for mt in (re.match(r"(?i)^\s*(?:Private\s+|Public\s+)?Sub\s+(\w+)", ln) for ln in code) if mt}
            blob, layout, why = blob_of(kind, nm), None, ""
            if kind == "report":
                why = "レポートの定義の直接読みは、実物で確かめていない"
            elif blob is None:
                why = "定義が見つからない"
            else:
                try:
                    tree, version = layout_tree(blob)
                    layout, ln = layout_text(tree, version, word, handlers, low_macros)
                    if ln["guess"]:
                        notes.append(f"フォーム「{nm}」のイベントの割り当てのうち、種類の番号が分からないものは、コードの手続きの名前から決めた: {'、'.join(ln['guess'][:8])}")
                    if ln["open"]:
                        notes.append(f"フォーム「{nm}」に、種類の分からないイベントの割り当てが残っている: {'、'.join(ln['open'][:8])}")
                    if ln["embedded"]:
                        notes.append(f"フォーム「{nm}」の埋め込みマクロ {ln['embedded']} 件は、直接読みでは読めていない")
                except Unreadable as e:
                    why = str(e)
            if layout is None:
                layout = stub_layout(word, why)
            out[kind].append({"name": nm, "layout": not why, "why": why, "code": bool(mod),
                              "lines": layout + ([f"CodeBehind{word}"] + code if mod else [])})
            if mod:
                mods.remove(mod)
    return {"version": db.version, "tables": tables_of(db, cat), "queries": queries, "modules": mods, "forms": out["form"], "reports": out["report"],
            "macros": macros, "locked": locked, "failed": failed, "notes": notes}


def write_export(path, out_dir):
    """直接読みの結果を、書き出しの道具と同じ形のテキストにして、out_dir に作る（前に作った書き出しは、作り直す）。→ read の結果"""
    src = pathlib.Path(path)
    res = read(src)
    out = pathlib.Path(out_dir)
    if out.exists():
        if any(x.is_dir() for x in out.iterdir()) or (any(out.iterdir()) and not (out / "access_export.txt").exists()):
            raise Unreadable(f"書き出し先 {out} に、ほかのものが入っている（空のフォルダか、前に作った書き出しを指定する）")
        shutil.rmtree(out)
    out.mkdir(parents=True)
    used = set()

    def put(prefix, name, suffix, lines):
        stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name) or "_"
        k, file = 1, f"{prefix}{stem}{suffix}"
        while file.lower() in used:
            k += 1
            file = f"{prefix}{stem}~{k}{suffix}"
        used.add(file.lower())
        (out / file).write_text("\n".join(lines) + "\n", encoding="utf-8")
        return file
    log = ["tool\tasis-flow access_direct", f"source\t{src.name}", "via\tdirect", f"access\tfile-format-{res['version']}",
           fingerprint(src)]
    for kind, key, prefix, suffix in (("form", "forms", "F_", ".form.txt"), ("report", "reports", "R_", ".report.txt")):
        for f in res[key]:
            log.append(f"{kind}\t{f['name']}\t{put(prefix, f['name'], suffix, f['lines'])}\tok" + ("" if f["layout"] else f" (layout unread: {f['why']})"))
    for m in res["macros"]:
        log.append(f"macro\t{m['name']}\t{put('X_', m['name'], '.macro.txt', m['lines'])}\tok")
    for m in res["modules"]:
        log.append(f"module\t{m['name']}\t{put('M_', m['name'], '.bas', m['code'].rstrip(chr(10)).split(chr(10)))}\tok")
    for kind, name, why in res["failed"]:
        if kind != "query":
            log.append(f"{kind}\t{name}\t\tFAILED {why}")
    lines = []
    for t in res["tables"]:
        if t["linked"]:
            lines.append(f"[linked]\t{t['name']}\tsource={t['source']}\tconnect={t['connect']}")
        else:
            lines.append(f"[table]\t{t['name']}" + (f"\trows={t['rows']}" if t["rows"] is not None else ""))
            lines += [f"  {f['name']}\t{f['type']}" + ("\tpk" if f["pk"] else "") for f in t["fields"]]
    (out / "access_tables.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    log.append(f"tables\t{len(res['tables'])}\taccess_tables.txt\tok")
    lines = []
    for q in res["queries"]:
        lines.append(f"-- query: {q['name']}\ttype={q['type']}")
        if q["connect"]:
            lines.append(f"-- connect: {q['connect']}")
        lines += q["sql"].split("\n") + [""]
    (out / "access_queries.sql").write_text("\n".join(lines) + "\n", encoding="utf-8")
    log.append(f"queries\t{len(res['queries'])}\taccess_queries.sql\tok")
    for kind, name, why in res["failed"]:
        if kind == "query":
            log.append(f"note\tクエリ「{name}」は、直接読みでは読めていない（{why}）")
    log += [f"note\t{n}" for n in res["notes"]]
    if res["locked"]:
        log.append("note\tVBA のプロジェクトに、閲覧のロックが掛かっている。読んでよいかを、利用者に確かめる")
    log.append("note\tリレーションと、起動時に開くフォームは、直接読みでは読んでいない")
    (out / "access_export.txt").write_text("\n".join(log) + "\n", encoding="utf-8")
    return res


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out", help="書き出しと同じ形のテキストを作るフォルダ（空のフォルダか、前に作った書き出し）")
    a = ap.parse_args()
    try:
        res = write_export(a.file, a.out) if a.out else read(a.file)
    except Unreadable as e:
        sys.exit(f"読めない: {e}")
    print(f"形式の版 {res['version']}  テーブル {len(res['tables'])}・クエリ {len(res['queries'])}・フォーム {len(res['forms'])}・レポート {len(res['reports'])}・"
          f"マクロ {len(res['macros'])}・モジュール {len(res['modules'])}" + ("・VBA に閲覧のロック" if res["locked"] else ""))
    for t in res["tables"]:
        print(f"  テーブル「{t['name']}」" + (f"（リンク: {t['source']}  {t['connect']}）" if t["linked"] else f"  項目 {len(t['fields'])}・{t['rows']} 件"))
    for q in res["queries"]:
        print(f"  クエリ「{q['name']}」（{q['type']}）  {q['sql'].splitlines()[0][:80] if q['sql'] else ''}")
    for key, word in (("forms", "フォーム"), ("reports", "レポート")):
        for f in res[key]:
            print(f"  {word}「{f['name']}」  定義: {'読めた' if f['layout'] else '読めていない（' + f['why'] + '）'}  コード: {'あり' if f['code'] else 'なし'}")
    for m in res["macros"]:
        print(f"  マクロ「{m['name']}」")
    for m in res["modules"]:
        print(f"  VBA「{m['name']}」（{m['kind']}）  {len(m['code'].splitlines())} 行")
    for kind, name, why in res["failed"]:
        print(f"  ※ 読めていない: {kind}「{name}」（{why}）")
    for n in res["notes"]:
        print(f"  ※ {n}")
    if a.out:
        print(f"→ {a.out}")


if __name__ == "__main__":
    sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
    main()
