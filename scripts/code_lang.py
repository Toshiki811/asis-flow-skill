#!/usr/bin/env python3
"""code_lang.py — VBA 以外の言語の、構造の取り出し（code_struct.py が使う。コードは実行しない・取り込まない）。

  Python                          標準の ast で構文から（parsed）。関数・分岐・繰り返し・例外・呼び出し・ファイルとテーブル・入力の求め
  SQL                             文の単位（heuristic）。読むテーブル・書くテーブル。手続き型の制御（IF・WHILE）は、文の中身として扱う
  バッチ・PowerShell・シェル      行の形から（heuristic）。if・繰り返し・関数・ラベル・呼び出し・ファイルの操作・外部プログラム
どれも、VBA と同じ形の木（seq・if・loop・call・interact・exit）にする。読めない書き方のファイルがあれば None を返し、構造は LLM の読取に回る。
"""
import ast
import pathlib
import re
import types

import code_struct as cs
import extract_access
from common import site



def group(nodes, comment_at):
    ps = cs.VbaParser.__new__(cs.VbaParser)
    return ps.group(nodes, types.SimpleNamespace(comment_at=comment_at))


def walk(nodes):
    for n in nodes:
        yield n
        for b in n.get("branches") or []:
            yield from walk(b["body"])
        yield from walk(n.get("body") or [])


def finish(pid, name, module, kind, lines, nodes, comment_at, info):
    body = group(nodes, comment_at)
    flat = list(walk(body))
    return {"id": pid, "name": name, "module": module, "kind": kind, "lines": lines, "body": body, "handlers": [], "public": True, "params": False,
            "calls": info["calls"], "io": info["io"], "interactions": info["inter"], "dynamic": info["dynamic"],
            "exec": sorted({ln for n in flat for ln in n.get("exec", [])}),
            "boxes": sum(1 for n in flat if n["kind"] in ("call", "interact", "if", "select", "loop") or (n["kind"] == "seq" and not n.get("tech")))}


def sql_tables(sql):
    """SQL の文字 → [(テーブル, 向き)]。書く先を先に、読む元をあとに（読み方は extract_access.sql_tables の 1 つ）"""
    outs, ins = extract_access.sql_tables(sql)
    return [(t, "out") for t in outs] + [(t, "in") for t in ins]


class Info:
    """手続き1つの、呼び出し・入出力・やり取りの記録"""

    def __init__(self, pid):
        self.pid, self.d = pid, {"calls": [], "io": [], "inter": [], "dynamic": []}

    def io(self, out, kind, name, way, line, dynamic=None):
        kind = cs.io_kind(kind)
        if (kind, name, way) not in [(x["kind"], x["name"], x["way"]) for x in out]:
            out.append({"kind": kind, "name": name, "way": way, "line": line})
        self.d["io"].append({"kind": kind, "name": name, "way": way, "line": line, "procedure": self.pid})
        if dynamic:
            self.d["dynamic"].append({"line": line, "what": kind, "expr": str(dynamic)[:80]})


# ---------- Python ----------

PY_READ = ("read_csv", "read_excel", "read_json", "read_table", "read_parquet", "load_workbook", "loadtxt", "genfromtxt")
PY_WRITE = ("to_csv", "to_excel", "to_json", "to_parquet", "savetxt", "savefig")
PY_LOG = re.compile(r"^(print|logging\.\w+|logger\.\w+|log\.\w+|pprint|warnings\.warn)$")


class PyParser:
    def __init__(self, modules):
        self.mods = modules
        self.procs, self.entries, self.funcs, self.consts = {}, [], {}, {}
        self.trees = {m["name"]: ast.parse("\n".join(m["lines"])) for m in modules}     # 読めない書き方なら SyntaxError
        for m in modules:
            for node in self.trees[m["name"]].body:
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    self.funcs.setdefault(node.name, []).append(f"{m['name']}.{node.name}")
                elif isinstance(node, ast.ClassDef):
                    for x in node.body:
                        if isinstance(x, (ast.FunctionDef, ast.AsyncFunctionDef)):
                            self.funcs.setdefault(x.name, []).append(f"{m['name']}.{node.name}.{x.name}")
                elif isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                    for t in node.targets:
                        if isinstance(t, ast.Name):
                            self.consts[t.id] = node.value.value
        for m in modules:
            self.module(m)

    def module(self, m):
        tree, self.m = self.trees[m["name"]], m
        self.comment_at = {k: ln.strip().lstrip("#").strip() for k, ln in enumerate(m["lines"], 1) if ln.strip().startswith("#")}
        main = []
        for node in tree.body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                self.define(node, f"{m['name']}.{node.name}", node.name)
            elif isinstance(node, ast.ClassDef):
                for x in node.body:
                    if isinstance(x, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        self.define(x, f"{m['name']}.{node.name}.{x.name}", f"{node.name}.{x.name}")
            elif isinstance(node, (ast.Import, ast.ImportFrom)) or (isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant)):
                continue
            elif isinstance(node, ast.If) and "__name__" in ast.unparse(node.test):
                main += node.body
            elif not (isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant)):
                main.append(node)
        if main:
            pid = f"{m['name']}.（本体）"
            self.info = Info(pid)
            nodes = self.stmts(main)
            self.procs[pid] = finish(pid, "（本体）", m["name"], "script", [main[0].lineno, max(x.end_lineno for x in main)], nodes, self.comment_at, self.info.d)
            self.entries.append({"procedure": pid, "trigger": "script", "label": f"スクリプト「{m['name']}」を実行", "evidence": "関数の外に書かれた文（スクリプトの本体）"})

    def define(self, node, pid, name):
        self.info = Info(pid)
        body = [x for x in node.body if not (isinstance(x, ast.Expr) and isinstance(x.value, ast.Constant))]
        nodes = self.stmts(body)
        self.procs[pid] = finish(pid, name, self.m["name"], "function", [node.lineno, node.end_lineno], nodes, self.comment_at, self.info.d)

    def text(self, a, b):
        return " ".join(x.strip() for x in self.m["lines"][a - 1:b])

    def head(self, node, body):
        a = node.lineno
        b = max(a, min(x.lineno for x in body) - 1) if body else node.end_lineno
        return [a, b]

    def stmts(self, body):
        out = []
        for n in body:
            a, b = n.lineno, n.end_lineno
            if isinstance(n, ast.If):
                node = {"kind": "if", "head": self.head(n, n.body), "lines": [a, b], "exec": list(range(*[self.head(n, n.body)[0], self.head(n, n.body)[1] + 1])), "branches": []}
                self.note(node, n.test)
                cur, label = n, "then"
                while True:
                    node["branches"].append({"label": label, "cond": ast.unparse(cur.test), "head": self.head(cur, cur.body), "body": self.stmts(cur.body)})
                    if len(cur.orelse) == 1 and isinstance(cur.orelse[0], ast.If) and cur.orelse[0].col_offset == cur.col_offset:
                        cur, label = cur.orelse[0], "elseif"
                        h = self.head(cur, cur.body)
                        node["exec"] += list(range(h[0], h[1] + 1))
                        self.analyse(cur.test, h[0], [])
                        continue
                    if cur.orelse:
                        node["branches"].append({"label": "else", "cond": "", "head": [cur.orelse[0].lineno, cur.orelse[0].lineno], "body": self.stmts(cur.orelse)})
                    break
                out.append(node)
            elif isinstance(n, (ast.For, ast.AsyncFor, ast.While)):
                h = self.head(n, n.body)
                cond = ast.unparse(n.test) if isinstance(n, ast.While) else f"{ast.unparse(n.target)} in {ast.unparse(n.iter)}"
                node = {"kind": "loop", "form": "while" if isinstance(n, ast.While) else "for", "cond": cond, "head": h, "lines": [a, b],
                        "exec": list(range(h[0], h[1] + 1)), "body": self.stmts(n.body)}
                self.note(node, n.test if isinstance(n, ast.While) else n.iter)
                out.append(node)
                out += self.stmts(n.orelse)
            elif isinstance(n, ast.Try):
                out += self.stmts(n.body)
                for h in n.handlers:
                    what = ast.unparse(h.type) if h.type is not None else ""
                    out.append({"kind": "if", "head": [h.lineno, h.lineno], "lines": [h.lineno, h.end_lineno], "exec": [h.lineno], "cond": f"例外が起きたとき（{what}）" if what else "例外が起きたとき",
                                "branches": [{"label": "then", "cond": "", "head": [h.lineno, h.lineno], "body": self.stmts(h.body)}]})
                out += self.stmts(n.orelse) + self.stmts(n.finalbody)
            elif isinstance(n, (ast.With, ast.AsyncWith)):
                h = self.head(n, n.body)
                out.append(self.plain(n.items[0].context_expr if n.items else n, h[0], h[1], self.text(h[0], h[1])))
                out += self.stmts(n.body)
            elif isinstance(n, (ast.Return, ast.Raise)) or (isinstance(n, ast.Expr) and isinstance(n.value, ast.Call)
                                                          and ast.unparse(n.value.func) in ("sys.exit", "exit", "quit", "os._exit")):
                if isinstance(n, ast.Return) and n.value is not None:
                    self.analyse(n.value, a, [])
                out.append({"kind": "exit", "what": "proc", "label": None, "lines": [a, b], "exec": list(range(a, b + 1))})
            elif isinstance(n, (ast.Break, ast.Continue)):
                out.append({"kind": "exit", "what": "loop" if isinstance(n, ast.Break) else "continue", "label": None, "lines": [a, b], "exec": list(range(a, b + 1))})
            elif isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Import, ast.ImportFrom, ast.Global, ast.Nonlocal, ast.Pass)):
                continue
            else:
                out.append(self.plain(n, a, b, self.text(a, b)))
        return out

    def note(self, node, expr):
        info = self.analyse(expr, node["head"][0], [])
        node.setdefault("cond", ast.unparse(expr))
        for k in ("interact", "calls", "io"):
            if info.get(k):
                node[k] = info[k]

    def name_of(self, arg):
        """ファイルを指す式 → (名前, 実行時に決まるか)"""
        if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
            return pathlib.PureWindowsPath(arg.value.replace("/", "\\")).name or arg.value, False
        if isinstance(arg, ast.Name) and arg.id in self.consts:
            return pathlib.PureWindowsPath(self.consts[arg.id].replace("/", "\\")).name, False
        if isinstance(arg, ast.JoinedStr):
            t = "".join(v.value if isinstance(v, ast.Constant) else f"＜{ast.unparse(v.value)[:20]}＞" for v in arg.values)
            tail = re.split(r"[\\/]", t)[-1]
            return (tail or t), not tail.strip("＜＞")
        if isinstance(arg, ast.Call) and arg.args and ast.unparse(arg.func).endswith(("join", "Path")):
            return self.name_of(arg.args[-1])
        if isinstance(arg, ast.BinOp) and isinstance(arg.op, (ast.Div, ast.Add)):
            nm, dyn = self.name_of(arg.right)
            return (nm, dyn) if not dyn else (f"＜{ast.unparse(arg)[:20]}＞", True)
        return cs.DYNAMIC, True

    def analyse(self, node, line, out):
        res = {"calls": [], "interact": None, "io": out, "log": True, "any": False}
        add = self.info.io
        for c in [x for x in ast.walk(node) if isinstance(x, ast.Call)]:
            fn = ast.unparse(c.func)
            last = fn.split(".")[-1]
            res["any"] = True
            res["log"] = res["log"] and bool(PY_LOG.match(fn))
            arg0 = c.args[0] if c.args else None
            cands = self.funcs.get(last, []) if (isinstance(c.func, ast.Name) or len(self.funcs.get(last, [])) == 1) else []
            if cands:
                pid = next((p for p in cands if p.startswith(self.m["name"] + ".")), cands[0])
                if pid not in res["calls"]:
                    res["calls"].append(pid)
                    self.info.d["calls"].append({"to": pid, "line": line, "head": False})
            if fn == "input":
                res["interact"] = {"how": "inputbox", "text": arg0.value if isinstance(arg0, ast.Constant) and isinstance(arg0.value, str) else ""}
            elif "messagebox." in fn:
                res["interact"] = {"how": "msgbox", "text": next((a.value for a in c.args[::-1] if isinstance(a, ast.Constant) and isinstance(a.value, str)), "")}
            if fn == "open" and arg0 is not None:
                mode = next((a.value for a in c.args[1:2] if isinstance(a, ast.Constant)), None) or next((k.value.value for k in c.keywords if k.arg == "mode" and isinstance(k.value, ast.Constant)), "r")
                nm, dyn = self.name_of(arg0)
                add(out, "ファイル", nm, "out" if any(x in str(mode) for x in "wax+") else "in", line, ast.unparse(arg0) if dyn else None)
            elif last in PY_READ + PY_WRITE and (arg0 is not None or c.keywords):
                nm, dyn = self.name_of(arg0) if arg0 is not None else (cs.DYNAMIC, True)
                add(out, "ファイル", nm, "in" if last in PY_READ else "out", line, ast.unparse(arg0) if dyn and arg0 is not None else None)
            elif last == "save" and arg0 is not None and isinstance(arg0, (ast.Constant, ast.JoinedStr, ast.Name)):
                nm, dyn = self.name_of(arg0)
                add(out, "ファイル", nm, "out", line, ast.unparse(arg0) if dyn else None)
            elif fn in ("shutil.copy", "shutil.copy2", "shutil.copyfile", "shutil.move", "os.rename", "os.replace") and len(c.args) >= 2:
                for arg, way in ((c.args[0], "in"), (c.args[1], "out")):
                    nm, dyn = self.name_of(arg)
                    add(out, "ファイル", nm, way, line, ast.unparse(arg) if dyn else None)
            elif fn in ("os.remove", "os.unlink", "os.makedirs", "os.mkdir", "shutil.rmtree") and arg0 is not None:
                nm, dyn = self.name_of(arg0)
                add(out, "ファイル", nm, "out", line, ast.unparse(arg0) if dyn else None)
            elif fn in ("subprocess.run", "subprocess.call", "subprocess.Popen", "subprocess.check_call", "subprocess.check_output", "os.system") and arg0 is not None:
                first = arg0.elts[0] if isinstance(arg0, (ast.List, ast.Tuple)) and arg0.elts else arg0
                nm = first.value.split(" ")[0] if isinstance(first, ast.Constant) and isinstance(first.value, str) else cs.DYNAMIC
                add(out, "外部プログラム", pathlib.PureWindowsPath(nm.replace("/", "\\")).name or nm, "out", line, None if nm != cs.DYNAMIC else ast.unparse(arg0))
            elif fn.startswith("smtplib.") or last in ("sendmail", "send_message"):
                add(out, "メール", "メール", "out", line)
            elif fn in ("requests.get", "requests.post", "requests.put", "requests.delete", "urllib.request.urlopen") and arg0 is not None:
                url = arg0.value if isinstance(arg0, ast.Constant) and isinstance(arg0.value, str) else ""
                host = re.sub(r"^\w+://([^/:]+).*$", r"\1", url) if url else cs.DYNAMIC
                add(out, "外部プログラム", host, "in" if last in ("get", "urlopen") else "out", line, None if url else ast.unparse(arg0))
            elif last == "connect" and fn.split(".")[0] in ("sqlite3", "pyodbc", "psycopg2", "pymysql", "cx_Oracle", "oracledb", "MySQLdb", "pymssql"):
                nm, _dyn = self.name_of(arg0) if arg0 is not None else (fn.split(".")[0], False)
                add(out, "データベース", nm if nm != cs.DYNAMIC else fn.split(".")[0], "in", line)
            if last in ("execute", "executemany", "read_sql", "read_sql_query", "read_sql_table") and arg0 is not None:
                sql = arg0.value if isinstance(arg0, ast.Constant) and isinstance(arg0.value, str) else \
                    "".join(v.value if isinstance(v, ast.Constant) else " ? " for v in arg0.values) if isinstance(arg0, ast.JoinedStr) else self.consts.get(getattr(arg0, "id", ""), "")
                tabs = sql_tables(sql)
                for t, way in tabs:
                    add(out, "テーブル", t, way, line)
                if not tabs:
                    add(out, "テーブル", cs.DYNAMIC, "in" if last.startswith("read") else "out", line, ast.unparse(arg0))
        return res

    def plain(self, node, a, b, text):
        io = []
        res = self.analyse(node, a, io)
        n = {"kind": "interact" if res["interact"] else "call" if res["calls"] else "stmt", "lines": [a, b], "exec": list(range(a, b + 1)), "text": text[:120]}
        if res["interact"]:
            n["interact"] = res["interact"]
            self.info.d["inter"].append(dict(res["interact"], line=a))
        if res["calls"]:
            n["to"] = res["calls"][0]
        if io:
            n["io"] = io
        if n["kind"] == "stmt" and not io and res["any"] and res["log"]:
            n["tech"] = "ログ"
        return n


def parse_python(modules):
    try:
        ps = PyParser(modules)
    except (SyntaxError, ValueError, RecursionError):
        return None
    return ps.procs, ps.entries, "parsed"


# ---------- SQL ----------

def parse_sql(modules):
    """SQL のスクリプト: 文（; まで）を1つの文として並べる。読む・書くテーブルを取る"""
    procs, entries = {}, []
    for m in modules:
        pid, info, nodes, buf, start = f"{m['name']}.（本体）", Info(f"{m['name']}.（本体）"), [], [], None
        comment_at = {k: ln.strip().lstrip("-").strip() for k, ln in enumerate(m["lines"], 1) if ln.strip().startswith("--")}
        block = False
        for k, raw in enumerate(m["lines"], 1):
            ln = raw.strip()
            if block or ln.startswith("/*"):
                block = "*/" not in ln
                continue
            code = re.sub(r"--.*$", "", ln).strip()
            if not code:
                continue
            start = start or k
            buf.append(code)
            if code.endswith(";") or re.fullmatch(r"(?i)(GO|/)", code):
                sql = " ".join(buf)
                io = []
                for t, way in sql_tables(sql):
                    info.io(io, "テーブル", t, way, start)
                n = {"kind": "stmt", "lines": [start, k], "exec": list(range(start, k + 1)), "text": sql[:120]}
                if io:
                    n["io"] = io
                elif re.match(r"(?i)\s*(SET\s|USE\s|GO$|/$|COMMIT|BEGIN\s*(TRAN|;)|PRINT\s|DECLARE\s)", sql):
                    n["tech"] = "変数の準備" if re.match(r"(?i)\s*(SET|DECLARE|USE)", sql) else "後片付け"
                nodes.append(n)
                buf, start = [], None
        if buf and start:
            nodes.append({"kind": "stmt", "lines": [start, len(m["lines"])], "exec": list(range(start, len(m["lines"]) + 1)), "text": " ".join(buf)[:120]})
        if nodes:
            procs[pid] = finish(pid, "（本体）", m["name"], "script", [nodes[0]["lines"][0], nodes[-1]["lines"][1]], nodes, comment_at, info.d)
            entries.append({"procedure": pid, "trigger": "script", "label": f"SQL スクリプト「{m['name']}」を実行", "evidence": "ファイルの中の文"})
    return procs, entries, "heuristic"


# ---------- バッチ・PowerShell・シェル（行の形から） ----------

LANGS = {
    "batch": {
        "comment": re.compile(r"(?i)^\s*(rem\b|::|@echo\s+off)"), "func": re.compile(r"^\s*:([\w\u0080-￿]+)\s*$"),
        "if": re.compile(r"(?i)^\s*if\s+(.+?)\s*\(\s*$"), "if1": re.compile(r"(?i)^\s*if\s+((?:not\s+)?(?:exist\s+\S+|errorlevel\s+\d+|defined\s+\S+|\S+\s*(?:==|equ|neq|lss|leq|gtr|geq)\s*\S+))\s+(.+)$"),
        "elif": None, "else": re.compile(r"(?i)^\s*\)\s*else\s*\(\s*$"), "loop": re.compile(r"(?i)^\s*for\s+(.+?)\s+do\s*\(\s*$"), "close": re.compile(r"^\s*\)\s*$"),
        "exit": re.compile(r"(?i)^\s*(exit(\s+/b)?(\s+\d+)?|goto\s+:?eof)\s*$"), "goto": re.compile(r"(?i)^\s*goto\s+:?([\w\u0080-￿]+)\s*$"),
        "call": re.compile(r"(?i)^\s*call\s+:([\w\u0080-￿]+)"), "ask": re.compile(r"(?i)^\s*(pause\b|set\s+/p\s+|choice\b)"), "log": re.compile(r"(?i)^\s*(echo\b|title\b|cls\b)"),
        "prep": re.compile(r"(?i)^\s*(set\s+(?!/p)|setlocal|endlocal|cd\b|pushd|popd|chcp)"),
        "io": [(re.compile(r"(?i)^\s*(?:copy|xcopy|robocopy|move)\s+(?:/\S+\s+)*(\S+)\s+(\S+)"), "ファイル", "pair"), (re.compile(r"(?i)^\s*(?:del|erase|rd|rmdir|md|mkdir|ren|rename)\s+(?:/\S+\s+)*(\S+)"), "ファイル", "out"),
               (re.compile(r"(?i)\b(?:sqlcmd|bcp|osql|sqlplus|mysql|psql)\b"), "データベース", "tool"), (re.compile(r"(?i)^\s*(?:start\s+(?:\"[^\"]*\"\s+)?|call\s+)?(\S+\.(?:exe|bat|cmd|vbs|ps1|py|jar))\b"), "外部プログラム", "out"),
               (re.compile(r">>?\s*(\S+)\s*$"), "ファイル", "out"), (re.compile(r"(?i)^\s*type\s+(\S+)"), "ファイル", "in")]},
    "powershell": {
        "comment": re.compile(r"^\s*#"), "func": re.compile(r"(?i)^\s*function\s+([\w\-\u0080-￿]+).*\{\s*$"),
        "if": re.compile(r"(?i)^\s*if\s*\((.+)\)\s*\{\s*$"), "if1": None, "elif": re.compile(r"(?i)^\s*\}?\s*elseif\s*\((.+)\)\s*\{\s*$"), "else": re.compile(r"(?i)^\s*\}?\s*else\s*\{\s*$"),
        "loop": re.compile(r"(?i)^\s*(?:foreach|for|while)\s*\((.+)\)\s*\{\s*$|^\s*(.+)\|\s*(?:ForEach-Object|%)\s*\{\s*$"), "close": re.compile(r"^\s*\}\s*$"),
        "try": re.compile(r"(?i)^\s*try\s*\{\s*$"), "catch": re.compile(r"(?i)^\s*\}?\s*(catch|finally)\b.*\{\s*$"),
        "exit": re.compile(r"(?i)^\s*(exit\b.*|return\b.*|throw\b.*)$"), "goto": None, "call": None, "ask": re.compile(r"(?i)\bRead-Host\b"),
        "log": re.compile(r"(?i)^\s*(Write-(Host|Output|Verbose|Debug|Information)|Start-Transcript|Stop-Transcript)\b"),
        "prep": re.compile(r"(?i)^\s*(\$ErrorActionPreference|Set-StrictMode|param\s*\(|Import-Module|Set-Location|cd\b)"),
        "io": [(re.compile(r"(?i)\b(?:Import-Csv|Get-Content|Import-Excel|Import-Clixml)\s+(?:-\w+\s+)*(\S+)"), "ファイル", "in"),
               (re.compile(r"(?i)\b(?:Export-Csv|Set-Content|Add-Content|Out-File|Export-Excel|Export-Clixml)\s+(?:-\w+\s+)*(\S+)"), "ファイル", "out"),
               (re.compile(r"(?i)\b(?:Copy-Item|Move-Item)\s+(?:-\w+\s+)*(\S+)\s+(?:-\w+\s+)*(\S+)"), "ファイル", "pair"), (re.compile(r"(?i)\b(?:Remove-Item|New-Item)\s+(?:-\w+\s+)*(\S+)"), "ファイル", "out"),
               (re.compile(r"(?i)\bInvoke-Sqlcmd\b"), "データベース", "tool"), (re.compile(r"(?i)\bSend-MailMessage\b"), "メール", "tool"),
               (re.compile(r"(?i)\b(?:Start-Process|&)\s+(?:-\w+\s+)*[\"']?([^\s\"']+)"), "外部プログラム", "out"), (re.compile(r"(?i)\bInvoke-(?:WebRequest|RestMethod)\b"), "外部プログラム", "tool")]},
    "shell": {
        "comment": re.compile(r"^\s*#"), "func": re.compile(r"^\s*(?:function\s+)?([\w\u0080-￿]+)\s*\(\)\s*\{?\s*$"),
        "if": re.compile(r"^\s*if\s+(.+?)\s*;?\s*(?:then)?\s*$"), "if1": None, "elif": re.compile(r"^\s*elif\s+(.+?)\s*;?\s*(?:then)?\s*$"), "else": re.compile(r"^\s*else\s*$"),
        "loop": re.compile(r"^\s*(?:for|while|until)\s+(.+?)\s*;?\s*(?:do)?\s*$"), "close": re.compile(r"^\s*(fi|done|\}|esac)\s*;?\s*$"),
        "exit": re.compile(r"^\s*(exit\b.*|return\b.*)$"), "goto": None, "call": None, "ask": re.compile(r"^\s*read\s+"), "log": re.compile(r"^\s*(echo|printf|logger)\b(?!.*>)"),
        "prep": re.compile(r"^\s*(set\s+-|export\s+|cd\s+|source\s+|\.\s+|[A-Za-z_]\w*=)"), "skip": re.compile(r"^\s*(then|do)\s*$"),
        "io": [(re.compile(r"^\s*(?:cp|mv|rsync|scp)\s+(?:-\S+\s+)*(\S+)\s+(\S+)"), "ファイル", "pair"), (re.compile(r"^\s*(?:rm|mkdir|touch|rmdir)\s+(?:-\S+\s+)*(\S+)"), "ファイル", "out"),
               (re.compile(r"\b(?:mysql|psql|sqlplus|sqlite3|bcp|sqlcmd)\b"), "データベース", "tool"), (re.compile(r"\b(?:mail|mailx|sendmail|mutt)\b"), "メール", "tool"),
               (re.compile(r"\b(?:curl|wget)\b"), "外部プログラム", "tool"), (re.compile(r"(?<![0-9&])>>?\s*([^\s&|;]+)"), "ファイル", "out"), (re.compile(r"(?<!<)<\s*([^\s&|;<]+)"), "ファイル", "in"),
               (re.compile(r"^\s*(?:cat|head|tail|grep\s+\S+|awk\s+'[^']*'|sort|wc\s+-\w+)\s+([^\s|>-][^\s|>]*)"), "ファイル", "in")]},
}


def clean_name(x):
    x = x.strip().strip("\"'")
    if re.search(r"[%$]", x):                                         # 変数が入った名前は、変数を ＜名前＞ にする
        t = re.sub(r"%~?\w*?([A-Za-z_]\w*)%|\$\{?([A-Za-z_]\w*)\}?|%%?~?\w", lambda mm: f"＜{mm.group(1) or mm.group(2) or '変数'}＞", x)
        tail = re.split(r"[\\/]", t)[-1]
        return (tail or t), not tail.strip("＜＞") or bool(re.fullmatch(r"＜[^＞]*＞", tail))
    return re.split(r"[\\/]", x)[-1] or x, False


class ScriptParser:
    """行の形から木を作る（バッチ・PowerShell・シェル）"""

    def __init__(self, m, lang):
        self.m, self.L, self.lang = m, LANGS[lang], lang
        self.rows = [(k, ln) for k, ln in enumerate(m["lines"], 1) if ln.strip() and not self.L["comment"].match(ln) and not (self.L.get("skip") and self.L["skip"].match(ln))]
        self.comment_at = {k: re.sub(r"(?i)^\s*(rem\s|::|#)\s*", "", ln).strip() for k, ln in enumerate(m["lines"], 1) if self.L["comment"].match(ln) and not ln.lower().strip().startswith("@echo")}
        self.funcs = {}                                               # 関数・ラベル → (開始の位置, 終了の位置)
        self.i = 0

    def stmt(self, k, ln, info):
        L, io, text = self.L, [], ln.strip()
        for rx, kind, way in L["io"]:
            mt = rx.search(text)
            if not mt:
                continue
            if way == "pair":
                for g, w in ((1, "in"), (2, "out")):
                    nm, dyn = clean_name(mt.group(g))
                    info.io(io, kind, nm, w, k, mt.group(g) if dyn else None)
            elif way == "tool":
                info.io(io, kind, mt.group(0).strip(), "out", k)
            else:
                nm, dyn = clean_name(mt.group(1))
                info.io(io, kind, nm, way, k, mt.group(1) if dyn else None)
            if kind != "ファイル" or way == "pair":
                break
        n = {"kind": "stmt", "lines": [k, k], "exec": [k], "text": text[:120]}
        if L["ask"] and L["ask"].search(text):
            n["kind"], n["interact"] = "interact", {"how": "inputbox", "text": cs.first_string(text)}
            info.d["inter"].append(dict(n["interact"], line=k))
        elif io:
            n["io"] = io
        elif L["log"].match(text):
            n["tech"] = "ログ"
        elif L["prep"].match(text):
            n["tech"] = "変数の準備"
        mt = L["call"].match(text) if L["call"] else None
        name = mt.group(1) if mt else next((w for w in re.findall(r"[\w\-\u0080-￿]+", text)[:1] if w in self.names and n["kind"] == "stmt" and not io), None)
        if name and name in self.names:
            n.update(kind="call", to=f"{self.m['name']}.{name}")
            n.pop("tech", None)
            info.d["calls"].append({"to": n["to"], "line": k, "head": False})
        return n

    def block(self, info):
        L, out = self.L, []
        while self.i < len(self.rows):
            k, ln = self.rows[self.i]
            if L["close"].match(ln) or (L["else"] and L["else"].match(ln)) or (L.get("elif") and L["elif"].match(ln)) or (L.get("catch") and L["catch"].match(ln)):
                break
            self.i += 1
            if L["func"].match(ln) and self.lang != "batch":
                depth = 1                                             # 関数の中身は、別の手続きとして読む（ここでは飛ばす）
                while self.i < len(self.rows) and depth:
                    t = self.rows[self.i][1]
                    depth += bool(re.search(r"\{\s*$", t)) - bool(L["close"].match(t))
                    self.i += 1
                continue
            mt = L["if"].match(ln)
            if mt and not (self.lang == "shell" and ln.strip().startswith("elif")):
                node = {"kind": "if", "head": [k, k], "exec": [k], "cond": mt.group(1), "branches": []}
                label, cond, head = "then", mt.group(1), [k, k]
                while True:
                    body = self.block(info)
                    node["branches"].append({"label": label, "cond": cond, "head": head, "body": body})
                    if self.i >= len(self.rows):
                        break
                    k2, t = self.rows[self.i]
                    self.i += 1
                    if L.get("elif") and L["elif"].match(t):
                        label, cond, head = "elseif", L["elif"].match(t).group(1), [k2, k2]
                        node["exec"].append(k2)
                    elif L["else"] and L["else"].match(t):
                        label, cond, head = "else", "", [k2, k2]
                    else:
                        break
                node["lines"] = [k, self.rows[self.i - 1][0]]
                out.append(node)
                continue
            mt = L["if1"].match(ln) if L["if1"] else None
            if mt:
                inner = self.stmt(k, mt.group(2), info)
                if L["exit"].match(mt.group(2)):
                    inner = {"kind": "exit", "what": "proc", "label": None, "lines": [k, k], "exec": []}
                elif L["goto"] and L["goto"].match(mt.group(2)):
                    inner = {"kind": "exit", "what": "goto", "label": L["goto"].match(mt.group(2)).group(1), "lines": [k, k], "exec": []}
                inner["exec"] = []
                out.append({"kind": "if", "single": True, "head": [k, k], "lines": [k, k], "exec": [k], "cond": mt.group(1),
                            "branches": [{"label": "then", "cond": mt.group(1), "head": [k, k], "body": [inner]}]})
                continue
            mt = L["loop"].match(ln)
            if mt:
                body = self.block(info)
                end = self.rows[self.i][0] if self.i < len(self.rows) else k
                self.i += 1
                out.append({"kind": "loop", "form": "for", "cond": next((g for g in mt.groups() if g), "").strip(), "head": [k, k], "lines": [k, end], "exec": [k], "body": body})
                continue
            if L.get("try") and L["try"].match(ln):
                out += self.block(info)
                while self.i < len(self.rows) and L["catch"].match(self.rows[self.i][1]):
                    k2, t = self.rows[self.i]
                    self.i += 1
                    body = self.block(info)
                    if "finally" in t.lower():
                        out += body
                    else:
                        out.append({"kind": "if", "head": [k2, k2], "lines": [k2, self.rows[min(self.i, len(self.rows) - 1)][0]], "exec": [k2], "cond": "エラーが起きたとき",
                                    "branches": [{"label": "then", "cond": "", "head": [k2, k2], "body": body}]})
                self.i += 1 if self.i < len(self.rows) and L["close"].match(self.rows[self.i][1]) else 0
                continue
            if L["exit"].match(ln):
                out.append({"kind": "exit", "what": "proc", "label": None, "lines": [k, k], "exec": [k]})
                continue
            mt = L["goto"].match(ln) if L["goto"] else None
            if mt:
                out.append({"kind": "exit", "what": "goto", "label": mt.group(1), "lines": [k, k], "exec": [k]})
                continue
            mt = L["func"].match(ln) if self.lang == "batch" else None
            if mt:
                out.append({"kind": "label", "name": mt.group(1), "lines": [k, k]})
                continue
            out.append(self.stmt(k, ln, info))
        return out

    def parse(self):
        m, L, procs = self.m, self.L, {}
        spans = []                                                    # 関数の範囲（バッチは、call :ラベル で呼ばれるラベルから、次の exit /b・goto :eof まで）
        lines = m["lines"]
        if self.lang == "batch":
            called = {mt.group(1).lower() for ln in lines for mt in [L["call"].match(ln)] if mt}
            for k, ln in enumerate(lines, 1):
                mt = L["func"].match(ln)
                if mt and mt.group(1).lower() in called:
                    end = next((j for j in range(k + 1, len(lines) + 1) if L["exit"].match(lines[j - 1]) or (L["func"].match(lines[j - 1]) and j > k + 1)), len(lines))
                    end = end if L["exit"].match(lines[end - 1]) else end - 1 if end < len(lines) else end
                    spans.append((mt.group(1), k, end))
        else:
            for k, ln in enumerate(lines, 1):
                mt = L["func"].match(ln)
                if mt:
                    depth, j = 1, k
                    while j < len(lines) and depth:
                        j += 1
                        t = lines[j - 1]
                        depth += bool(re.search(r"\{\s*$", t)) - bool(re.match(r"^\s*\}\s*$", t))
                    spans.append((mt.group(1), k, j))
        self.names = {nm for nm, _a, _b in spans}
        inside = lambda k: any(a <= k <= b for _n, a, b in spans)
        every = self.rows
        for nm, a, b in spans:
            self.rows, self.i = [(k, ln) for k, ln in every if a < k <= b and not (self.lang != "batch" and k == b)], 0
            info = Info(f"{m['name']}.{nm}")
            nodes = []
            while self.i < len(self.rows):
                nodes += self.block(info)
                self.i += 1
            procs[info.pid] = finish(info.pid, nm, m["name"], "function", [a, b], nodes, self.comment_at, info.d)
        self.rows, self.i = [(k, ln) for k, ln in every if not inside(k)], 0
        info, nodes = Info(f"{m['name']}.（本体）"), []
        while self.i < len(self.rows):
            nodes += self.block(info)
            self.i += 1
        entries = []
        if nodes:
            procs[info.pid] = finish(info.pid, "（本体）", m["name"], "script", [self.rows[0][0], self.rows[-1][0]], nodes, self.comment_at, info.d)
            entries.append({"procedure": info.pid, "trigger": "script", "label": f"スクリプト「{m['name']}」を実行", "evidence": "関数の外に書かれた行（スクリプトの本体）"})
        return procs, entries


def parse_script(modules, lang):
    procs, entries = {}, []
    for m in modules:
        p, e = ScriptParser(m, lang).parse()
        procs.update(p)
        entries += e
    return procs, entries, "heuristic"


# ---------- 構造を LLM が読む言語（Java・C#・COBOL ほか） ----------
# LLM が書くのは、ブロックの木（種類と行の範囲）だけ: code/<まとまり>.tree.json。機械が、範囲・入れ子・行の網羅・見出しの語を確かめて、構造にする。

C_LIKE = ("java", "csharp", "javascript", "typescript", "c", "cpp", "go", "kotlin", "php")
HEAD_WORDS = {"if": r"(?i)\b(if|unless|when|catch|except|rescue|on\s+error|at\s+end|invalid\s+key|on\s+exception|on\s+size\s+error)\b|\?",
              "select": r"(?i)\b(switch|select|case|evaluate|when|match)\b",
              "loop": r"(?i)\b(for|foreach|while|do|loop|perform|until|each|repeat|times)\b|\.forEach\b|\.map\b"}
EXIT_WORDS = r"(?i)\b(return|break|continue|exit|goto|go\s+to|throw|raise|stop\s+run|goback|die|next|leave|abort|abend|cee3abd|ilbo?abn0|panic|halt)\b"
NO_OP = re.compile(r"(?i)^(continue|next\s+sentence)\s*\.?$")           # COBOL の、何もしない文（どのブロックにも入れなくてよい）
STRUCTURAL = re.compile(r"(?i)^[\s{}()\[\];,.]*$|^\s*\}?\s*(else|try|finally|do|begin|then)\s*[{:.]?\s*$|"
                        r"^\s*(end|end\s+\w+|end-\w+|fi|done|esac|next(\s+\w+)?|loop|wend)\s*[;.]?\s*$")


def plain_lines(lang, lines):
    """空の行・コメントだけの行・かっこや end だけの行を除いた、行の番号の集合（どれかのノードに入っているべき行）と、コメントだけの行 → 文字"""
    keep, notes, block = set(), {}, False
    marks = ("//", "#") if lang == "php" else ("//",) if lang in C_LIKE else ("'", "rem ") if lang == "vbnet" else ("//*",) if lang == "jcl" \
        else ("#",) if lang in ("ruby", "perl") else ("*>",) if lang == "cobol" else ("//", "#", "--")
    for k, raw in enumerate(lines, 1):
        s = raw.strip()
        if lang == "cobol" and len(raw) > 6 and raw[:6].strip(" 0123456789") == "":   # 固定形式: 1〜6 桁目は行の番号、7 桁目が * か / なら注釈行
            if raw[6] in "*/":
                notes[k] = raw[7:].strip()
                continue
            s = raw[7:].strip()
        if block:
            block = "*/" not in s
            notes[k] = s.strip("*/ ").strip()
            continue
        if s.startswith("/*") and lang != "cobol":
            block = "*/" not in s
            notes[k] = s.strip("/* ").strip()
            continue
        if not s:
            continue
        if s.lower().startswith(marks):
            notes[k] = re.sub(r"^(//\*?|#|--|'|rem\s|\*>)\s*", "", s, flags=re.I).strip()
            continue
        if not STRUCTURAL.match(s) and not (lang == "cobol" and NO_OP.match(s)):
            keep.add(k)
    return keep, notes


def comment_hints(notes):
    """コメントだけの行 → 名付けの手がかりにする文字（飾りの線だけの行は、上に続くコメントの文字を使う。文字が無ければ空）"""
    out = {}
    for k in notes:
        j = k
        while j in notes and not re.search(r"[0-9A-Za-z\u3040-\u30ff\u4e00-\u9fff]", notes[j]):
            j -= 1
        out[k] = re.sub(r"^[\s*\-=#/]+|[\s*\-=#/]+$", "", notes[j]) if j in notes else ""
    return out


def tree_nodes(nodes):
    for n in nodes or []:
        if isinstance(n, dict):
            yield n
            for b in n.get("branches") or []:
                yield from tree_nodes(b.get("body") if isinstance(b, dict) else [])
            yield from tree_nodes(n.get("body"))


def tree_problems(tree, modules, lang):
    """LLM が書いた木の誤り。空なら、構造にできる"""
    errs, mods = [], {m["name"]: m for m in modules}
    span = lambda x: isinstance(x, list) and len(x) == 2 and all(isinstance(v, int) for v in x) and 1 <= x[0] <= x[1]
    ids, used = set(), {}
    all_ids = {f"{q.get('module')}.{q.get('name')}": q.get("module") for q in tree.get("procedures") or []}
    for p in tree.get("procedures") or []:
        pid = f"{p.get('module')}.{p.get('name')}"
        m = mods.get(p.get("module"))
        if m is None:
            errs.append(site(f"{pid}: モジュール「{p.get('module')}」が、このまとまりに無い（あるのは {'・'.join(list(mods)[:6])}）"))
            continue
        if not span(p.get("lines")) or p["lines"][1] > len(m["lines"]):
            errs.append(site(f"{pid}: 行の範囲 {p.get('lines')} が、ソース（{len(m['lines'])} 行）に合わない"))
            continue
        if pid in ids:
            errs.append(site(f"{pid}: 同じ名前の手続きが2つある（引数の違うものは、名前に引数を添えて分ける）"))
        ids.add(pid)
        lo, hi = p["lines"]
        for q, (x, y) in used.get(p["module"], []):
            if lo <= y and x <= hi:
                errs.append(site(f"{pid}: 行の範囲（{lo}〜{hi}）が、{q}（{x}〜{y}）と重なる"))
        used.setdefault(p["module"], []).append((pid, (lo, hi)))
        keep, _notes = plain_lines(lang, m["lines"])
        text = lambda a, b: " ".join(x.strip() for x in m["lines"][a - 1:b])
        head = p.get("head") or [lo, lo]
        owned = {}

        def own(a, b, what):
            for ln in range(a, b + 1):
                if ln in keep:
                    owned.setdefault(ln, []).append(what)

        def visit(nodes, a, b, where):
            last = a - 1
            for n in nodes:
                k = n.get("kind")
                if not span(n.get("lines")):
                    errs.append(site(f"{pid}: {where}のノードの行の範囲が無い・違う（{n.get('lines')}）"))
                    continue
                x, y = n["lines"]
                tag = f"{k}（{x}〜{y} 行）"
                if not (a <= x and y <= b):
                    errs.append(site(f"{pid}: {tag}が、{where}の範囲（{a}〜{b} 行）の外"))
                    continue
                if x <= last:
                    errs.append(site(f"{pid}: {tag}が、前のノードと重なっているか、順が逆"))
                last = y
                if k in ("if", "select", "loop"):
                    if n.get("leave") is not None:
                        errs.append(site(f"{pid}: {tag}に leave は付けられない（seq・call・interact だけ）"))
                    h = n.get("head") or [x, x]
                    if not span(h) or not (x <= h[0] and h[1] <= y):
                        errs.append(site(f"{pid}: {tag}の head {h} が、ノードの範囲の外"))
                        continue
                    own(h[0], h[1], tag)
                    heads = [h] + [br["head"] for br in n.get("branches") or [] if span(br.get("head"))] + ([n["tail"]] if span(n.get("tail")) else [])
                    for to in n.get("calls") or []:
                        if to not in all_ids:
                            errs.append(site(f"{pid}: {tag}の calls「{to}」が、木の手続きに無い"))
                        elif not any(str(to).split(".")[-1].split("(")[0].lower() in text(*hh).lower() for hh in heads):
                            errs.append(site(f"{pid}: {tag}の条件の行に、calls の名前「{str(to).split('.')[-1]}」が無い"))
                    if not re.search(HEAD_WORDS[k], text(*h)):
                        errs.append(site(f"{pid}: {tag}の先頭の行（{h[0]} 行）に、{ {'if': '分岐', 'select': '振り分け', 'loop': '繰り返し'}[k] }の語が無い: {text(*h)[:60]}"))
                    if k == "loop":
                        t = n.get("tail")
                        if t is not None:
                            if span(t) and h[1] < t[0] and t[1] <= y:
                                own(t[0], t[1], tag)
                            else:
                                errs.append(site(f"{pid}: {tag}の tail {t} が、ノードの範囲の外"))
                        visit(n.get("body") or [], h[1] + 1, (t[0] - 1) if span(t) else y, "繰り返し")
                    else:
                        if not n.get("branches"):
                            errs.append(site(f"{pid}: {tag}に、出口（branches）が無い"))
                        at = h[1] + 1
                        for br in n.get("branches") or []:
                            bh = br.get("head")
                            if bh is not None:
                                if not span(bh) or not (at <= bh[0] and bh[1] <= y):
                                    errs.append(site(f"{pid}: {tag}の出口「{br.get('label')}」の head {bh} が、順に並んでいない・範囲の外"))
                                    continue
                                own(bh[0], bh[1], tag)
                                at = bh[1] + 1
                            body = br.get("body") or []
                            visit(body, at, y, f"出口「{br.get('label')}」")
                            at = max([at] + [c["lines"][1] + 1 for c in body if span(c.get("lines"))])
                else:
                    own(x, y, tag)
                    t = text(x, y)
                    if k == "exit" and not re.search(EXIT_WORDS, t):
                        errs.append(site(f"{pid}: {tag}に、抜ける・戻る語（return・break・throw など）が無い: {t[:60]}"))
                    if k == "exit" and n.get("io"):
                        errs.append(site(f"{pid}: {tag}に io は付けられない（読む・書くものがある return は、seq にして \"leave\": \"proc\" を添える）"))
                    if n.get("calls"):
                        errs.append(site(f"{pid}: {tag}に calls は付けられない（if・select・loop の条件の行だけ。ほかは call のブロックにする）"))
                    if n.get("leave") is not None and (k == "exit" or not re.search(EXIT_WORDS, t)):
                        errs.append(site(f"{pid}: {tag}の leave: 行に、抜ける・戻る語（return など）が無い・exit のブロックには要らない: {t[:60]}"))
                    if k == "interact" and n.get("text") and re.sub(r"\s+", " ", n["text"]).strip().lower() not in t.lower():
                        errs.append(site(f"{pid}: {tag}の text「{n['text'][:30]}」が、その行の文字に無い"))
                    if k == "call":
                        to = str(n.get("to") or "")
                        if to not in all_ids:
                            errs.append(site(f"{pid}: {tag}の呼び先「{to}」が、木の手続きに無い（範囲の外の呼び出しは、call でなく seq にする）"))
                        elif to.split(".")[-1].split("(")[0].lower() not in t.lower():
                            errs.append(site(f"{pid}: {tag}の行に、呼び先の名前「{to.split('.')[-1]}」が無い: {t[:60]}"))
                for io in n.get("io") or []:
                    ln = io.get("line")
                    if not isinstance(ln, int) or not (1 <= ln <= len(m["lines"])):   # ふつうはノードの中の行。名前が定数なら、その定数を決めている行でもよい
                        errs.append(site(f"{pid}: {tag}の入出力「{io.get('name')}」の行（{ln}）が、ソースに無い"))
                    elif io.get("name") != cs.DYNAMIC and str(io.get("name")).lower() not in m["lines"][ln - 1].lower():
                        errs.append(site(f"{pid}: {tag}の入出力の名前「{io.get('name')}」が、{ln} 行の文字に無い（実行するときに決まる名前は「{cs.DYNAMIC}」と書く）"))
        visit(p.get("body") or [], head[1] + 1 if span(head) else lo, hi, "手続き")
        lost = sorted(ln for ln in keep if head[1] < ln <= hi and ln not in owned)
        twice = sorted(ln for ln, w in owned.items() if len(w) > 1)
        if lost:
            errs.append(site(f"{pid}: どのノードにも入っていない行がある（{len(lost)} 行: {', '.join(map(str, lost[:8]))}{' …' if len(lost) > 8 else ''}）。"
                        f"先頭は {lost[0]} 行「{m['lines'][lost[0] - 1].strip()[:40]}」"))
        if twice:
            errs.append(site(f"{pid}: 2つのノードに入っている行がある（{', '.join(map(str, twice[:8]))}）"))
    for e in tree.get("entries") or []:
        pid = str(e.get("procedure"))
        if pid not in ids:
            errs.append(site(f"起動点「{e.get('label')}」の手続き「{pid}」が、木の手続きに無い"))
            continue
        m = mods[all_ids[pid]]
        ln = e.get("line")
        if not isinstance(ln, int) or not (1 <= ln <= len(m["lines"])):
            errs.append(site(f"起動点「{e.get('label')}」の根拠の行（{ln}）が、ソースに無い"))
            continue
        also = e.get("also")
        if also is not None and not (isinstance(also, int) and 1 <= also <= len(m["lines"])):
            errs.append(site(f"起動点「{e.get('label')}」の also の行（{also}）が、ソースに無い"))
            also = None
        here, front = (m["lines"][ln - 1] + " " + m["name"]).lower(), m["lines"][also - 1].lower() if also else ""
        for name in re.findall(r"「([^」]+)」", str(e.get("label") or "")):
            low = name.lower()                                       # also があれば、その行の文字と根拠の行の文字を、この順につないだ名前でもよい（クラスに付いた URL の前置き ＋ メソッドの URL）
            if low not in here and not any(low[:i] in front and low[i:] in here for i in range(1, len(low)) if front):
                errs.append(site(f"起動点「{e.get('label')}」: かぎかっこの中の名前「{name}」が、根拠の行（{m['name']} {ln} 行）にも、ファイルの名前にも無い"
                            + ("（also の行の文字と、根拠の行の文字を、この順につないだものでもない）" if also else "（名前の前半が別の行にあるときは、その行を also に書く）")))
        for b in e.get("before") or []:
            if b not in ids or b == pid:
                errs.append(site(f"起動点「{e.get('label')}」の before「{b}」が、木の手続きに無い（「モジュール.手続き」で書く。起動点の手続きそのものは書けない）"))
    return errs


def from_tree(tree, modules, lang):
    """LLM が書いた木（誤りの無いもの）→ (手続き, 起動点)。条件の文字・実行文の行・コメントの手がかりは、機械がソースから入れる"""
    mods = {m["name"]: m for m in modules}
    procs, entries, before = {}, [], {}
    for e in tree.get("entries") or []:                              # 枠組みが、起動点の手続きの前に呼ぶ手続き（図の先頭に、呼び出しとして入れる）
        before.setdefault(e["procedure"], []).extend(b for b in e.get("before") or [] if b not in before.get(e["procedure"], []))
    for p in tree.get("procedures") or []:
        m = mods[p["module"]]
        pid, info = f"{p['module']}.{p['name']}", Info(f"{p['module']}.{p['name']}")
        keep, notes = plain_lines(lang, m["lines"])
        body = lambda x: x[7:] if lang == "cobol" and len(x) > 6 and x[:6].strip(" 0123456789") == "" else x     # COBOL の固定形式: 行の番号の欄を除く
        text = lambda a, b: re.sub(r"\s+", " ", " ".join(body(x).strip() for x in m["lines"][a - 1:b])).strip()
        mean = lambda a, b: [ln for ln in range(a, b + 1) if ln in keep]
        quoted = lambda t: cs.first_string(t) or next(iter(re.findall(r"'([^']+)'", t)), "")

        def conv_all(nodes):
            out = []
            for n in nodes or []:
                x, y = n["lines"]
                cuts = [k for k in range(x + 1, y) if k in notes] if n["kind"] == "seq" and all(x <= i["line"] <= y for i in n.get("io") or []) else []
                parts, a = [], x                                     # まとまりの中の、コメントだけの行で分ける（機械が構文を読む言語と同じ区切り）
                for k in cuts + [y + 1]:
                    if mean(a, k - 1):
                        parts.append(dict(n, lines=[a, k - 1], io=[i for i in n.get("io") or [] if a <= i["line"] <= k - 1]))
                    a = k + 1
                for part in parts if cuts else [n]:
                    node = conv(part)
                    if node["kind"] == "stmt" and not node["exec"]:   # 何もしない文だけ（COBOL の CONTINUE）
                        continue
                    out.append(node)
                    if n.get("leave") and part is (parts[-1] if cuts else n):   # 1行が「する」と「戻る」を兼ねる（行は、先のノードが持つ）
                        out.append({"kind": "exit", "what": n["leave"], "label": None, "lines": [part["lines"][1], part["lines"][1]], "exec": []})
            return out

        def conv(n):
            x, y = n["lines"]
            k = n["kind"]
            io = []
            for i in n.get("io") or []:
                info.io(io, i["kind"], i["name"], i["way"], i["line"], m["lines"][i["line"] - 1].strip() if i["name"] == cs.DYNAMIC else None)
            if k in ("if", "select"):
                h = n.get("head") or [x, x]
                node = {"kind": k, "lines": [x, y], "head": h, "exec": mean(*h), "cond": text(*h)[:160], "branches": []}
                for br in n["branches"]:
                    bh = br.get("head")
                    if bh:
                        node["exec"] += mean(*bh)
                    node["branches"].append({"label": br["label"], "cond": text(*bh)[:160] if bh else node["cond"] if br["label"] == "then" else "",
                                             "head": bh or h, "body": conv_all(br.get("body"))})
            elif k == "loop":
                h, t = n.get("head") or [x, x], n.get("tail")
                node = {"kind": "loop", "form": "", "lines": [x, y], "head": h, "exec": mean(*h) + (mean(*t) if t else []), "cond": text(*h)[:160],
                        "body": conv_all(n.get("body"))}
            elif k == "call":
                node = {"kind": "call", "lines": [x, y], "exec": mean(x, y), "to": n["to"], "text": text(x, y)[:120]}
                info.d["calls"].append({"to": n["to"], "line": x, "head": False})
            elif k == "interact":
                node = {"kind": "interact", "lines": [x, y], "exec": mean(x, y), "text": text(x, y)[:120],
                        "interact": {"how": n.get("how", "msgbox"), "text": re.sub(r"\s+", " ", n["text"]).strip() if n.get("text") else quoted(text(x, y))}}
                info.d["inter"].append(dict(node["interact"], line=x))
            elif k == "exit":
                node = {"kind": "exit", "what": n.get("what", "proc"), "label": None, "lines": [x, y], "exec": mean(x, y)}   # select: 振り分けの出口の終わり（switch の break）
            else:                                                    # まとまりは、機械が構文を読む言語と同じ決まりでつなぎ直す（group: コメントの行・触る先・技術的な文の境で切る）
                node = {"kind": "stmt", "lines": [x, y], "exec": mean(x, y), "text": text(x, y)[:120]}
                if n.get("tech"):
                    node["tech"] = n["tech"]
            if io:
                node["io"] = io
            for to in n.get("calls") or [] if k in ("if", "select", "loop") else []:   # 条件の中の呼び出し
                info.d["calls"].append({"to": to, "line": (n.get("head") or [x, x])[0], "head": True})
            return node
        nodes = conv_all(p.get("body"))
        for b in reversed(before.get(pid, [])):                      # 行は、手続きの先頭の行（実行文は持たない）
            nodes.insert(0, {"kind": "call", "lines": [p["lines"][0]] * 2, "exec": [], "to": b, "text": "", "framework": True})
            info.d["calls"].insert(0, {"to": b, "line": p["lines"][0], "head": False})
        procs[pid] = finish(pid, p["name"], p["module"], "function", list(p["lines"]), nodes, comment_hints(notes), info.d)
    for e in tree.get("entries") or []:
        entries.append({"procedure": e["procedure"], "trigger": e.get("trigger", "entry"), "label": e["label"], "evidence": f"{procs[e['procedure']]['module']} {e['line']} 行"})
    return procs, entries


def tree_agreement(t1, t2):
    """同じソースを2回読んだ木の一致: {"procedures": [一致, 全体], "nodes": [一致, 全体], "rate": 割合}。ノードは（手続き・種類・行の範囲）で比べる"""
    def joined(nodes):
        out = []
        for n in nodes or []:
            n = dict(n, **({"branches": [dict(b, body=joined(b.get("body"))) for b in n["branches"]]} if n.get("branches") else {}),
                     **({"body": joined(n["body"])} if n.get("body") else {}))
            if n.get("kind") == "seq" and out and out[-1].get("kind") == "seq" and out[-1].get("tech") == n.get("tech") and not out[-1].get("leave"):
                out[-1] = dict(out[-1], lines=[out[-1]["lines"][0], n["lines"][1]], leave=n.get("leave"))
            else:
                out.append(n)
        return out

    def keys(t, join):
        ps = {f"{p['module']}.{p['name']}": p for p in t.get("procedures") or []}
        return set(ps), {(pid, n.get("kind"), tuple(n.get("lines") or [])) for pid, p in ps.items() for n in tree_nodes(joined(p.get("body")) if join else p.get("body"))}
    (p1, n1), (p2, n2) = keys(t1, True), keys(t2, True)
    r1, r2 = keys(t1, False)[1], keys(t2, False)[1]
    flow = lambda ns: {k for k in ns if k[1] != "seq"}
    both, every = len(n1 & n2), len(n1 | n2)
    return {"procedures": [len(p1 & p2), len(p1 | p2)], "nodes": [both, every], "rate": round(both / every, 3) if every else 1.0,
            "raw": [len(r1 & r2), len(r1 | r2)], "flow": [len(flow(n1) & flow(n2)), len(flow(n1) | flow(n2))],
            "only_first": sorted(n1 - n2)[:20], "only_second": sorted(n2 - n1)[:20]}


def hints(lang, lines):
    """手続きの見出しらしい行（LLM が読む範囲を決める手がかり。構造には使わない）: [(行, 文字)]"""
    if lang == "cobol":
        rx = re.compile(r"(?i)^.{6}\s{1,4}([\w-]+)(\s+SECTION)?\.\s*$|^.{6}\s+(PROCEDURE\s+DIVISION|PROGRAM-ID)\b")
    else:
        rx = re.compile(r"(?i)^\s*(?:(?:public|private|protected|internal|static|final|abstract|async|override|virtual|synchronized|export|default)\s+)*"
                        r"(?:function\s+\w+|def\s+\w+|sub\s+\w+|func\s+\w+|fun\s+\w+|[\w<>\[\],.?]+\s+\w+)\s*\([^;]*\)?\s*(?:throws\s+[\w., ]+)?\s*[{:]?\s*$")
    skip = re.compile(r"(?i)^\s*(if|for|foreach|while|switch|catch|return|else|using|lock|new|throw)\b")
    return [(k, ln.strip()[:100]) for k, ln in enumerate(lines, 1) if rx.match(ln) and not skip.match(ln)]


def parse(language, modules):
    """(手続き, 起動点, 構造の段) か None（機械で読めない。LLM の読取に回す）"""
    if language == "python":
        return parse_python(modules)
    if language == "sql":
        return parse_sql(modules)
    if language in LANGS:
        return parse_script(modules, language)
    return None
