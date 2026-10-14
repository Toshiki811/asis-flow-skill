#!/usr/bin/env python3
"""coverage.py — 点検（check_flow・check_layout・check_extract）の FAIL のうち、自己点検で一度も出ないものを数える。

  python3 evals/coverage.py -w <自己点検の作業フォルダ>     # 自己点検を、FAIL を出した場所を控えながら回して、数える（自己点検と同じ時間）
  python3 evals/coverage.py --log <控え>                   # 控え（ファイル:行 の並び）から数えるだけ

点検の場所は、スクリプトの中の .fail( の呼び出し（構文から拾う）。控えは、環境変数 ASIS_FAIL_LOG があるとき、Report.fail が書き足す。
試験しない点検は、evals/untested_checks.md の表に、理由つきで並べる（ファイル・関数・文の一部で指す。行の番号は書かない）。
一覧に無いのに一度も出ない点検があれば、終了コード 1（試験を足すか、試験しない理由を書いて、一覧に足す）。
一覧にあるのに出た点検・一覧のどの点検にも当たらない行も知らせる（一覧を直す）。
"""
import argparse
import ast
import collections
import os
import pathlib
import re
import subprocess
import sys
import tempfile

SKILL = pathlib.Path(__file__).resolve().parent.parent
FILES = ("check_flow.py", "check_layout.py", "check_extract.py")
LIST = SKILL / "evals" / "untested_checks.md"
ENV = "ASIS_FAIL_LOG"                                                 # common.FAIL_LOG と同じ名前


def text_of(node):
    """FAIL の文の、決まった部分（変わる所は …）"""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        return "".join(v.value if isinstance(v, ast.Constant) else "…" for v in node.values)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
        return text_of(node.left) + text_of(node.right)
    if isinstance(node, ast.IfExp):
        return text_of(node.body)
    return "…"


def sites(root=SKILL):
    """点検の場所: [{file, func, line, end, text}]"""
    out = []
    for name in FILES:
        tree = ast.parse((root / "scripts" / name).read_text(encoding="utf-8"))
        funcs = [(n.lineno, n.end_lineno, n.name) for n in ast.walk(tree) if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
        for n in ast.walk(tree):
            if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute) and n.func.attr == "fail" and n.args:
                func = min((f for f in funcs if f[0] <= n.lineno <= f[1]), key=lambda f: f[1] - f[0], default=(0, 0, "（関数の外）"))[2]
                out.append({"file": name, "func": func, "line": n.lineno, "end": n.end_lineno, "text": re.sub(r"\s+", " ", text_of(n.args[0]))})
    return sorted(out, key=lambda s: (s["file"], s["line"]))


def listed(path=LIST):
    """試験しない点検の一覧: [{file, func, part, why}]（表の行）"""
    out = []
    for ln in path.read_text(encoding="utf-8").splitlines() if path.exists() else []:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) == 4 and cells[0].endswith(".py") and cells[0] in FILES:
            out.append({"file": cells[0], "func": cells[1].strip("`"), "part": cells[2].strip("「」"), "why": cells[3]})
    return out


def matches(row, s):
    return row["file"] == s["file"] and row["func"] == s["func"] and row["part"] in s["text"]


def tally(hits, root=SKILL, path=LIST):
    """(場所, 出た場所, 一覧の行, 出ない場所で一覧に無いもの, 一覧にあるのに出た場所, どの場所にも当たらない一覧の行)"""
    ss, rows = sites(root), listed(path)
    fired = [s for s in ss if any(f == s["file"] and s["line"] <= ln <= s["end"] for f, ln in hits)]
    quiet = [s for s in ss if s not in fired]
    loose = [s for s in quiet if not any(matches(r, s) for r in rows)]
    noisy = [s for s in fired if any(matches(r, s) for r in rows)]
    stale = [r for r in rows if not any(matches(r, s) for s in ss)]
    return ss, fired, rows, loose, noisy, stale


def read_log(path):
    hits = collections.Counter()
    for ln in pathlib.Path(path).read_text(encoding="utf-8").splitlines():
        f, _, n = ln.strip().rpartition(":")
        if f and n.isdigit():
            hits[(f, int(n))] += 1
    return hits


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("-w", "--work", help="自己点検の作業フォルダ（evals/selftest.py -w と同じ）")
    g.add_argument("--log", help="FAIL を出した場所の控え")
    a = ap.parse_args()
    if a.work:
        log = pathlib.Path(tempfile.mkstemp(prefix="asis_fail_", suffix=".log")[1])
        r = subprocess.run([sys.executable, str(SKILL / "evals" / "selftest.py"), "-w", a.work], env=dict(os.environ, **{ENV: str(log)}),
                           capture_output=True, text=True)
        said = log.with_suffix(".selftest.txt")
        said.write_text(r.stdout + r.stderr, encoding="utf-8")
        print(f"自己点検: 終了コード {r.returncode}（{r.stdout.strip().splitlines()[-1] if r.stdout.strip() else '出力なし'}）。控え: {log}・自己点検の出力: {said}")
        for ln in [x for x in r.stdout.splitlines() if x.startswith("NG")][:10]:
            print(f"  {ln[:150]}")
    else:
        log = pathlib.Path(a.log)
    ss, fired, rows, loose, noisy, stale = tally(read_log(log))
    print(f"点検の FAIL {len(ss)} か所: 自己点検で出た {len(fired)}・試験しない（{LIST.name}）{len(ss) - len(fired) - len(loose)}・出ないのに一覧に無い {len(loose)}")
    for s in loose:
        print(f"  出ない  {s['file']}:{s['line']} [{s['func']}] {s['text'][:110]}")
    for s in noisy:
        print(f"  一覧にあるのに出た  {s['file']}:{s['line']} [{s['func']}] {s['text'][:80]}（一覧から外す）")
    for r in stale:
        print(f"  一覧の行が、どの点検にも当たらない  {r['file']} {r['func']}「{r['part']}」（点検の文・関数の名前が変わった。一覧を直す）")
    sys.exit(1 if loose or stale else 0)


if __name__ == "__main__":
    main()
