#!/usr/bin/env python3
"""coverage.py — 比較の点検（sc_check・sc_layout）の FAIL のうち、自己点検で一度も出ないものを数える。

  python3 evals/coverage.py                       # 自己点検を、FAIL を出した場所を控えながら回して、数える（自己点検と同じ時間）
  python3 evals/coverage.py -w <自己点検の作業>    # 自己点検の組み立ての場所（selftest.py -w と同じ。あれば組み立てを飛ばす）
  python3 evals/coverage.py --log <控え>          # 控え（ファイル:行 の並び）から数えるだけ

点検の場所（構文から拾う）:
  FAIL     .fail( と .fail_if( の呼び出し。配置の点検（asis の check_layout.page_problems）の FAIL を中継する行（at= のあるもの）は数えない
           （中身は asis の自己点検が試す）
  呼び元   形の点検の部品 shape(…) を呼ぶ所（ファイルの種類ごとに、形の FAIL が出るかを見る）
控えは、環境変数 ASIS_FAIL_LOG があるとき、asis の Report.fail・fail_if（FAIL のときだけ）・common.site が書き足す（shape は呼び元の行も控える）。
数えないもの: 落ちないための守り（asis の common.guarded）。
試験しない点検は、evals/untested_checks.md の表に、理由つきで並べる（ファイル・関数・文の一部で指す。行の番号は書かない）。
終了コード 1: 一覧に無いのに一度も出ない点検がある（試験を足すか、試験しない理由を書いて一覧に足す）／一覧にあるのに出た点検・
どの点検にも当たらない一覧の行がある（一覧を直す）／自己点検が NG。
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
FILES = ("sc_check.py", "sc_layout.py")
HELPERS = ("shape",)
LIST = SKILL / "evals" / "untested_checks.md"
FAIL_LOG = "ASIS_FAIL_LOG"


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


def site_of(n):
    """構文の節が点検の場所なら (種類, 文の節)。種類は fail（FAIL）／call（形の点検の呼び元）"""
    if not isinstance(n, ast.Call):
        return None
    f = n.func
    name = f.attr if isinstance(f, ast.Attribute) else f.id if isinstance(f, ast.Name) else None
    if name == "fail" and n.args and isinstance(f, ast.Attribute):
        if any(k.arg == "at" for k in n.keywords):
            return None
        return "fail", n.args[0]
    if name == "fail_if" and len(n.args) > 1:
        return "fail", n.args[1]
    if name in HELPERS and isinstance(f, ast.Name) and len(n.args) >= 4:
        return "call", n.args[3]
    return None


def sites(root=SKILL):
    """点検の場所: [{file, func, line, end, kind, text}]"""
    out = []
    for name in FILES:
        tree = ast.parse((root / "scripts" / name).read_text(encoding="utf-8"))
        funcs = [(n.lineno, n.end_lineno, n.name) for n in ast.walk(tree) if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
        within = lambda n: min((f for f in funcs if f[0] <= n.lineno <= f[1]), key=lambda f: f[1] - f[0], default=(0, 0, "（関数の外）"))[2]
        for n in ast.walk(tree):
            hit = site_of(n)
            if hit is None:
                continue
            kind, arg = hit
            func = within(n)
            if kind == "fail" and func in HELPERS:                     # 部品の中の FAIL は、呼び元で数える
                continue
            text = re.sub(r"\s+", " ", text_of(arg))
            if kind == "call":
                text = f"shape（{text}）"
            out.append({"file": name, "func": func, "line": n.lineno, "end": n.end_lineno, "kind": kind, "text": text})
    return sorted(out, key=lambda s: (s["file"], s["line"]))


def listed(path=LIST):
    """試験しない点検の一覧: [{file, func, part, why}]（表の行）"""
    out = []
    for ln in path.read_text(encoding="utf-8").splitlines() if path.exists() else []:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) == 4 and cells[0] in FILES:
            out.append({"file": cells[0], "func": cells[1].strip("`"), "part": cells[2].strip("「」"), "why": cells[3]})
    return out


def matches(row, s):
    return row["file"] == s["file"] and row["func"] == s["func"] and row["part"] in s["text"]


def tally(hits, root=SKILL, path=LIST):
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
    ap.add_argument("-w", "--work", help="自己点検の組み立ての場所（selftest.py -w と同じ）")
    ap.add_argument("--log", help="FAIL を出した場所の控え（自己点検を回さない）")
    a = ap.parse_args()
    if a.log:
        log, broken = pathlib.Path(a.log), False
    else:
        if a.work:
            pathlib.Path(a.work).mkdir(parents=True, exist_ok=True)
        log = pathlib.Path(tempfile.mkstemp(prefix="sc_fail_", suffix=".log", dir=a.work or None)[1])
        cmd = [sys.executable, str(SKILL / "evals" / "selftest.py")] + (["-w", a.work] if a.work else [])
        r = subprocess.run(cmd, env=dict(os.environ, **{FAIL_LOG: str(log)}), capture_output=True, text=True)
        broken = r.returncode != 0
        said = log.with_suffix(".selftest.txt")
        said.write_text(r.stdout + r.stderr, encoding="utf-8")
        print(f"自己点検: 終了コード {r.returncode}（{r.stdout.strip().splitlines()[-1] if r.stdout.strip() else '出力なし'}）。控え: {log}・自己点検の出力: {said}")
        for ln in [x for x in r.stdout.splitlines() if x.startswith("NG ")][:10]:
            print(f"  {ln[:150]}")
    ss, fired, rows, loose, noisy, stale = tally(read_log(log))
    kinds = collections.Counter(s["kind"] for s in ss)
    print(f"点検の場所 {len(ss)} か所（FAIL {kinds['fail']}・形の点検の呼び元 {kinds['call']}）: "
          f"自己点検で出た {len(fired)}・試験しない（{LIST.name}）{len(ss) - len(fired) - len(loose)}・出ないのに一覧に無い {len(loose)}")
    for s in loose:
        print(f"  出ない  {s['file']}:{s['line']} [{s['func']}] {s['text'][:110]}")
    for s in noisy:
        print(f"  一覧にあるのに出た  {s['file']}:{s['line']} [{s['func']}] {s['text'][:80]}（一覧から外す）")
    for r in stale:
        print(f"  一覧の行が、どの点検にも当たらない  {r['file']} {r['func']}「{r['part']}」（点検の文・関数の名前が変わった。一覧を直す）")
    if broken:
        print("自己点検が NG（上の NG の行）")
    sys.exit(1 if loose or stale or noisy or broken else 0)


if __name__ == "__main__":
    main()
