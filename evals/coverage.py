#!/usr/bin/env python3
"""coverage.py — 点検（check_flow・check_layout・check_extract）の FAIL のうち、自己点検で一度も出ないものを数える。

  python3 evals/coverage.py -w <自己点検の作業フォルダ>     # 自己点検を、FAIL を出した場所を控えながら回して、数える（自己点検と同じ時間）
  python3 evals/coverage.py --log <控え>                   # 控え（ファイル:行 の並び）から数えるだけ

点検の場所（構文から拾う）:
  FAIL       .fail( と .fail_if(（段によって WARN になる点検）の呼び出し、配置の点検の yield "fail", …
  門の WARN  .warn( のうち、納品の門が FAIL にする WARN（deliver.PROMOTE の語を含むもの）。ほかの WARN は数えない（門を止めない）
  呼び元     点検の共通の部品（common.FAIL_HELPERS: 形の点検・読めない・根拠のページ ほか）を呼ぶ所。ファイルの種類ごとに、部品の FAIL が出るかを見る
  中継の内側 点検が 1 行で中継する、ほかのモジュールの誤りの文: 直しの一覧（apply_edits.py の self.bad・self.err）・アクティビティ図の直し（activity.py の同じ形）・
             台帳の紐付け（systems.py の link_errors・excluded_errors・folded_errors）・LLM が書いたブロックの木（code_lang.py の tree_problems）・
             取り出した構造（code_struct.py の tree_errors と、木が読めない・形が崩れている）・用語の対応（terms.py の problems）・
             照合の報告と照合に出した版の記録（review.py）・業務マトリクス一覧と確認事項一覧の読み戻しと処理時間の基準の表（table_template.py の read_side・read_back・time_table）・
             人の値の控え（common.py の answers.json・times.json の形）。誤りの文ごとに 1 か所（common.site で包んだ所と、self.bad の呼び出し）。
             入れ子の関数の中の文は、その関数の名前で数える（ブロックの木の点検の多くは visit）
控えは、環境変数 ASIS_FAIL_LOG があるとき、Report.fail・fail_if（FAIL になったときだけ）・warn・common.site が書き足す（共通の部品の中なら、呼び元の行も）。
部品の中の FAIL も、呼び元も、それぞれ 1 か所に数える。数えないもの: 落ちないための守り（common.guarded の FAIL）。
試験しない点検は、evals/untested_checks.md の表に、理由つきで並べる（ファイル・関数・文の一部で指す。行の番号は書かない）。
終了コード 1: 一覧に無いのに一度も出ない点検がある（試験を足すか、試験しない理由を書いて一覧に足す）／一覧にあるのに出た点検・
どの点検にも当たらない一覧の行がある（一覧を直す）／門が FAIL にする語（deliver.PROMOTE）が、どの WARN の文にも無い（文を変えて語が外れた）／
-w で回した自己点検が NG。
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
sys.path.insert(0, str(SKILL / "scripts"))
from common import FAIL_HELPERS, FAIL_LOG  # noqa: E402
from deliver import PROMOTE  # noqa: E402

FILES = ("check_flow.py", "check_layout.py", "check_extract.py")
RELAYS = ("apply_edits.py", "activity.py", "systems.py", "code_lang.py", "code_struct.py", "terms.py",   # 中継の内側（誤りの文ごとに数える）
          "review.py", "table_template.py", "common.py")
BAD_HOLDERS = ("apply_edits.py", "activity.py")                      # 誤りを溜める部品 bad（self.bad(i, op, 文)）のあるファイル
LIST = SKILL / "evals" / "untested_checks.md"


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
    """構文の節が点検の場所なら (種類, 文の節)。種類は fail（FAIL）／warn（門の WARN の候補）／call（共通の部品の呼び元）"""
    if isinstance(n, ast.Call):
        f = n.func
        name = f.attr if isinstance(f, ast.Attribute) else f.id if isinstance(f, ast.Name) else None
        if name == "fail" and n.args and isinstance(f, ast.Attribute):
            if any(k.arg == "at" for k in n.keywords):               # yield で返した点検を出す所（場所は yield の行で数える）
                return None
            return "fail", n.args[0]
        if name == "fail_if" and len(n.args) > 1:
            return "fail", n.args[1]
        if name == "warn" and n.args and isinstance(f, ast.Attribute):
            return "warn", n.args[0]
        if name in FAIL_HELPERS:
            arg = next((a for a in n.args if isinstance(a, ast.JoinedStr) or (isinstance(a, ast.Constant) and isinstance(a.value, str))), None)
            return "call", arg
    if isinstance(n, ast.Yield) and isinstance(n.value, ast.Tuple) and len(n.value.elts) == 2 \
            and isinstance(n.value.elts[0], ast.Constant) and n.value.elts[0].value == "fail":
        return "fail", n.value.elts[1]
    return None


def relay_of(n, func, name=""):
    """中継の内側の誤りの文なら ("relay", 文の節): common.site で包んだ文（誤りを溜める部品 bad の中は除く）と、self.bad の呼び出し"""
    holder = name in BAD_HOLDERS
    if isinstance(n, ast.Call):
        f = n.func
        if isinstance(f, ast.Name) and f.id == "site" and n.args and not (holder and func == "bad"):
            return "relay", n.args[0]
        if holder and isinstance(f, ast.Attribute) and f.attr == "bad" and isinstance(f.value, ast.Name) and f.value.id == "self":
            return "relay", n.args[2] if len(n.args) > 2 else None
    return None


def sites(root=SKILL):
    """点検の場所: [{file, func, line, end, kind, text}]。call の文は「部品の名前（呼び元が渡した名前）」"""
    out = []
    for name in FILES + RELAYS:
        tree = ast.parse((root / "scripts" / name).read_text(encoding="utf-8"))
        funcs = [(n.lineno, n.end_lineno, n.name) for n in ast.walk(tree) if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
        within = lambda n: min((f for f in funcs if f[0] <= n.lineno <= f[1]), key=lambda f: f[1] - f[0], default=(0, 0, "（関数の外）"))[2]
        for n in ast.walk(tree):
            hit = (relay_of(n, within(n), name) if hasattr(n, "lineno") else None) if name in RELAYS else site_of(n)
            if hit is None:
                continue
            kind, arg = hit
            text = re.sub(r"\s+", " ", text_of(arg) if arg is not None else "…")
            if kind == "warn" and not any(k in text for k in PROMOTE):
                continue
            if kind == "call":
                text = f"{n.func.attr if isinstance(n.func, ast.Attribute) else n.func.id}（{text}）"
            func = within(n)
            out.append({"file": name, "func": func, "line": n.lineno, "end": n.end_lineno, "kind": kind, "text": text})
    return sorted(out, key=lambda s: (s["file"], s["line"]))


def listed(path=LIST):
    """試験しない点検の一覧: [{file, func, part, why}]（表の行）"""
    out = []
    for ln in path.read_text(encoding="utf-8").splitlines() if path.exists() else []:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cells) == 4 and cells[0].endswith(".py") and cells[0] in FILES + RELAYS:
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
        r = subprocess.run([sys.executable, str(SKILL / "evals" / "selftest.py"), "-w", a.work], env=dict(os.environ, **{FAIL_LOG: str(log)}),
                           capture_output=True, text=True)
        broken = r.returncode != 0
        said = log.with_suffix(".selftest.txt")
        said.write_text(r.stdout + r.stderr, encoding="utf-8")
        print(f"自己点検: 終了コード {r.returncode}（{r.stdout.strip().splitlines()[-1] if r.stdout.strip() else '出力なし'}）。控え: {log}・自己点検の出力: {said}")
        for ln in [x for x in r.stdout.splitlines() if x.startswith("NG")][:10]:
            print(f"  {ln[:150]}")
    else:
        log, broken = pathlib.Path(a.log), False
    ss, fired, rows, loose, noisy, stale = tally(read_log(log))
    dead = [k for k in PROMOTE if not any(s["kind"] == "warn" and k in s["text"] for s in ss)]
    kinds = collections.Counter(s["kind"] for s in ss)
    print(f"点検の場所 {len(ss)} か所（FAIL {kinds['fail']}・門の WARN {kinds['warn']}・共通の部品の呼び元 {kinds['call']}・中継の内側 {kinds['relay']}）: "
          f"自己点検で出た {len(fired)}・試験しない（{LIST.name}）{len(ss) - len(fired) - len(loose)}・出ないのに一覧に無い {len(loose)}")
    for s in loose:
        print(f"  出ない  {s['file']}:{s['line']} [{s['func']}] {s['text'][:110]}")
    for s in noisy:
        print(f"  一覧にあるのに出た  {s['file']}:{s['line']} [{s['func']}] {s['text'][:80]}（一覧から外す）")
    for r in stale:
        print(f"  一覧の行が、どの点検にも当たらない  {r['file']} {r['func']}「{r['part']}」（点検の文・関数の名前が変わった。一覧を直す）")
    for k in dead:
        print(f"  門が FAIL にする語「{k}」が、どの WARN の文にも無い（点検の文を変えた。deliver.PROMOTE か点検の文を直す）")
    if broken:
        print("自己点検が NG（上の NG の行）")
    sys.exit(1 if loose or stale or noisy or dead or broken else 0)


if __name__ == "__main__":
    main()
