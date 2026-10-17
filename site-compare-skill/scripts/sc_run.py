#!/usr/bin/env python3
"""sc_run.py — 拠点間比較の、機械だけで進められる工程を順に回す。LLM・人の出番になったら止まり、次にやることを出す。

  python3 scripts/sc_run.py -w <比較の作業フォルダ>              # 進められる所まで進める
  python3 scripts/sc_run.py -w <比較の作業フォルダ> --dry-run    # 何を実行するかだけ見る

次の工程は sc_status.py と同じ判定（作業フォルダの中身だけを見る）。機械のコマンド（このスキルの scripts/ と、asis-flow-skill の scripts/ の棚卸し・
ページ化など）を回し、次のときに止まる: 次の仕事が LLM・人のもの（読取・書き分け・用語の対応・比べる行・タスクの対応・照合・目視）／
コマンドが FAIL を出した／同じコマンドをもう一度勧められた（結果を見て、LLM が書く番）／納品の門まで来た（出力先は利用者が決める）。
"""
import argparse
import pathlib
import shlex
import subprocess
import sys

from sc_common import ASIS_SCRIPTS
from sc_status import W, Status

S = pathlib.Path(__file__).resolve().parent
MAX_ROUNDS = 12


def commands(todo):
    """次にやることの 1 行から、実行するコマンドの並び（コメントを外し、&& で分ける）。コマンドでなければ空"""
    line = todo.split("    #", 1)[0].strip()
    if not (line.startswith("python3 scripts/") or line.startswith(f"python3 {ASIS_SCRIPTS}/")):
        return []
    return [shlex.split(part) for part in line.split(" && ")]


def run_one(argv, cmp):
    script = pathlib.Path(argv[1])
    path = S / script.name if argv[1].startswith("scripts/") else script
    args = [sys.executable, str(path)] + [str(cmp) if x == W else x for x in argv[2:]]
    r = subprocess.run(args, capture_output=True, text=True)
    tail = [ln for ln in (r.stdout + r.stderr).splitlines() if ln.strip()][-12:]
    return r.returncode, tail


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    cmp = pathlib.Path(a.work).expanduser().resolve()
    if not (cmp / "sites.json").exists():
        sys.exit(f"sites.json が無い（python3 scripts/sc_init.py <入力フォルダ> -w {cmp}）")
    done, why = [], ""
    for _ in range(MAX_ROUNDS):
        st = Status(cmp, fast=False)
        st.run()
        if not st.todo:
            why = "すべて済み。納品の門が写した出力先の一式を渡す"
            break
        todo = st.todo[0]
        if "sc_deliver.py" in todo:
            why = f"納品の門まで来た。出力先を決めて python3 scripts/sc_deliver.py -w {W} --to <出力先>"
            break
        cmds = commands(todo)
        if not cmds:
            why = "次は LLM・人の仕事"
            break
        if todo in done:
            why = "機械の仕事は済んだ。その結果を見て、LLM・人が書く番"
            break
        done.append(todo)
        print(f"▶ {todo.split('    #', 1)[0].strip().replace(W, shlex.quote(str(cmp)))}")
        if a.dry_run:
            why = "--dry-run なので、ここまで（実行していない）"
            break
        failed = False
        for argv in cmds:
            code, tail = run_one(argv, cmp)
            print("\n".join("    " + ln[:200] for ln in tail))
            if code != 0:
                failed, why = True, f"{pathlib.Path(argv[1]).name} が FAIL を出した（終了コード {code}）。上の行を直してから、もう一度"
                break
        if failed:
            break
    else:
        why = f"{MAX_ROUNDS} 工程を回した。もう一度実行すると続きから進む"
    st = Status(cmp, fast=False)
    st.run()
    here = shlex.quote(str(cmp))
    print(f"\n== 進み具合（{cmp.name}）==")
    print("\n".join(x.replace(W, here) for x in st.rows))
    if st.todo:
        print("\n次にやること:")
        for t in st.todo[:3]:
            print("  " + t.replace(W, here))
    print(f"\n止めた理由: {why.replace(W, here)}")


if __name__ == "__main__":
    main()
