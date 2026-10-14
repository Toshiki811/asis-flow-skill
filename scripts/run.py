#!/usr/bin/env python3
"""run.py — 機械だけで進められる工程を、順に回す。LLM・人の出番になったら止まり、次にやることを出す。

  python3 scripts/run.py -w <作業フォルダ>              # 進められる所まで進める
  python3 scripts/run.py -w <作業フォルダ> --dry-run    # 何を実行するかだけ見る

どの工程を次にやるかは、status.py と同じ判定（作業フォルダの中身だけを見る）。親（進行役）がコマンドの順番を覚えておく必要は無い:
書き出し（番号の振り直し → 図 → 配置の点検と PNG → 業務マトリクス一覧 → 点検）や、報告書の作り直しを、抜かさずに回す。最後は納品の門（deliver.py）。
止まるのは次のとき。どれも、理由と、次にやることを出す。
  ・次の仕事が LLM・人のもの（階層の判定を書く・ページを読む・分類を決める・フローを書く・原典照合・利用者に見せる確認点）
  ・点検が FAIL を出した（直すのは、そのページ・フローの担当）
  ・同じコマンドをもう一度勧められた（機械の仕事は済んでいて、その結果を見て LLM が書く番）
"""
import argparse
import pathlib
import shlex
import subprocess
import sys

from common import load
from status import Status, W

S = pathlib.Path(__file__).resolve().parent
MAX_ROUNDS = 12                                                      # 念のための上限（1回の実行で回す工程の数）


def commands(todo):
    """次にやることの1行から、実行するコマンドの並び（コメントを外し、&& で分ける）。コマンドでなければ空"""
    line = todo.split("    #", 1)[0].strip()
    if not line.startswith("python3 scripts/"):
        return []
    return [shlex.split(part) for part in line.split(" && ")]


def run_one(argv, work):
    """python3 scripts/<名前> … を実行し、(終了コード, 出力の末尾) を返す"""
    args = [sys.executable, str(S / pathlib.Path(argv[1]).name)] + [str(work) if x == W else x for x in argv[2:]]
    r = subprocess.run(args, capture_output=True, text=True)
    tail = [ln for ln in (r.stdout + r.stderr).splitlines() if ln.strip()][-12:]
    return r.returncode, tail


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    work = pathlib.Path(a.work).expanduser().resolve()
    done, why = [], ""
    for _ in range(MAX_ROUNDS):
        m = load(work / "manifest.json") if work.is_dir() else None
        if m is None:
            why = f"棚卸しがまだ（入力フォルダを決めて、python3 scripts/inventory.py <入力フォルダ> -w {W} を実行する）"
            break
        st = Status(work, m, fast=False)
        st.run()
        if not st.todo:
            why = "すべて済み。納品の門が写した出力先の一式を渡す"
            break
        todo = st.todo[0]
        cmds = commands(todo)
        if "deliver.py" in todo:                                     # 出力先は、利用者が決める
            why = f"納品の門まで来た。出力先を決めて python3 scripts/deliver.py -w {W} --to <出力先>"
            break
        if not cmds:
            why = "次は LLM・人の仕事"
            break
        if todo in done:
            why = "機械の仕事は済んだ。その結果を見て、LLM・人が書く番"
            break
        done.append(todo)
        print(f"▶ {todo.split('    #', 1)[0].strip()}")
        if a.dry_run:
            why = "--dry-run なので、ここまで（実行していない）"
            break
        failed = False
        for argv in cmds:
            code, tail = run_one(argv, work)
            print("\n".join("    " + ln[:200] for ln in tail))
            if code != 0:
                failed, why = True, f"{pathlib.Path(argv[1]).name} が FAIL を出した（終了コード {code}）。上の行を、そのページ・フローの担当に直させてから、もう一度"
                break
        if failed:
            break
    else:
        why = f"{MAX_ROUNDS} 工程を回した。もう一度実行すると続きから進む"
    m = load(work / "manifest.json") if work.is_dir() else None
    if m is not None:
        st = Status(work, m, fast=False)                             # 点検も回して見る（FAIL が残っているのに「済み」と出さない）
        st.run()
        print(f"\n== 進み具合（{work.name}）==")
        print("\n".join(st.rows))
        if st.todo:
            print("\n次にやること:")
            for t in st.todo[:3]:
                print("  " + t)
    print(f"\n止めた理由: {why}")


if __name__ == "__main__":
    main()
