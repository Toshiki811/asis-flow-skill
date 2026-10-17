#!/usr/bin/env python3
"""fixture_tools.py — 架空の 3 拠点の作業フォルダを、期待データから組み立てる（自己点検・大量の試験・手での確かめが使う。LLM は使わない）。

  python3 evals/fixture_tools.py build <比較の作業フォルダ> [--in <入力の置き場>] [--stage sites|compare]

  sites    拠点ごとの asis の作業フォルダを、期待データ（カード・階層・用語・分類・フロー）で作り、照合済みにする
           （照合済みは、試験のためだけの手順: 照合に出す版を控え、全部「支持」の報告を書き、review.py draft。実際は、作り方を知らない別のエージェントが照合する）
  compare  sites に続けて、比較の期待データ（書き分け・用語の対応・比べる行・タスク）を置く（既定）
"""
import argparse
import json
import pathlib
import shutil
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
SKILL = HERE.parent
FX = HERE / "fixture_expected"
sys.path.insert(0, str(SKILL / "scripts"))

from sc_common import ASIS_SCRIPTS  # noqa: E402


def run(script, *args, asis=False, check=True):
    path = (ASIS_SCRIPTS if asis else SKILL / "scripts") / script
    r = subprocess.run([sys.executable, str(path), *map(str, args)], capture_output=True, text=True)
    if check and r.returncode != 0:
        sys.exit(f"{script} {' '.join(map(str, args))} が失敗した:\n{r.stdout[-1500:]}\n{r.stderr[-1500:]}")
    return r.returncode, r.stdout + r.stderr


def sites_json(inp):
    return {"baseline": "TK", "shared": {"inputs": [str(inp / "共通")]},
            "sites": [{"id": "TK", "name": "東京支店", "aliases": ["東京"], "language": "ja", "inputs": [str(inp / "東京支店")]},
                      {"id": "OS", "name": "大阪支店", "aliases": ["大阪", "Osaka"], "language": "ja", "inputs": [str(inp / "大阪支店")]},
                      {"id": "SG", "name": "シンガポール拠点", "aliases": ["シンガポール", "Singapore"], "language": "en",
                       "inputs": [str(inp / "シンガポール拠点")], "uses_shared": False}]}


def mark_reviewed(w):
    """asis の作業フォルダの全フローを、照合済みにする（試験だけの手順）"""
    sys.path.insert(0, str(ASIS_SCRIPTS))
    import review as RV
    run("build_drawio.py", "-w", w, asis=True)
    figs = RV.figures(w)
    sid = RV.record(w, figs, "全体")
    (w / "reports").mkdir(exist_ok=True)
    for F, fl in figs.items():
        vs = [{"node": n["id"], "verdict": "支持", "evidence": f"「{n['label']}」とある"} for n in fl["nodes"] if n["type"] not in ("start", "end")]
        (w / "reports" / f"review_{F}.json").write_text(json.dumps({"sheet": sid[F], "verdicts": vs, "findings": []}, ensure_ascii=False), encoding="utf-8")
    run("review.py", "draft", "-w", w, asis=True)


def build_sites(cmp, inp):
    """入力を作り、共通と拠点ごとの asis の作業フォルダを、期待データで組み立てる"""
    subprocess.run([sys.executable, str(HERE / "make_fixture.py"), str(inp)], check=True, capture_output=True)
    cmp.mkdir(parents=True, exist_ok=True)
    (cmp / "sites.json").write_text(json.dumps(sites_json(inp), ensure_ascii=False, indent=1), encoding="utf-8")
    sh = cmp / "asis" / "shared"
    run("inventory.py", inp / "共通", "-w", sh, asis=True)
    run("render_pages.py", "-w", sh, asis=True)
    shutil.copytree(FX / "asis" / "shared" / "cards", sh / "cards", dirs_exist_ok=True)
    for sid, ins in (("TK", ["共通", "東京支店"]), ("OS", ["共通", "大阪支店"]), ("SG", ["シンガポール拠点"])):
        run("inventory.py", *[inp / x for x in ins], "-w", cmp / "asis" / sid, asis=True)
    run("sc_share.py", "-w", cmp, "--all")
    for sid in ("TK", "OS", "SG"):
        w = cmp / "asis" / sid
        run("render_pages.py", "-w", w, asis=True)
        shutil.copytree(FX / "asis" / sid, w, dirs_exist_ok=True)
        mark_reviewed(w)


def fake_verify(cmp, rows=None):
    """対応の照合を済ませる（試験だけの手順: シートを作り、全部「支持」の報告を書き、採否表を作る。実際は、作り方を知らない別のエージェントが照合する）"""
    run("sc_verify.py", "sheet", "-w", cmp, *(["--only", ",".join(rows)] if rows else []))
    vs = json.loads((cmp / "review" / "versions.json").read_text(encoding="utf-8"))
    last = {}
    for v in vs["versions"]:
        last[v["row"]] = v
    (cmp / "reports").mkdir(exist_ok=True)
    for rid, v in last.items():
        if rows and rid not in rows:
            continue
        rep = {"sheet": v["id"], "row": rid, "verdicts": [{"claim": c, "verdict": "支持", "evidence": "原典のとおり"} for c in v["claims"]], "findings": []}
        (cmp / "reports" / f"verify_{rid}.json").write_text(json.dumps(rep, ensure_ascii=False), encoding="utf-8")
    run("sc_verify.py", "draft", "-w", cmp)


def build_compare(cmp):
    for name in ("applies.json", "terms.json", "axis.json"):
        if (FX / name).exists():
            shutil.copy2(FX / name, cmp / name)
    if (FX / "rows").exists():
        shutil.copytree(FX / "rows", cmp / "rows", dirs_exist_ok=True)
    run("sc_collect.py", "-w", cmp)
    run("sc_compute.py", "-w", cmp)
    fake_verify(cmp)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["build"])
    ap.add_argument("work")
    ap.add_argument("--in", dest="inp")
    ap.add_argument("--stage", choices=["sites", "compare"], default="compare")
    a = ap.parse_args()
    cmp = pathlib.Path(a.work).expanduser().resolve()
    if cmp.exists():
        sys.exit(f"もうある: {cmp}（消してから作る）")
    inp = pathlib.Path(a.inp).expanduser().resolve() if a.inp else cmp.parent / (cmp.name + "_in")
    build_sites(cmp, inp)
    if a.stage == "compare":
        build_compare(cmp)
    print(f"架空の 3 拠点の作業フォルダ → {cmp}（入力 {inp}）")


if __name__ == "__main__":
    main()
