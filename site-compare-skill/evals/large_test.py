#!/usr/bin/env python3
"""large_test.py — 拠点と行を増やした比較で、材料の上限・ページの分け方・点検の出力の畳み・途中からの再開・Excel の行数・時間を確かめる（LLM なし）。

  python3 evals/large_test.py                       # 一時フォルダで（終わったら消す）
  python3 evals/large_test.py -w <フォルダ> [--reps 10]

架空の 3 拠点（evals/fixture_tools.py）を組み立て、拠点の写しを 6 拠点に（東京支店・大阪支店・シンガポール拠点と、その写しの名古屋支店・
福岡支店・ベトナム拠点）、業務（経費申請・承認・支払処理の 3 行）を reps 回に複製し（業務分類も reps）、さらに 4 回分の 3 行を 1 行にまとめた
大きな行（タスク 28）を足す。記載なしの候補は機械の候補をそのまま写し、照合は試験だけの手順（全部「支持」）で済ませる。
拠点の as-is（asis の作業フォルダ）は複製しないので、点検は用語の対応から納品物まで（terms・axis・rows・compute・verify・outputs）を見る。
"""
import argparse
import copy
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import time

HERE = pathlib.Path(__file__).resolve().parent
SKILL = HERE.parent
S = SKILL / "scripts"
sys.path.insert(0, str(S))
sys.path.insert(0, str(HERE))

import fixture_tools as FT  # noqa: E402

CLONES = {"NA": ("TK", "名古屋支店", ["名古屋"], "ja"), "FK": ("OS", "福岡支店", ["福岡"], "ja"), "VN": ("SG", "ベトナム拠点", ["ベトナム", "Vietnam"], "vi")}
FLOW = re.compile(r"(?<![A-Za-z0-9])F(\d+)(?![0-9])")
BIG = 4                                                              # 大きな行にまとめる回数
NG = []


def say(ok, name, detail=""):
    if not ok:
        NG.append(name)
    print(f"{'OK' if ok else 'NG'}  {name}" + ("" if ok else f" — {detail}"))


def run(script, *args, env=None):
    t = time.time()
    r = subprocess.run([sys.executable, str(S / script), *map(str, args)], capture_output=True, text=True, env=env)
    return r.returncode, r.stdout + r.stderr, time.time() - t


def jload(p):
    return json.loads(pathlib.Path(p).read_text(encoding="utf-8"))


def jsave(p, d):
    pathlib.Path(p).write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def shift(obj, off):
    """フロー ID（F01・F03-1）と、それを含む箱の鍵を off だけずらす"""
    return json.loads(FLOW.sub(lambda m: f"F{int(m.group(1)) + off:02d}", json.dumps(obj, ensure_ascii=False)))


def clone_site(obj, src, dst):
    """拠点 ID の入った所（members・absent の鍵・内容の差の拠点と根拠）を、写しの拠点にも足す"""
    for t in obj.get("tasks") or []:
        if src in t["members"]:
            t["members"][dst] = list(t["members"][src])
        if src in (t.get("absent") or {}):
            t["absent"][dst] = copy.deepcopy(t["absent"][src])
        for c in t.get("content") or []:
            if src in c["sites"]:
                c["sites"].append(dst)
            c["evidence"] += [dict(e, site=dst) for e in c["evidence"] if e["site"] == src]


def grow(cmp, reps):
    """架空の 3 拠点の比較の作業フォルダを、6 拠点 × (3 × reps + 1) 行に育てる"""
    n = reps + BIG
    sites = jload(cmp / "sites.json")
    by = {x["id"]: x for x in sites["sites"]}
    for cid, (src, name, aliases, lang) in CLONES.items():
        sites["sites"].append(dict(copy.deepcopy(by[src]), id=cid, name=name, aliases=aliases, language=lang))
    jsave(cmp / "sites.json", sites)
    for sid in ("TK", "OS", "SG"):
        snap = jload(cmp / "snap" / f"{sid}.json")
        big = dict(snap, categories=[], flows={}, families={}, nodes={}, excluded=[], questions=snap["questions"] * n)
        for r in range(n):
            x = shift({k: snap[k] for k in ("categories", "flows", "families", "nodes", "excluded")}, 3 * r)
            big["categories"] += x["categories"]
            big["excluded"] += x["excluded"]
            for k in ("flows", "families", "nodes"):
                big[k].update(x[k])
        jsave(cmp / "snap" / f"{sid}.json", big)
        for cid, (src, *_rest) in CLONES.items():
            if src == sid:
                jsave(cmp / "snap" / f"{cid}.json", dict(big, site=cid))
    tm = jload(cmp / "terms.json")
    for g in tm["groups"]:
        g["members"] += [dict(m, site=cid) for cid, (src, *_r) in CLONES.items() for m in g["members"] if m["site"] == src]
    jsave(cmp / "terms.json", tm)
    ax = jload(cmp / "axis.json")
    base_rows, groups, rows = ax["rows"], [], []

    def with_clones(sites_):
        out = list(sites_)
        for cid, (src, *_r) in CLONES.items():
            out += [dict(copy.deepcopy(x), site=cid) for x in sites_ if x["site"] == src]
        return out
    for r in range(reps):
        groups.append({"id": f"G{r + 1:02d}", "name": f"経費精算（その {r + 1}）", "name_from": None})
        for k, br in enumerate(base_rows):
            x = shift(br, 3 * r)
            x.update(id=f"B{3 * r + k + 1:02d}", group=f"G{r + 1:02d}", name=f"{br['name']}（その {r + 1}）", sites=with_clones(x["sites"]))
            rows.append(x)
    bid = f"B{3 * reps + 1:02d}"
    groups.append({"id": f"G{reps + 1:02d}", "name": "経費精算のまとめ", "name_from": None})
    bsites = []
    for sid in ("TK", "OS", "SG"):
        parts = [p for r in range(reps, n) for br in base_rows for x in shift(br, 3 * r)["sites"] if x["site"] == sid for p in x["parts"]]
        bsites.append({"site": sid, "relation": "same" if sid != "SG" else "split", "parts": parts, "reason": "" if sid != "SG" else "申請と承認を 1 本のフローに書いている"})
    rows.append({"id": bid, "group": f"G{reps + 1:02d}", "parent": None, "name": "経費精算のまとめ", "name_from": shift(base_rows[0]["name_from"], 3 * reps),
                 "origin": "baseline", "sites": with_clones(bsites)})
    jsave(cmp / "axis.json", dict(ax, groups=groups, rows=rows))
    base_m = {br["id"]: jload(cmp / "rows" / f"{br['id']}.match.json") for br in base_rows}
    for p in list((cmp / "rows").glob("*.json")):
        p.unlink()
    for r in range(reps):
        for k, br in enumerate(base_rows):
            m = shift(base_m[br["id"]], 3 * r)
            rid = f"B{3 * r + k + 1:02d}"
            m["row"] = rid
            for q in m["questions"]:
                q["row"] = rid
            for cid, (src, *_r) in CLONES.items():
                clone_site(m, src, cid)
            jsave(cmp / "rows" / f"{rid}.match.json", m)
    tasks = []
    for r in range(reps, n):
        for br in base_rows:
            for t in shift(base_m[br["id"]], 3 * r)["tasks"]:
                tasks.append(dict(t, id=f"t{len(tasks) + 1}"))
    m = {"row": bid, "tasks": tasks, "questions": []}
    for cid, (src, *_r) in CLONES.items():
        clone_site(m, src, cid)
    jsave(cmp / "rows" / f"{bid}.match.json", m)
    import sc_match
    from sc_common import sites_of
    from sc_rows import Model
    mdl = Model(cmp, sites_of(cmp))
    for rid in [x["id"] for x in rows]:
        p = cmp / "rows" / f"{rid}.match.json"
        m = jload(p)
        for t in m["tasks"]:
            for s, ab in (t.get("absent") or {}).items():
                ab["checked"] = [c[0] for c in sc_match.candidates(mdl, rid, t, s)]
        jsave(p, m)
    for name in ("numbers.json",):
        (cmp / name).unlink(missing_ok=True)
    for d in ("review", "reports", "out"):
        shutil.rmtree(cmp / d, ignore_errors=True)
    return [x["id"] for x in rows], bid, len(tasks)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", help="組み立てる場所（省くと一時フォルダ。終わったら消す）")
    ap.add_argument("--reps", type=int, default=10)
    a = ap.parse_args()
    root = pathlib.Path(a.work).expanduser().resolve() if a.work else pathlib.Path(tempfile.mkdtemp(prefix="sc_large_"))
    root.mkdir(parents=True, exist_ok=True)
    cmp = root / "large"
    shutil.rmtree(cmp, ignore_errors=True)
    t0 = time.time()
    FT.build_sites(cmp, root / "large_in")
    FT.build_compare(cmp)
    rows, bid, big_tasks = grow(cmp, a.reps)
    nsites = 3 + len(CLONES)
    print(f"拠点 {nsites}・行 {len(rows)}（大きな行 {bid}: タスク {big_tasks}）→ {cmp}（組み立て {time.time() - t0:.0f} 秒）")
    times = {}
    rc, out, times["計算"] = run("sc_compute.py", "-w", cmp)
    say(rc == 0 and f"計算した行 {len(rows)}" in out, "計算: 全部の行", out[-300:])
    t = time.time()
    FT.fake_verify(cmp)
    times["照合（試験の手順）"] = time.time() - t
    rc, out, times["Excel"] = run("sc_book.py", "-w", cmp)
    say(rc == 0, "書き出し: Excel", out[-300:])
    rc, out, times["draw.io"] = run("sc_drawio.py", "-w", cmp)
    say(rc == 0, "書き出し: draw.io", out[-300:])
    rc, out, times["配置の点検"] = run("sc_layout.py", "-w", cmp, "--no-png")
    say(rc == 0, "配置の点検: FAIL 0（全ページ）", out[-400:])
    total = 0
    for st in ("terms", "axis", "rows", "compute", "verify", "outputs"):
        rc, out, sec = run("sc_check.py", "-w", cmp, "--stage", st, "--fast")
        total += sec
        say(rc == 0, f"点検 {st}: FAIL 0", "\n".join(x for x in out.splitlines() if x.startswith("FAIL"))[:600])
    times["点検（6 段）"] = total

    # 材料の上限
    rc, out, _ = run("sc_axis.py", "plan", "-w", cmp)
    sizes = {m.group(1): int(m.group(2).replace(",", "")) for m in re.finditer(r"^(B\d+) .*?材料 ([\d,]+) 字", out, re.M)}
    say(len(sizes) == len(rows) and max(sizes.values()) <= 40000 and "★" not in out, "材料: どの行も上限（4 万字）まで", f"最大 {max(sizes.values()) if sizes else '?'}")
    rc, out, _ = run("sc_axis.py", "plan", "-w", cmp, "--cap", str(sizes.get(bid, 0) - 1))
    say("★" in out.split(bid)[-1].splitlines()[0] if bid in out else False, "材料: 上限を超える行に ★（子の行に分ける）", out[-300:])
    rc, out, _ = run("sc_axis.py", "material", "-w", cmp)
    say(rc == 0 and "材料" in out.splitlines()[-1], "材料: 業務種別の対応の材料の大きさ", out.splitlines()[-1] if out else "")

    # ページの分け方
    lay = jload(cmp / "out" / "拠点比較フロー.layout.json")
    names = [p["name"] for p in lay["pages"]]
    want = 1 + 3 * a.reps * 2 + 2 * 2
    say(len(names) == want, f"ページ: 1 ページ目＋行ごとに（6 拠点は基準と 3 拠点ずつで 2 ページ・タスク 25 超えはタスクの境で）= {want}", f"{len(names)}")
    say(any("（1/2）" in x and "［2/2］" in x for x in names) and any(x.startswith(bid) and "（2/2）［1/2］" in x for x in names),
        "ページ: 拠点の分け方と、タスクの境の分け方の名前", f"{[x for x in names if x.startswith(bid)]}")
    per = {}
    for p in lay["pages"]:
        if p["id"].startswith("row-"):
            per[p["id"]] = sum(1 for k in p["nodes"] if k.startswith("TASK."))
    say(per and max(per.values()) <= 25, "ページ: 1 ページのタスクは 25 まで", f"{max(per.values()) if per else '?'}")
    first = next(p for p in lay["pages"] if p["id"].startswith(f"row-{bid}"))
    say(all(k.split(".")[0] in ("TK", "OS", "SG", "NA", "TASK") for k in first["nodes"] if "." in k), "ページ: 基準（東京支店）が、どのページにも左に出る",
        f"{sorted({k.split('.')[0] for k in first['nodes']})}")

    # Excel の行数
    import openpyxl
    wb = openpyxl.load_workbook(cmp / "out" / "拠点比較.xlsx")
    ntask = sum(len(jload(cmp / "rows" / f"{r}.result.json")["tasks"]) for r in rows)
    filled = lambda ws: sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if any(v not in (None, "") for v in row))
    got = {"拠点タスク一覧": filled(wb["拠点タスク一覧"]), "拠点別の明細": filled(wb["拠点別の明細"]), "複雑性ヒートマップ": filled(wb["複雑性ヒートマップ"])}
    say(got == {"拠点タスク一覧": ntask, "拠点別の明細": ntask * nsites, "複雑性ヒートマップ": len(rows)},
        f"Excel: 行数（タスク {ntask}・タスク × 拠点 {ntask * nsites}・行 {len(rows)}）", f"{got}")
    head = [c.value for c in wb["拠点タスク一覧"][1]]
    say(head[5:5 + nsites] == ["【基準】東京支店", "大阪支店", "シンガポール拠点", "名古屋支店", "福岡支店", "ベトナム拠点"], "Excel: 拠点の列（6 拠点。基準が左）", f"{head}")

    # 点検の出力の畳み
    d = root / "fold"
    shutil.rmtree(d, ignore_errors=True)
    shutil.copytree(cmp, d)
    for p in (d / "rows").glob("*.match.json"):
        m = jload(p)
        for t in m["tasks"]:
            t["absent"] = {}
        jsave(p, m)
    rc, out, _ = run("sc_check.py", "-w", d, "--stage", "rows", "--fast", "--report", d / "out" / "fold.txt")
    shown = sum(1 for x in out.splitlines() if x.startswith("FAIL") and "記載なしの根拠（absent）が無い" in x)
    full = sum(1 for x in (d / "out" / "fold.txt").read_text(encoding="utf-8").splitlines() if "記載なしの根拠（absent）が無い" in x)
    say(rc == 1 and shown == 10 and full > 10 and "同じ種類があと" in out, f"畳み: 画面は同じ種類を 10 件まで（全 {full} 件は --report のファイルに）", f"画面 {shown}・ファイル {full}")

    # 途中からの再開（タスクの対応の済んだ行だけを計算し、戻した行は照合し直さなくてよい）
    for p in (d / "rows").glob("*.match.json"):
        shutil.copy2(cmp / "rows" / p.name, p)
    gone = rows[1::3]
    for r in gone:
        (d / "rows" / f"{r}.match.json").rename(d / f"{r}.keep")
        (d / "rows" / f"{r}.result.json").unlink()
    rc, out, _ = run("sc_compute.py", "-w", d)
    rc2, out2, _ = run("sc_check.py", "-w", d, "--stage", "rows", "--fast", "--all-lines")
    miss = [r for r in gone if f"rows/{r}.match.json が無い" in out2]
    say(f"計算した行 {len(rows) - len(gone)}" in out and miss == gone, f"再開: タスクの対応の無い行（{len(gone)}）を点検が挙げ、計算は済んだ行だけ",
        out[-200:] + " / " + "・".join(miss[:5]))
    for r in gone:
        (d / f"{r}.keep").rename(d / "rows" / f"{r}.match.json")
    rc, out, _ = run("sc_compute.py", "-w", d)
    rc2, out2, _ = run("sc_check.py", "-w", d, "--stage", "verify", "--fast")
    say(f"計算した行 {len(rows)}" in out and rc2 == 0, "再開: 戻した行も計算し、中身の変わらない行は照合し直さなくてよい", out2[-300:])
    shutil.rmtree(d)

    # 揺れ（文字のハッシュの種を変えても、計算・Excel・draw.io が同じ。set の順に頼ると、ここで割れる）
    outs = set()
    for seed in ("1", "4", "11", "12345"):
        env = dict(os.environ, PYTHONHASHSEED=seed)
        for script in ("sc_compute.py", "sc_book.py", "sc_drawio.py"):
            run(script, "-w", cmp, env=env)
        outs.add(tuple([(cmp / "out" / "拠点比較フロー.drawio").read_text(encoding="utf-8"), (cmp / "out" / "拠点比較.mark").read_text(encoding="utf-8")]
                       + [p.read_text(encoding="utf-8") for p in sorted((cmp / "rows").glob("*.result.json"))]))
    say(len(outs) == 1, "揺れ: ハッシュの種を変えて 4 回書き出しても、計算・Excel・draw.io が同じ", f"{len(outs)} 通り")

    print("\n時間（秒）: " + "・".join(f"{k} {v:.1f}" for k, v in times.items()))
    print(f"\nNG {len(NG)}" + (f"（{'・'.join(NG[:8])}）" if NG else ""))
    if not a.work:
        shutil.rmtree(root, ignore_errors=True)
    sys.exit(1 if NG else 0)


if __name__ == "__main__":
    main()
