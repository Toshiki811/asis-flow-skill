#!/usr/bin/env python3
"""sc_deliver.py — 拠点間比較の納品の門。書き出しと点検を決まった順に回し直し、全部通ったら出力先へ写す（工程 9）。

  python3 scripts/sc_deliver.py -w <比較の作業フォルダ> --to <出力先>
  python3 scripts/sc_deliver.py -w <比較の作業フォルダ> --check          # 門の点検だけ（写さない）

回すもの（この順）: 計算（sc_compute）→ Excel（sc_book）→ draw.io（sc_drawio）→ 配置の点検と PNG（sc_layout）→ 機械チェック全段（sc_check。拠点の門も回す）。
門を通る条件: 機械チェックが全段 FAIL 0（拠点ごとの asis のフローと原典照合・対応の照合の支持率 95% 以上を含む）・配置の点検 FAIL 0・
全ページの PNG がそろっていて、図より新しい（draw.io デスクトップが要る。無い環境では門を通らない。下見は sc_layout.py --preview）。
写すもの: <出力先>/納品物/（拠点比較.xlsx・拠点比較フロー.drawio・png/）と <出力先>/作業の記録/（チェック結果・対応の照合シート・照合の採否表・比較の数）。
出力先の 納品物/ に、この門が写したもの以外のファイルがあれば、上書きせずに止まる。
"""
import argparse
import datetime
import json
import pathlib
import shutil
import subprocess
import sys

from sc_common import BOOK, FIG, load, pct, result_path, save, sites_of, try_load, work_dir

S = pathlib.Path(__file__).resolve().parent
ROUNDS = 2


def run(script, cmp, *args):
    r = subprocess.run([sys.executable, str(S / script), "-w", str(cmp), *args], capture_output=True, text=True)
    return r.returncode, r.stdout + r.stderr


def pngs(cmp):
    """いまの図の全ページの PNG: [(道, ページの名前)]"""
    lay = load(cmp / "out" / FIG.replace(".drawio", ".layout.json"), None) or {"pages": []}
    stem = pathlib.Path(FIG).stem
    return [(cmp / "out" / "png" / f"{stem}_p{pg['index']:02d}_{pg['id']}.png", pg["name"]) for pg in lay["pages"]]


def numbers(cmp):
    """比較の数（比較の数.md・.json）"""
    from sc_rows import Model
    import sc_compute
    import sc_verify
    mdl = Model(cmp, sites_of(cmp))
    cfg = sc_compute.config()
    res = [r for r in (try_load(result_path(cmp, x["id"]))[0] for x in mdl.rows()) if isinstance(r, dict)]
    tasks = [x for r in res for x in r["tasks"]]
    cls = {k: sum(1 for x in tasks if x["class"] == k) for k in ("共通", "差異あり", "拠点固有")}
    tot = ok = 0
    shown = {v["id"]: v for v in sc_verify.versions(cmp)["versions"]}
    for r in res:
        rv = try_load(cmp / "review" / f"{r['row']}.review.json")[0] or {}
        got = {v["claim"]: v for v in rv.get("verdicts") or []}
        for c, (mark, _b) in sc_verify.claims(mdl, r["row"]).items():
            tot += 1
            v = got.get(c)
            ok += bool(v and v["verdict"] == "支持" and shown.get(v["sheet"], {}).get("claims", {}).get(c) == mark)
    per = {}
    for sid in mdl.order:
        per[sid] = {"name": mdl.names[sid], "tasks": sum(1 for x in tasks if x["marks"].get(sid) in ("○", "△")),
                    "absent": sum(1 for x in tasks if x["marks"].get(sid) == "—"), "diff": sum(1 for x in tasks if x["marks"].get(sid) == "△"),
                    "out": sum(1 for x in tasks if x["marks"].get(sid) == "対象外"),
                    "coarse": sum(1 for x in tasks if sid in x["grain"]["coarse"]),
                    "levels": {lv["name"]: sum(1 for r in res if (r["sites"].get(sid) or {}).get("complexity") and
                                               r["sites"][sid]["complexity"][r["basis"]]["level"] == lv["name"]) for lv in cfg["levels"]},
                    "requests": sum(1 for r in res for q in r["requests"] if q["site"] == sid)}
    return {"date": datetime.date.today().isoformat(), "sites": per, "baseline": mdl.names.get(mdl.baseline), "rows": len(res), "tasks": len(tasks),
            "class": cls, "support": {"ok": ok, "total": tot}, "basis": cfg["basis"], "weights": cfg["weights"], "levels": cfg["levels"]}


def numbers_md(n):
    lines = [f"# 比較の数（{n['date']}）", "", f"- 基準の拠点: {n['baseline'] or 'なし'}", f"- 比べる行（業務種別）: {n['rows']}",
             f"- タスク: {n['tasks']}（共通 {n['class']['共通']}・差異あり {n['class']['差異あり']}・拠点固有 {n['class']['拠点固有']}）",
             f"- 対応の照合の支持率: {n['support']['ok']}/{n['support']['total']}" + (f"（{pct(n['support']['ok'] / n['support']['total'])}）" if n['support']['total'] else ""),
             f"- 複雑性: {'そろえた粒度' if n['basis'] == 'aligned' else '原典の粒度'}で、分岐 × {n['weights']['branch']} ＋ 受け渡し × {n['weights']['handoff']}。境 "
             + "・".join(f"{lv['name']} {lv['max'] if lv['max'] is not None else '超'}" for lv in n["levels"]), "",
             "| 拠点 | 記載のあるタスク | 差のあるタスク（△） | 記載なし（—） | 対象外 | 粗く書いたタスク | 複雑性（行の数） | 追加の資料の依頼 |",
             "| --- | --- | --- | --- | --- | --- | --- | --- |"]
    for s in n["sites"].values():
        lines.append(f"| {s['name']} | {s['tasks']} | {s['diff']} | {s['absent']} | {s.get('out', 0)} | {s['coarse']} | "
                     + "・".join(f"{k} {v}" for k, v in s["levels"].items())
                     + f" | {s['requests']} |")
    return "\n".join(lines) + "\n"


def look(cmp):
    """門の点検: [(何, 直し方)]"""
    bad = []
    code, out = run("sc_check.py", cmp, "--report", str(cmp / "out" / "チェック_比較.txt"))
    if code:
        first = [x.strip()[len("FAIL"):].strip()[:110] for x in out.splitlines() if x.startswith("FAIL")][:3]
        bad.append(("機械チェック（sc_check）に FAIL がある" + "".join(f"\n      ・{x}" for x in first), "out/チェック_比較.txt の FAIL を直す"))
    code, out = run("sc_layout.py", cmp, "--no-png", "--report", str(cmp / "out" / "チェック_配置.txt"))
    if code:
        bad.append(("配置の点検（sc_layout）に FAIL がある", "out/チェック_配置.txt。行を子の行に分ける・タスクを見直す"))
    fig = cmp / "out" / FIG
    ps = pngs(cmp)
    lack = [name for p, name in ps if not p.exists() or (fig.exists() and p.stat().st_mtime < fig.stat().st_mtime)]
    if lack:
        bad.append((f"PNG が無い・図より古いページ {len(lack)}（{'・'.join(lack[:3])}）", "draw.io デスクトップを入れて python3 scripts/sc_layout.py -w <比較の作業フォルダ>。PNG を全ページ目で見る"))
    return bad


def chain(cmp):
    for script, args in (("sc_compute.py", []), ("sc_book.py", []), ("sc_drawio.py", []), ("sc_layout.py", [])):
        code, out = run(script, cmp, *args)
        if code and script != "sc_layout.py":
            sys.exit(f"{script} が止まった:\n{out[-1500:]}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--to")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    if not a.check and not a.to:
        ap.error("--to <出力先> か --check")
    bad = []
    for _ in range(ROUNDS):
        chain(cmp)
        bad = look(cmp)
        if not bad:
            break
    if bad:
        print("納品の門を通らない:")
        for what, how in bad:
            print(f"  ✗ {what} — {how}")
        sys.exit(1)
    n = numbers(cmp)
    save(cmp / "out" / "比較の数.json", n)
    (cmp / "out" / "比較の数.md").write_text(numbers_md(n), encoding="utf-8")
    if a.check:
        print("納品の門の点検は通った（--check なので、写していない）")
        return
    dest = pathlib.Path(a.to).expanduser().resolve()
    goods, log = dest / "納品物", dest / "作業の記録"
    mine = {BOOK, FIG, "png"}
    if goods.exists() and any(p.name not in mine for p in goods.iterdir()):
        sys.exit(f"{goods} に、この門が写したもの以外のファイルがある（{', '.join(p.name for p in goods.iterdir() if p.name not in mine)}）。別の出力先にするか、片付けてから")
    goods.mkdir(parents=True, exist_ok=True)
    log.mkdir(parents=True, exist_ok=True)
    for name in (BOOK, FIG):
        shutil.copy2(cmp / "out" / name, goods / name)
    if (goods / "png").exists():
        shutil.rmtree(goods / "png")
    (goods / "png").mkdir()
    for p, _name in pngs(cmp):
        shutil.copy2(p, goods / "png" / p.name)
    import sc_verify
    from sc_rows import Model
    (log / "対応の照合_採否表.md").write_text(sc_verify.review_md(cmp, Model(cmp, sites_of(cmp))), encoding="utf-8")
    for p in sorted((cmp / "out").glob("対応の照合シート_*.md")) + [cmp / "out" / x for x in ("チェック_比較.txt", "チェック_配置.txt", "比較の数.md")]:
        if p.exists():
            shutil.copy2(p, log / p.name)
    print(f"納品物 → {goods}\n作業の記録 → {log}\n")
    print(numbers_md(n))


if __name__ == "__main__":
    main()
