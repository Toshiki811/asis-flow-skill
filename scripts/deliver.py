#!/usr/bin/env python3
"""deliver.py — 納品の門。書き出しと点検を決まった順に回し、全部通ったら、納品物と作業の記録を出力先に写す。

  python3 scripts/deliver.py -w <作業フォルダ> --to <出力先>     # 門を通して、出力先の 納品物/・作業の記録/ に写す
  python3 scripts/deliver.py -w <作業フォルダ> --check            # 門を見るだけ（書き出し直さない・写さない。納品の数と現状の処理時間の写しは out/ に書く）

門の中身（全部通ったら写す）:
  1 書き出しのあとの点検が、全段で FAIL 0（check_flow の全段・check_layout・check_extract）。
    門では、次の WARN も FAIL にする: 文章の中で資料を ID で呼んでいる／業務マトリクス一覧に計算の結果が保存されていない／
    紐付け・関係・「続く」の根拠の文が出典のページに無い／読めない・曖昧な所の確認事項が無い／引用が原典と完全には一致しない／同じ手順を 2 つのフローに描いている
  2 図・業務マトリクス一覧・報告書が、もとのデータ（人が入れた値・採否表を含む）より新しい
  3 PNG が、いまの図より新しく、全ページある（draw.io が無ければ作れないので、通らない）
  4 照合シートが全フロー分あり、フロー・PNG より新しい
  5 原典照合（工程 8）が、全部の図で済んでいる（採否表・採否の記入・支持率 95% 以上は check_flow --stage review）
やり直し: 2〜4 と「書き出し直す」と言う FAIL は、門が書き出しをやり直して（連鎖の順: build_drawio → build_activity → check_layout（PNG）→
build_matrix → build_reports → make_compare）、もう一度見る（2 回まで）。データを直す必要があるもの（ほかの FAIL・照合がまだ）は止めて、
直す工程とコマンド（status.py と同じ）を出す。人が入れた値を取り込めないブック（開いたまま・列を動かした）も、書き出しが止まり、門も止まる。
通ったら: out/納品の数.md・.json を書き、出力先の 納品物/ に図・PNG・業務マトリクス一覧・確認事項一覧・出典対応表・階層と業務分類を、
作業の記録/ にチェック結果・照合シート・採否表・納品の数・現状の処理時間の写し（BPR の比較の基準。基準の表の版つき）を写す。出力先に、いまの一式に無いファイル（前の納品物の残り）があれば、写さずに止まる
（消さない。別の出力先を指定するか、人が片付けてから）。
納品の数: 資料・ページ・フロー・ノードの数、網羅率、引用照合率（完全一致の率も）、推定の数、確認事項の重要度別の件数（回答済みで残っているもの）、
支持率、人が入れた処理時間の割合・処理時間の合計・基準の表の版、残った WARN。
終了コード: 門を通れば 0、通らなければ 1。
"""
import argparse
import datetime
import pathlib
import shutil
import subprocess
import sys
import tempfile

from common import (ACT_FIG, BASELINE_FILE, DELIVERABLES, FLOW_FIG, MATRIX_BOOK, NUMBERS_FILE, RECORDS, REPORTS, REVIEW_MIN, flow_key, load, load_flows, save, split_figs,
                    work_dir)
from status import Status, newest, source_files
from validate import validate

S = pathlib.Path(__file__).resolve().parent
ROUNDS = 2                                                           # 書き出しをやり直す回数の上限
PROMOTE = {"資料を ID で呼んでいる": "data", "計算の結果が保存されていない式のセル": "export",   # 門では FAIL にする WARN（直し方の種類）
           "が、出典のページに見当たらない": "data",                      # 紐付け・関係・「続く」の根拠の文が原典に無い（機械が写すので、言い換え・作った文）
           "を出典にした確認事項がどのフローにも無い": "data",              # 読めない・曖昧な所を、確認事項にしていない
           "の文字と完全には一致しない": "data",                          # 手で写した引用（行の番号から機械が入れていない）
           "つのフローに描かれている": "data",                            # 同じ手順を、上位と下位に重ねて描いている疑い
           "まだ取り込んでいない人の値": "export"}                        # 人がブックに入れた値（書き出し直すと取り込む。現状の処理時間の写しの前に）
CHECKS = (("check_flow.py", "チェック_フロー.txt", ["--stage", "all"]), ("check_layout.py", "チェック_配置.txt", ["--no-png"]),
          ("check_extract.py", "チェック_読取.txt", []))
GOODS, LOG = "納品物", "作業の記録"


def run(script, work, *args):
    r = subprocess.run([sys.executable, str(S / script), "-w", str(work), *args], capture_output=True, text=True)
    return r.returncode, r.stdout + r.stderr


def activities(work):
    return sorted((work / "activities").glob("*.activity.json")) if (work / "activities").exists() else []


def pngs(work):
    """いまの図の全ページの PNG: [(PNG, 図のファイル)]（配置 *.layout.json のページの並びから）"""
    out = []
    for lp in sorted((work / "out").glob("*.layout.json")):
        lay = load(lp, {}) or {}
        stem = pathlib.Path(str(lay.get("file", ""))).stem
        out += [(work / "out" / "png" / f"{stem}_p{pg['index']:02d}_{pg['id']}.png", lp.with_name(str(lay.get("file", ""))))
                for pg in lay.get("pages") or [] if isinstance(pg, dict) and "index" in pg and "id" in pg]
    return out


def stale(work):
    """書き出し直せば直る所: 無い・もとのデータより古い出力"""
    out, said = work / "out", []
    flows, acts = load_flows(work), activities(work)
    src = newest(source_files(work, acts))
    old = lambda p, since: not p.exists() or p.stat().st_mtime < since
    drawio = sorted(out.glob(f"{FLOW_FIG}*.drawio"))
    if flows and not drawio:
        said.append(f"{FLOW_FIG}.drawio が無い")
    for p in drawio + ([out / ACT_FIG] if acts else []):
        if old(p, src):
            said.append(f"{p.name} が無い・もとのデータより古い")
    if flows or acts:
        book = out / MATRIX_BOOK
        if old(book, max(src, newest([work / "times.json", work / "answers.json"]))):
            said.append(f"{book.name} が無い・もとのデータ（人が入れた値を含む）より古い")
    since = max(src, newest(sorted((work / "review").glob("*.review.json")) + [work / "answers.json"]))
    lost = [name for name in REPORTS if old(out / name, since)]
    if lost:
        said.append(f"報告書（{'・'.join(lost)}）が無い・もとのデータ・採否表・回答より古い")
    pages = pngs(work)
    bad = [p.name for p, d in pages if old(p, d.stat().st_mtime if d.exists() else 0)]
    if bad:
        said.append(f"PNG が無い・図より古い {len(bad)} 枚（{'・'.join(bad[:4])}{' …' if len(bad) > 4 else ''}）")
    for fl in flows:
        F = flow_key(fl)
        sheets = list(out.glob(f"照合シート_{F}.pdf")) + list(out.glob(f"照合シート_{F}_*.pdf"))
        made = min((p.stat().st_mtime for p in sheets), default=0)
        if not sheets or made < newest([work / "flows" / f"{F}.flow.json"] + list((out / "png").glob(f"*_flow-{F}.png"))):
            said.append(f"照合シート_{F} が無い・フロー・図より古い")
    return said


def chain(work):
    """書き出しを、決まった順に回す。人が入れた値を守るために書き出しが止まったら、その理由を返す（門も止まる）"""
    flows, acts = load_flows(work), activities(work)
    steps = ([("build_drawio.py", ["--split"] if split_figs(work / "out") else [])] if flows else []) + ([("build_activity.py", [])] if acts else []) \
        + [("check_layout.py", []), ("build_matrix.py", []), ("build_reports.py", [])] + ([("make_compare.py", [])] if flows else [])
    for script, args in steps:
        code, text = run(script, work, *args)
        print(f"    ▶ {script}" + ("" if code == 0 else f"（終了コード {code}）"))
        if code != 0 and script.startswith("build_"):                 # 点検（配置）・照合シートの不足は、門が見て知らせる
            return f"{script} が止まった: " + " ／ ".join(ln for ln in text.strip().splitlines()[-3:])
    return None


def look(work, numbers_path):
    """門を見る: ([(種類, 文)], 数, 残った WARN)。種類は export（書き出し直せば直る）／data（データを直す）"""
    probs = [("export", x) for x in stale(work)]
    from check_layout import drawio_cli
    if drawio_cli() is None and any(not p.exists() for p, _d in pngs(work)):   # 書き出し直しても PNG は作れない
        probs.append(("data", "draw.io のコマンドが無いので、PNG を作れない（draw.io デスクトップを入れてから、もう一度）"))
    import review as RV
    figs = RV.figures(work)
    if figs and not any((work / "review").glob("*.review.json")):    # 採否表が 1 つも無い（check_flow は、照合の段を見ない。照合シートの記録だけでは、照合は始まっていない）
        probs.append(("data", f"原典照合（工程 8）がまだ: 図 {len(figs)} 枚に採否表が無い"))
    warns = []
    has_cards = (work / "cards").exists() and any((work / "cards").glob("*/*.json"))
    for script, report, args in CHECKS:
        if script == "check_extract.py" and not has_cards:
            continue
        path = work / "out" / report
        path.unlink(missing_ok=True)                                 # 前の回の記録を、いまの合否と取り違えない
        more = ["--numbers", str(numbers_path)] if script == "check_flow.py" else []
        _code, text = run(script, work, *args, "--report", str(path), *more)
        lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
        if not any(" FAIL / " in ln for ln in lines):                 # 点検が途中で止まった: 合格と取り違えない
            probs.append(("data", f"{script} が動かない（{(text.strip().splitlines() or ['出力なし'])[-1][:100]}）"))
            continue
        for ln in lines:
            if ln.startswith("FAIL  "):
                probs.append(("export" if "書き出し直" in ln else "data", f"{report}: {ln[6:]}"))
            elif ln.startswith("WARN  "):
                hit = next((k for k in PROMOTE if k in ln), None)
                if hit:
                    probs.append((PROMOTE[hit], f"{report}: {ln[6:]}（門では FAIL）"))
                else:
                    warns.append(f"{report}: {ln[6:]}")
    nums = load(numbers_path, {}) or {}
    return probs, nums, warns


def numbers_text(work, nums, warns, dest):
    """納品の数（md の行）"""
    m = load(work / "manifest.json", {}) or {}
    docs = [d for d in m.get("docs") or [] if not d.get("duplicate_of") and d.get("kind") != "unsupported"]
    rate = lambda a, b: f"{a / b:.1%}" if b else "—"
    q, t, rv = nums.get("questions") or {}, nums.get("times") or {}, nums.get("review") or {}
    rows = [("資料", f"読む対象 {len(docs)}（うちプログラムのソース {sum(d.get('kind') == 'code' for d in docs)}）"),
            ("ページ", f"内容のページ {nums.get('pages_content', 0)}・フローに使った {nums.get('pages_used', 0)}（網羅率 {rate(nums.get('pages_used', 0), nums.get('pages_content', 0))}。"
                       "使っていないページの理由は、出典対応表）"),
            ("手順・分岐の項目", f"{nums.get('items_all', 0)}（フローに入れた {nums.get('items_in', 0)}・対象外（理由つき）{nums.get('items_out', 0)}・どちらでもない {nums.get('items_lost', 0)}）"),
            ("業務フロー", f"{nums.get('flows', 0)} 本・出典つきノード {nums.get('nodes', 0)}・推定 {nums.get('inferred', 0)} か所"),
            ("アクティビティ図", f"{nums.get('activities', 0)} 枚"),
            ("引用照合率", f"{nums.get('quotes_ok', 0)}/{nums.get('quotes', 0)}（{rate(nums.get('quotes_ok', 0), nums.get('quotes', 0))}）。"
                          f"うち原典の文字と完全に一致 {nums.get('quotes_exact', 0)}（{rate(nums.get('quotes_exact', 0), nums.get('quotes', 0))}）・画像由来（カードとの照合）{nums.get('quotes_image', 0)}"),
            ("確認事項", f"全 {q.get('all', 0)} 件（高 {q.get('high', 0)}・中 {q.get('mid', 0)}・低 {q.get('low', 0)}）。回答済みで、まだ図に反映していないもの {q.get('answered_open', 0)}")]
    for kind, name in (("flow", "業務フロー"), ("activity", "アクティビティ図")):
        s = rv.get(kind)
        if s:
            rows.append((f"原典照合（{name}）", f"支持率 {rate(s['support'], s['nodes'])}（支持 {s['support']} ÷ 開始・終了を除く全ノード {s['nodes']}。"
                                                f"判定 {s['judged']}・一部支持 {s['partial']}・不支持 {s['against']}。照合した図 {s['reviewed']}/{s['figures']}。下限 {REVIEW_MIN:.0%}）"))
    if t:
        rows.append(("処理時間", f"人が入れた値 {t.get('human', 0)} セル・仮の値 {t.get('soft', 0)} セル（人の値の割合 {rate(t.get('human', 0), t.get('human', 0) + t.get('soft', 0))}）・件数 {t.get('counts', 0)} 行"
                                 + (f"・合計 {t['total']} 分・基準の表の版 {t['table_version']}（見直しの前と後は、同じ版の表で比べる。写しは {BASELINE_FILE}）" if t.get("table_version") else "")))
    rows.append(("残った WARN", f"{len(warns)} 件（中身は、作業の記録のチェック結果）"))
    return ["# 納品の数", "", f"作成: {datetime.date.today().isoformat()}（deliver.py。納品の門を通ったときの数）" + (f"。写した先: {dest}" if dest else ""), "",
            "| 項目 | 数 |", "| --- | --- |"] + [f"| {a} | {b} |" for a, b in rows] + [""] + ([f"- {w}" for w in warns[:50]] + ([f"- ほか {len(warns) - 50} 件"] if len(warns) > 50 else []) if warns else [])


def plan_copy(work, dest):
    """写すもの: {出力先のファイル: out/ のファイル}"""
    out = work / "out"
    goods = sorted({p for pat in DELIVERABLES for p in out.glob(pat)})
    recs = sorted({p for pat in RECORDS for p in out.glob(pat)})
    plan = {dest / GOODS / p.name: p for p in goods}
    plan.update({dest / GOODS / "png" / p.name: p for p, _d in pngs(work)})
    plan.update({dest / LOG / p.name: p for p in recs})
    return plan


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--to", help="出力先のフォルダ（中に 納品物/・作業の記録/ を作る）")
    ap.add_argument("--check", action="store_true", help="門を見るだけ（書き出し直さない・写さない。納品の数と現状の処理時間の写しは out/ に書く）")
    a = ap.parse_args()
    if not a.to and not a.check:
        ap.error("--to <出力先> か --check を指定する")
    work = work_dir(a.work)
    dest = pathlib.Path(a.to).expanduser().resolve() if a.to else None
    if dest is not None and (dest == work or work in dest.parents):
        raise SystemExit("出力先は、作業フォルダの外にする（作業フォルダは、作り直しのたびに変わる）")
    if not (work / "manifest.json").exists():
        raise SystemExit("manifest.json が無い（棚卸しがまだ）")
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="asisflow_gate_"))
    try:
        for rnd in range(ROUNDS + 1):
            print(f"== 納品の門（{'見る' if rnd == 0 else f'書き出しをやり直して、もう一度 {rnd}/{ROUNDS}'}）==")
            probs, nums, warns = look(work, tmp / "numbers.json")
            if not probs:
                break
            data = [x for k, x in probs if k == "data"]
            if len(data) == len(probs) or a.check or rnd == ROUNDS:   # 書き出し直せば直る所があれば、先にやり直す（データの FAIL が、古い書き出しのせいのこともある）
                why = "データを直す所がある" if data else "--check なので、書き出し直していない" if a.check else f"{ROUNDS} 回やり直しても通らない"
                for k, x in sorted(probs, key=lambda p: p[0] != "data")[:30]:
                    print(f"FAIL  {'［データを直す］' if k == 'data' else '［書き出し直す］'}{x}")
                if len(probs) > 30:
                    print(f"FAIL  … ほか {len(probs) - 30} 件（チェック_*.txt）")
                m = load(work / "manifest.json")
                st = Status(work, m, fast=False)
                st.run()
                print("\n== 進み具合 ==\n" + "\n".join(st.rows))
                todo = [t for t in st.todo if "deliver.py" not in t]
                if todo:
                    print("\n直す工程とコマンド（直してから、もう一度 deliver.py）:\n" + "\n".join("  " + t for t in todo[:3]))
                print(f"\n門を通らない: {why}（データ {len(data)} 件・書き出し {len(probs) - len(data)} 件）")
                sys.exit(1)
            for x in [x for k, x in probs if k == "export"][:8]:
                print(f"  {x}")
            print("  書き出しをやり直す:")
            stop = chain(work)
            if stop:
                print(f"FAIL  {stop}")
                print("\n門を通らない: 書き出しが止まった（人が入れた値を守るため。理由を見て直してから、もう一度 deliver.py）")
                sys.exit(1)
        out = work / "out"
        import build_matrix as BM
        base = BM.baseline(work)                                     # 現状の処理時間の写し（BPR の比較の基準）。形を確かめてから書く
        if base is not None:
            errs = validate(base, "baseline")
            if errs:
                raise SystemExit(f"{BASELINE_FILE} の形が崩れる（{'／'.join(errs[:3])}）。build_matrix.py で書き出し直してから、もう一度")
            save(out / BASELINE_FILE, base)
            nums["times"] = dict(nums.get("times") or {}, total=base["total"], table_version=base["table_version"])
        elif (out / BASELINE_FILE).exists():
            (out / BASELINE_FILE).unlink()                           # 処理時間の行の無い様式に変えた: 前の写しを残さない
        lines = numbers_text(work, nums, warns, dest)
        (out / f"{NUMBERS_FILE}.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        save(out / f"{NUMBERS_FILE}.json", dict(nums, warn_lines=warns, date=datetime.date.today().isoformat(), to=str(dest) if dest else None))
        print("\n".join(lines[4:lines.index("", 4)]))                  # 表だけ（残った WARN の中身は、ファイル）
        if a.check:
            print("\n門を通った（--check なので、写していない）")
            return
        plan = plan_copy(work, dest)
        extra = sorted(str(p.relative_to(dest)) for folder in (dest / GOODS, dest / LOG) if folder.exists()
                       for p in folder.rglob("*") if p.is_file() and p not in plan and not p.name.startswith("."))
        if extra:
            print(f"\n出力先に、いまの一式に無いファイルがある（前の納品物の残り）: {'・'.join(extra[:8])}{' …' if len(extra) > 8 else ''}")
            raise SystemExit("写さずに止めた（消さない）。別の出力先を指定するか、残りを片付けてから、もう一度")
        for to, src in plan.items():
            to.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, to)
        print(f"\n門を通った。{dest / GOODS} に {sum(1 for p in plan if GOODS in p.parts)} ファイル・{dest / LOG} に {sum(1 for p in plan if LOG in p.parts)} ファイルを写した")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
