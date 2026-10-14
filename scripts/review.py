#!/usr/bin/env python3
"""review.py — 照合の報告（reports/）から、採否表（review/<図の番号>.review.json）の下書きを作る。

  python3 scripts/review.py draft -w <作業フォルダ>                  # 報告のある図を全部
  python3 scripts/review.py draft -w <作業フォルダ> --only F03,A-01-2

判定と指摘は、報告から機械が写す（親は写さない。写し違い・言い換えが入らない）。親が書くのは、指摘ごとの採否（decision: 採用／不採用／保留）と、
対応・理由（action）だけ。前の採否表があれば、同じ指摘（種別・対象・文が同じ）の採否と対応を引き継ぐ。中身が変わらなければ、ファイルを書き直さない。
  業務フロー     reports/review_<フローID>.json・review_<フローID>_<分冊>.json（分冊は1つにまとめる。同じノードを2回判定していれば、厳しいほう）
  アクティビティ図 reports/code_review_<先頭の番号>.json（中の activities ごとに分ける。見た箱（checked）のうち、指摘の無い箱は「支持」、指摘のある箱は「不支持」。
                 checked に無い箱は判定なし。報告に無い図・図に無い箱を指す指摘は、報告を直すまで止まる）
採否の書かれていない指摘があるうちは、点検（check_flow.py --stage review）が FAIL にする。
"""
import argparse
import json
import pathlib

from common import load, save, try_load, work_dir
from validate import validate

ORDER = {"支持": 0, "一部支持": 1, "不支持": 2}
NO_FINDING = "指摘なし（照合の報告で、この箱を見て、指摘が無い）"


def flow_reports(work, F):
    """業務フロー F の照合の報告（分冊を含む）: [(ファイル, 中身)]"""
    out = []
    for f in sorted((work / "reports").glob(f"review_{F}*.json")) if (work / "reports").exists() else []:
        rest = f.stem[len("review_"):].split("_")
        if rest[0] == F and (len(rest) == 1 or (len(rest) == 2 and rest[1].isdigit())):
            out.append((f, load(f)))
    return out


def code_reports(work, aid):
    """アクティビティ図 aid を含む照合の報告: [(ファイル, 中身)]"""
    out = []
    for f in sorted((work / "reports").glob("code_review_*.json")) if (work / "reports").exists() else []:
        d = load(f)
        if isinstance(d, dict) and aid in (d.get("activities") or []):
            out.append((f, d))
    return out


def from_reports(work, fid, fig):
    """報告から写した (判定, 指摘)。報告が無ければ None。形の崩れた報告は ValueError（どのファイルの何が違うか）"""
    work = pathlib.Path(work)
    act = fid.startswith("A-")
    reps = code_reports(work, fid) if act else flow_reports(work, fid)
    if not reps:
        return None
    for f, d in reps:
        errs = validate(d, "report_code_review" if act else "report_review")
        if errs:
            raise ValueError(f"reports/{f.name} の形が崩れている（{'／'.join(errs[:3])}）")
    verdict, findings = {}, []
    if act:
        boxes = [n["id"] for n in fig.get("nodes") or [] if isinstance(n, dict) and n.get("type") not in ("start", "end")]
        for f, d in reps:
            stray = [x for x in d.get("findings") or [] if x.get("activity") not in d["activities"]]
            if stray:
                raise ValueError(f"reports/{f.name} の指摘が、報告の activities に無い図を指す（{'・'.join(sorted({str(x.get('activity')) for x in stray}))}。図の番号を直す）")
            lost = [str(x["node"]) for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node") and x["node"] not in boxes]
            if lost:
                raise ValueError(f"reports/{f.name} の指摘が、{fid} に無い箱を指す（{'・'.join(lost[:8])}。箱の番号を直す）")
        hits = [x for _f, d in reps for x in d.get("findings") or [] if x.get("activity") == fid]
        seen = {n for _f, d in reps for n in (d.get("checked") or {}).get(fid) or []}
        for n in boxes:
            mine = [x for x in hits if x.get("node") == n]
            if mine:
                verdict[n] = {"node": n, "verdict": "不支持", "evidence": " ／ ".join(x["what"] for x in mine)}
            elif n in seen:
                verdict[n] = {"node": n, "verdict": "支持", "evidence": NO_FINDING}
        findings = [{"kind": x["kind"], "target": x.get("node") or "図にしない手続き", "text": x["what"]} for x in hits]
    else:
        for _f, d in reps:
            for v in d.get("verdicts") or []:
                old = verdict.get(v["node"])
                if old is None or ORDER[v["verdict"]] > ORDER[old["verdict"]]:
                    verdict[v["node"]] = {"node": v["node"], "verdict": v["verdict"], "evidence": v["evidence"]}
            findings += [{"kind": x["kind"], "target": x["target"], "text": x["text"]} for x in d.get("findings") or []]
    return list(verdict.values()), findings


def key(x):
    return (x.get("kind"), x.get("target"), x.get("text"))


def draft(work, fid, fig):
    """採否表の下書き。(中身, 採否の書かれていない指摘の数)。報告が無ければ (None, 0)"""
    got = from_reports(work, fid, fig)
    if got is None:
        return None, 0
    verdicts, findings = got
    old = try_load(work / "review" / f"{fid}.review.json")[0]
    kept = {key(x): x for x in (old.get("findings") or [] if isinstance(old, dict) else []) if isinstance(x, dict)}
    out = []
    for i, x in enumerate(findings, 1):
        was = kept.get(key(x)) or {}
        out.append(dict(id=f"r{i}", **x, decision=was.get("decision", ""), action=was.get("action", "")))
    open_ = sum(1 for x in out if x["decision"] not in ("採用", "不採用", "保留") or not str(x["action"]).strip())
    return {"flow_id": fid, "verdicts": verdicts, "findings": out}, open_


def tally(figs, reviews):
    """支持率の集計（業務フローとアクティビティ図を分ける）。reviews: 図の番号 → 採否表の中身。
    → {"flow"|"activity": {"figs": 図の数, "nodes": 開始・終了を除くノードの数, "judged": 判定したノードの数, "tot": 判定ごとの数, "dropped": 判定のあとに無くしたノードの数, "rest": 照合していない図}}"""
    import collections
    out = {}
    for kind in ("flow", "activity"):
        mine = {F: fig for F, fig in figs.items() if F.startswith("A-") == (kind == "activity")}
        nodes = {(F, n.get("id")) for F, fig in mine.items() for n in fig.get("nodes") or [] if isinstance(n, dict) and n.get("type") not in ("start", "end")}
        verdict = {}                                                 # (図, ノード) → 判定。同じノードを2回数えない
        for F, rv in reviews.items():
            if F in mine and isinstance(rv, dict):
                for v in rv.get("verdicts") if isinstance(rv.get("verdicts"), list) else []:
                    if isinstance(v, dict):
                        verdict[(F, v.get("node"))] = v.get("verdict") if v.get("verdict") in ORDER else "その他"
        out[kind] = {"figs": len(mine), "nodes": len(nodes), "judged": len(set(verdict) & nodes),
                     "tot": collections.Counter(v for k, v in verdict.items() if k in nodes),   # 指摘を受けて無くしたノードの判定は数えない（支持率が上に振れる）
                     "dropped": len(set(verdict) - nodes), "rest": sorted(set(mine) - set(reviews))}
    return out


def figures(work):
    """図の番号 → 図のデータ（業務フローの flow.json・アクティビティ図の activity.json）"""
    from common import load_flows, flow_key
    out = {flow_key(fl): fl for fl in load_flows(work)}
    for f in sorted((work / "activities").glob("*.activity.json")) if (work / "activities").exists() else []:
        a = try_load(f)[0]
        if isinstance(a, dict) and a.get("flow_id"):
            out[a["flow_id"]] = a
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="図の番号（カンマ区切り）")
    a = ap.parse_args()
    work = work_dir(a.work)
    only = {x.strip() for x in a.only.split(",") if x.strip()} if a.only else None
    figs = figures(work)
    made = bad = 0
    for fid, fig in figs.items():
        if only and fid not in only:
            continue
        try:
            rv, open_ = draft(work, fid, fig)
        except ValueError as e:
            print(f"{fid}: {e}。報告を直してから")
            bad += 1
            continue
        if rv is None:
            continue
        path = work / "review" / f"{fid}.review.json"
        old = try_load(path)[0]
        if json.dumps(old, ensure_ascii=False, sort_keys=True) != json.dumps(rv, ensure_ascii=False, sort_keys=True):
            path.parent.mkdir(parents=True, exist_ok=True)
            save(path, rv)
            made += 1
        n = len(rv["verdicts"])
        sup = sum(v["verdict"] == "支持" for v in rv["verdicts"])
        print(f"{fid}: 判定 {n}（支持 {sup}）・指摘 {len(rv['findings'])}" + (f"・採否が未記入 {open_}" if open_ else ""))
    if only and not set(only) & set(figs):
        raise SystemExit(f"--only の番号が、業務フローにもアクティビティ図にも無い: {', '.join(sorted(only))}")
    print(f"採否表を書いた・書き直した図 {made}。採否が未記入の指摘には、decision（採用／不採用／保留）と action（何を直したか・理由）を書く"
          "（references/source-review-prompt.md）")
    raise SystemExit(1 if bad else 0)


if __name__ == "__main__":
    main()
