#!/usr/bin/env python3
"""review.py — 照合の報告（reports/）から、採否表（review/<図の番号>.review.json）の下書きを作る。

  python3 scripts/review.py draft -w <作業フォルダ>                  # 報告のある図を全部
  python3 scripts/review.py draft -w <作業フォルダ> --only F03,A-01-2
  python3 scripts/review.py snapshot -w <作業フォルダ> --only F03 --assume-current   # 前の版の採否表に、照合した版の写しを足す（利用者が認めたときだけ）

判定と指摘は、報告から機械が写す（親は写さない。写し違い・言い換えが入らない）。親が書くのは、指摘ごとの採否（decision: 採用／不採用／保留）と、
対応・理由（action）だけ。前の採否表があれば、同じ指摘（種別・対象・文が同じ）の採否と対応を引き継ぐ。中身が変わらなければ、ファイルを書き直さない。
  業務フロー     reports/review_<フローID>.json・review_<フローID>_<分冊>.json（分冊は1つにまとめる。同じノードを2回判定していれば、厳しいほう）
  アクティビティ図 reports/code_review_<先頭の番号>.json（中の activities ごとに分ける。見た箱（checked）のうち、指摘の無い箱は「支持」、指摘のある箱は「不支持」。
                 checked に無い箱は判定なし。報告に無い図・図に無い箱を指す指摘は、報告を直すまで止まる）
  差分照合       reports/review_<フローID>_差分<N>.json・code_review_<先頭の番号>_差分<N>.json（照合のあとに変わった箱だけを照合し直した報告。
                 判定した箱は、前の報告の判定を、この報告の判定で置き換える。N の小さい順に重ねる。図に無い箱を判定していれば止まる）
照合した版の写し（reviewed）: 採否表に、照合したときの箱ごとの要約（種類・名前・レーン・補足・出典のページ・前後の線）と、報告の印を残す。
新しい報告・変わった報告があれば、その報告が見た箱の写しを、いまの図で取り直す（全体の報告なら全部の箱、差分照合なら判定した箱だけ）。
点検（check_flow.py --stage review）は、いまの図と写しが違う箱を「照合していない箱」として FAIL にする（差分照合で、その箱だけを照合し直す）。
写しの無い前の版の採否表には、写しを勝手に作らない（点検は WARN。いまの図を照合済みと認めるとき・全体を照合し直したときは snapshot --assume-current）。
採否の書かれていない指摘があるうちは、点検が FAIL にする。
"""
import argparse
import datetime
import hashlib
import json
import pathlib
import re

from common import load, save, try_load, work_dir
from validate import validate

ORDER = {"支持": 0, "一部支持": 1, "不支持": 2}
NO_FINDING = "指摘なし（照合の報告で、この箱を見て、指摘が無い）"
DIFF = re.compile(r"_差分(\d+)$")                                     # 差分照合の報告（照合のあとに変わった箱だけを照合し直した）


def diff_no(f):
    """差分照合の報告なら、その番号。全体の報告なら 0"""
    m = DIFF.search(pathlib.Path(f).stem)
    return int(m.group(1)) if m else 0


def flow_reports(work, F):
    """業務フロー F の照合の報告（分冊・差分照合を含む）: [(ファイル, 中身)]。全体の報告 → 差分照合（番号の順）"""
    out = []
    for f in sorted((work / "reports").glob(f"review_{F}*.json")) if (work / "reports").exists() else []:
        rest = DIFF.sub("", f.stem)[len("review_"):].split("_")
        if rest[0] == F and (len(rest) == 1 or (len(rest) == 2 and rest[1].isdigit())):
            out.append((f, load(f)))
    return sorted(out, key=lambda x: (diff_no(x[0]), x[0].name))


def code_reports(work, aid):
    """アクティビティ図 aid を含む照合の報告: [(ファイル, 中身)]。全体の報告 → 差分照合（番号の順）"""
    out = []
    for f in sorted((work / "reports").glob("code_review_*.json")) if (work / "reports").exists() else []:
        d = load(f)
        if isinstance(d, dict) and aid in (d.get("activities") or []):
            out.append((f, d))
    return sorted(out, key=lambda x: (diff_no(x[0]), x[0].name))


def reports(work, fid):
    return code_reports(work, fid) if fid.startswith("A-") else flow_reports(work, fid)


def judged(d, fid):
    """報告 d が判定した箱の番号"""
    if fid.startswith("A-"):
        return set((d.get("checked") or {}).get(fid) or []) | {x.get("node") for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node")}
    return {v.get("node") for v in d.get("verdicts") or []}


def from_reports(work, fid, fig):
    """報告から写した (判定, 指摘)。報告が無ければ None。形の崩れた報告は ValueError（どのファイルの何が違うか）"""
    work = pathlib.Path(work)
    act = fid.startswith("A-")
    reps = reports(work, fid)
    if not reps:
        return None
    for f, d in reps:
        errs = validate(d, "report_code_review" if act else "report_review")
        if errs:
            raise ValueError(f"reports/{f.name} の形が崩れている（{'／'.join(errs[:3])}）")
    boxes = [n["id"] for n in fig.get("nodes") or [] if isinstance(n, dict) and n.get("type") not in ("start", "end")]
    for f, d in reps:
        lost = sorted(judged(d, fid) - set(boxes)) if diff_no(f) else []
        if lost:
            raise ValueError(f"reports/{f.name}（差分照合）が、{fid} に無い箱を判定している（{'・'.join(map(str, lost[:8]))}。箱の番号を直す）")
    verdict, findings = {}, []
    if act:
        for f, d in reps:
            stray = [x for x in d.get("findings") or [] if x.get("activity") not in d["activities"]]
            if stray:
                raise ValueError(f"reports/{f.name} の指摘が、報告の activities に無い図を指す（{'・'.join(sorted({str(x.get('activity')) for x in stray}))}。図の番号を直す）")
            lost = [str(x["node"]) for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node") and x["node"] not in boxes]
            if lost:
                raise ValueError(f"reports/{f.name} の指摘が、{fid} に無い箱を指す（{'・'.join(lost[:8])}。箱の番号を直す）")
        layers = [[(f, d) for f, d in reps if not diff_no(f)]] + [[(f, d)] for f, d in reps if diff_no(f)]   # 全体の報告 → 差分照合の順に重ねる
        for layer in layers:
            hits = [x for _f, d in layer for x in d.get("findings") or [] if x.get("activity") == fid]
            seen = {n for _f, d in layer for n in judged(d, fid)}
            for n in boxes:
                mine = [x for x in hits if x.get("node") == n]
                if mine:
                    verdict[n] = {"node": n, "verdict": "不支持", "evidence": " ／ ".join(x["what"] for x in mine)}
                elif n in seen:
                    verdict[n] = {"node": n, "verdict": "支持", "evidence": NO_FINDING}
            findings += [{"kind": x["kind"], "target": x.get("node") or "図にしない手続き", "text": x["what"]} for x in hits]
    else:
        for f, d in reps:
            again = {}
            for v in d.get("verdicts") or []:
                old = verdict.get(v["node"]) if not diff_no(f) else again.get(v["node"])
                if old is None or ORDER[v["verdict"]] > ORDER[old["verdict"]]:
                    again[v["node"]] = verdict[v["node"]] = {"node": v["node"], "verdict": v["verdict"], "evidence": v["evidence"]}
            findings += [{"kind": x["kind"], "target": x["target"], "text": x["text"]} for x in d.get("findings") or []]
    return list(verdict.values()), findings


def summaries(fig):
    """図の箱ごとの要約（照合した版と、いまの図を比べる）: {番号: {種類・名前・レーン・補足・推定・リンク先・出典のページ・前後の線}}"""
    lanes = {x.get("id"): x.get("name") for x in fig.get("lanes") or [] if isinstance(x, dict)}
    edges = [e for e in fig.get("edges") or [] if isinstance(e, dict)]
    out = {}
    for n in fig.get("nodes") or []:
        if not isinstance(n, dict) or n.get("type") in ("start", "end"):
            continue
        i = n.get("id")
        out[str(i)] = {"type": str(n.get("type") or ""), "label": str(n.get("label") or ""), "lane": str(lanes.get(n.get("lane"), n.get("lane")) or ""),
                       "detail": str(n.get("detail") or ""), "inferred": bool(n.get("inferred")), "link": str(n.get("link_flow") or ""),
                       "pages": sorted({f"{s.get('doc_id')} {s.get('page')}" for s in n.get("sources") or [] if isinstance(s, dict)}),
                       "edges": sorted([f"→{e.get('to')} {e.get('label') or ''}".rstrip() for e in edges if e.get("from") == i]
                                       + [f"←{e.get('from')} {e.get('label') or ''}".rstrip() for e in edges if e.get("to") == i])}
    return out


SEEN_JA = {"type": "種類", "label": "名前", "lane": "レーン", "detail": "補足", "inferred": "推定の印", "link": "リンク先", "pages": "出典のページ", "edges": "線"}


def unreviewed(rv, fig):
    """照合した版の写しと、いまの図が違う箱: [(番号, 何が違うか)]。写しが無ければ None"""
    seen = rv.get("reviewed") if isinstance(rv, dict) else None
    if not isinstance(seen, dict) or not isinstance(seen.get("nodes"), dict):
        return None
    out = []
    for i, now in summaries(fig).items():
        was = seen["nodes"].get(i)
        if not isinstance(was, dict):
            out.append((i, "照合のあとに足した"))
        elif was != now:
            out.append((i, "・".join(SEEN_JA.get(k, k) for k in now if was.get(k) != now[k]) + "が変わった"))
    return out


def report_marks(work, fid):
    """報告ごとの印（中身から作る。報告が変われば変わる）: {ファイル名: 印}"""
    return {f.name: hashlib.sha1(f.read_bytes()).hexdigest()[:12] for f, _d in reports(pathlib.Path(work), fid)}


def snapshot(work, fid, fig, old):
    """照合した版の写し（reviewed）。新しい報告・変わった報告が見た箱だけ、いまの図で取り直す。
    写しの無い前の版の採否表（old が dict で reviewed が無い）には作らない（None。どの報告が、どの版の図を見たかが分からない。
    ファイルの日時は、作業フォルダを写すと並びが変わるので、手がかりにしない）"""
    if isinstance(old, dict) and "reviewed" not in old:
        return None
    work = pathlib.Path(work)
    reps, marks, now = reports(work, fid), report_marks(work, fid), summaries(fig)
    prev = old.get("reviewed") if isinstance(old, dict) and isinstance(old.get("reviewed"), dict) else {"nodes": {}, "reports": {}}
    new = [(f, d) for f, d in reps if (prev.get("reports") or {}).get(f.name) != marks[f.name]]
    nodes = dict(now) if any(not diff_no(f) for f, _d in new) else dict(prev.get("nodes") or {})
    for f, d in new:
        if diff_no(f):
            nodes.update({i: now[i] for i in judged(d, fid) if i in now})
    at = {k: v for k, v in (prev.get("figures") or {}).items() if k in marks}
    at.update({f.name: figure_mark(fig) for f, _d in new})                # 報告を受けたときの図（採用した指摘を、図に反映したかを見る）
    out = {"nodes": {i: v for i, v in sorted(nodes.items()) if i in now}, "reports": marks, "figures": at}
    if prev.get("assumed") and not any(not diff_no(f) for f, _d in new):
        out["assumed"] = prev["assumed"]
    return out


def figure_mark(fig):
    """図のデータ全体の印（中身から作る。1 か所でも直せば変わる）"""
    return hashlib.sha1(json.dumps(fig, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:12]


def key(x):
    return (x.get("kind"), x.get("target"), x.get("text"))


def origins(work, fid):
    """指摘（種別・対象・文）→ その指摘を書いた報告のファイル名（同じ指摘が 2 つの報告にあれば、先の報告）"""
    out = {}
    for f, d in reports(pathlib.Path(work), fid):
        for x in d.get("findings") or [] if isinstance(d, dict) else []:
            if not isinstance(x, dict):
                continue
            if fid.startswith("A-"):
                if x.get("activity") == fid:
                    out.setdefault((x.get("kind"), x.get("node") or "図にしない手続き", x.get("what")), f.name)
            else:
                out.setdefault(key(x), f.name)
    return out


def draft(work, fid, fig):
    """採否表の下書き。(中身, 採否の書かれていない指摘の数)。報告が無ければ (None, 0)"""
    got = from_reports(work, fid, fig)
    if got is None:
        return None, 0
    verdicts, findings = got
    path = pathlib.Path(work) / "review" / f"{fid}.review.json"
    old = try_load(path)[0]
    kept = {key(x): x for x in (old.get("findings") or [] if isinstance(old, dict) else []) if isinstance(x, dict)}
    out = []
    for i, x in enumerate(findings, 1):
        was = kept.get(key(x)) or {}
        out.append(dict(id=f"r{i}", **x, decision=was.get("decision", ""), action=was.get("action", "")))
    open_ = sum(1 for x in out if x["decision"] not in ("採用", "不採用", "保留") or not str(x["action"]).strip())
    rv = {"flow_id": fid, "verdicts": verdicts, "findings": out}
    seen = snapshot(work, fid, fig, old)
    if seen is not None:
        rv["reviewed"] = seen
    return rv, open_


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
    ap.add_argument("cmd", choices=["draft", "snapshot"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="図の番号（カンマ区切り）")
    ap.add_argument("--assume-current", action="store_true", help="snapshot: いまの図を、照合済みとして写す（利用者が認めたときだけ）")
    a = ap.parse_args()
    work = work_dir(a.work)
    only = {x.strip() for x in a.only.split(",") if x.strip()} if a.only else None
    figs = figures(work)
    if only and not set(only) & set(figs):
        raise SystemExit(f"--only の番号が、業務フローにもアクティビティ図にも無い: {', '.join(sorted(only))}")
    if a.cmd == "snapshot":
        if not a.assume_current:
            raise SystemExit("snapshot は、いまの図を照合済みとして写す。照合のあとに図を直していれば、照合していない箱まで済みになる。"
                             "利用者が認めたときだけ、--assume-current を付けて実行する")
        made = 0
        for fid, fig in figs.items():
            path = work / "review" / f"{fid}.review.json"
            rv = try_load(path)[0]
            if (only and fid not in only) or not isinstance(rv, dict):
                continue
            gaps = unreviewed(rv, fig) or []
            rv["reviewed"] = {"nodes": summaries(fig), "reports": report_marks(work, fid), "figures": {},   # 採用した指摘は、反映済みとして認める
                              "assumed": datetime.date.today().isoformat()}
            save(path, rv)
            made += 1
            print(f"{fid}: いまの図（箱 {len(rv['reviewed']['nodes'])}）を、照合済みとして写した"
                  + (f"。前の写しと違った箱: {'・'.join(i for i, _w in gaps[:10])}" if gaps else ""))
        print(f"照合した版の写しを書いた採否表 {made}（写した日を reviewed.assumed に残した）")
        return
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
        gaps = unreviewed(rv, fig)
        print(f"{fid}: 判定 {n}（支持 {sup}）・指摘 {len(rv['findings'])}" + (f"・採否が未記入 {open_}" if open_ else "")
              + ("・照合した版の写しなし（前の版の採否表）" if gaps is None else f"・照合していない箱 {len(gaps)}（{'・'.join(i for i, _w in gaps[:8])}）" if gaps else ""))
    print(f"採否表を書いた・書き直した図 {made}。採否が未記入の指摘には、decision（採用／不採用／保留）と action（何を直したか・理由）を書く"
          "（references/source-review-prompt.md）")
    raise SystemExit(1 if bad else 0)


if __name__ == "__main__":
    main()
