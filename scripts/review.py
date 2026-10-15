#!/usr/bin/env python3
"""review.py — 照合に出す版を控え、照合の報告（reports/）から、採否表（review/<図の番号>.review.json）の下書きを作る。

  python3 scripts/review.py draft -w <作業フォルダ>                  # 報告のある図を全部
  python3 scripts/review.py draft -w <作業フォルダ> --only F03,A-01-2
  python3 scripts/review.py record -w <作業フォルダ> --only A-01-2 [--nodes P-003,P-007]   # 照合に出す版を控え、シートの番号（S-<n>）を出す
  python3 scripts/review.py snapshot -w <作業フォルダ> --only F03 --assume-current   # 前の版の採否表の図を、照合済みと認める（利用者が認めたときだけ）

照合に出した版の記録（review/sheets.json。機械だけが書く。手で直さない）: 照合に出すたびに、箱ごとの要約（種類・名前・レーン・補足・推定の印・
リンク先・使うシステム・帳票の向きと付く箱・出典のページと行・前後の線 ほか）と、図全体の印を、シートの番号（S-<n>）で残す。
業務フローは make_compare.py が照合シートを作るときに控える（シートの見出しに番号が出る）。アクティビティ図は、照合に出す前に record で控える。
同じ図・同じ中身・同じ種類の最後の控えがあれば、その番号を使う（書き出し直しでは増やさない）。
照合の報告には、照合したシートの番号を sheet に書かせる。どの報告がどの版の図を見たかは、この番号で決まる
（採否表を消して作り直しても、報告を書き直しても、変わらない）。

判定と指摘は、報告から機械が写す（親は写さない。写し違い・言い換えが入らない）。親が書くのは、指摘ごとの採否（decision: 採用／不採用／保留）と、
対応・理由（action）だけ。前の採否表があれば、同じシートの同じ指摘（種別・対象・文が同じ）の採否と対応を引き継ぐ
（あとのシートで同じ指摘がまた出たら、直っていないということなので、新しい指摘として採否を書く）。中身が変わらなければ、ファイルを書き直さない。
  業務フロー     reports/review_<フローID>.json・review_<フローID>_<分冊>.json（分冊は1つにまとめる。同じノードを2回判定していれば、厳しいほう）
  アクティビティ図 reports/code_review_<先頭の番号>.json（中の activities ごとに分ける。見た箱（checked）のうち、指摘の無い箱は「支持」、指摘のある箱は「不支持」。
                 checked に無い箱は判定なし。報告に無い図を指す指摘は、報告を直すまで止まる）
  差分照合       reports/review_<フローID>_差分<N>.json・code_review_<図の番号>_差分<N>.json（照合のあとに変わった箱だけを照合し直した報告。N は 1、2…）
判定は、照合に出した順（記録の順）に重ね、あとのシートの判定が、前の判定を置き換える（全体を照合し直せば、前の差分照合より新しい）。
照合に出していない箱を判定した報告・記録に無い番号の報告・名前が決まりと違う報告は、報告を直すまで止まる。
点検（check_flow.py --stage review）は、いまの図と、各箱をいちばん新しく照合に出した版が違う箱を「照合したあとに変わった・足した箱」として FAIL にする
（差分照合で、その箱だけを照合し直す）。採用した指摘は、その指摘を受けたシートの図と、いまの図が同じなら FAIL（まだ反映していない）。
報告に sheet の無い前の版の採否表は、点検は WARN。いまの図を照合済みと認めるとき（全体を照合し直して、採用した指摘を図に反映したあとも）、
利用者が snapshot --assume-current で「認めた」版を控える（sheet のある報告のある図には使えない）。
採否の書かれていない指摘があるうちは、点検が FAIL にする。
"""
import argparse
import datetime
import hashlib
import itertools
import json
import pathlib
import re

from common import BadJson, load, save, try_load, work_dir
from validate import validate

ORDER = {"支持": 0, "一部支持": 1, "不支持": 2}
NO_FINDING = "指摘なし（照合の報告で、この箱を見て、指摘が無い）"
DIFF = re.compile(r"_差分([1-9][0-9]*)$")                            # 差分照合の報告（照合のあとに変わった箱だけを照合し直した）
RECORD = "sheets.json"                                               # review/ の下。照合に出した版の記録
ASSUMED = "認めた"                                                    # 利用者が、いまの図を照合済みと認めた版（前の版の採否表）


def diff_no(f):
    """差分照合の報告なら、その番号。全体の報告なら 0"""
    m = DIFF.search(pathlib.Path(f).stem)
    return int(m.group(1)) if m else 0


def read_report(f):
    try:
        return load(f)
    except BadJson as e:
        raise ValueError(f"reports/{f.name} が JSON として読めない（{str(e)[-60:]}）") from None


def flow_reports(work, F):
    """業務フロー F の照合の報告（分冊・差分照合を含む）: [(ファイル, 中身)]。名前が決まりと違う報告があれば ValueError"""
    out = []
    for f in sorted((work / "reports").glob(f"review_{F}*.json")) if (work / "reports").exists() else []:
        rest = f.stem[len("review_"):]
        if rest != F and not rest.startswith(F + "_"):
            continue                                                 # 下位フロー（F03-1）の報告
        tail = rest[len(F) + 1:]
        if rest != F and not re.fullmatch(r"[1-9][0-9]*|差分[1-9][0-9]*", tail):
            raise ValueError(f"reports/{f.name}: 報告の名前が決まりと違う（review_{F}.json・review_{F}_<分冊の番号>.json・review_{F}_差分<N>.json。N は 1、2…）")
        out.append((f, read_report(f)))
    return out


def code_reports(work, aid):
    """アクティビティ図 aid を含む照合の報告: [(ファイル, 中身)]"""
    out = []
    for f in sorted((work / "reports").glob("code_review_*.json")) if (work / "reports").exists() else []:
        if "差分" in f.stem and not DIFF.search(f.stem):
            raise ValueError(f"reports/{f.name}: 報告の名前が決まりと違う（code_review_<図の番号>_差分<N>.json。N は 1、2…）")
        d = read_report(f)
        if isinstance(d, dict) and aid in (d.get("activities") or []):
            out.append((f, d))
    return out


def reports(work, fid):
    return code_reports(work, fid) if fid.startswith("A-") else flow_reports(work, fid)


def judged(d, fid):
    """報告 d が判定した箱の番号"""
    if fid.startswith("A-"):
        return set((d.get("checked") or {}).get(fid) or []) | {x.get("node") for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node")}
    return {v.get("node") for v in d.get("verdicts") or []}


def summaries(fig):
    """図の箱ごとの要約（照合に出した版と、いまの図を比べる）: {番号: {種類・名前・レーン・補足・推定・リンク先・使うシステム・帳票の向きと付く箱・
    出典のページと行・前後の線（推定の印も）・アクティビティ図の呼び出し・繰り返し・畳んだ手続き}}"""
    lanes = {x.get("id"): x.get("name") for x in fig.get("lanes") or [] if isinstance(x, dict)}
    edges = [e for e in fig.get("edges") or [] if isinstance(e, dict)]
    line = lambda e, end: f"{end}{(' ' + str(e.get('label'))) if e.get('label') else ''}{'（推定）' if e.get('inferred') else ''}"
    out = {}
    for n in fig.get("nodes") or []:
        if not isinstance(n, dict) or n.get("type") in ("start", "end"):
            continue
        i = n.get("id")
        s = {"type": str(n.get("type") or ""), "label": str(n.get("label") or ""), "lane": str(lanes.get(n.get("lane"), n.get("lane")) or ""),
             "detail": str(n.get("detail") or ""), "inferred": bool(n.get("inferred")), "link": str(n.get("link_flow") or ""),
             "pages": sorted({f"{s.get('doc_id')} {s.get('page')}" + (f" {s['lines'][0]}-{s['lines'][-1]}" if isinstance(s.get("lines"), list) and s["lines"] else "")
                              for s in n.get("sources") or [] if isinstance(s, dict)}),
             "edges": sorted([line(e, f"→{e.get('to')}") for e in edges if e.get("from") == i] + [line(e, f"←{e.get('from')}") for e in edges if e.get("to") == i])}
        for k in ("systems", "direction", "attach_to", "calls", "loop", "folded"):   # 種類によって無い項目は、写しにも入れない
            if n.get(k) not in (None, "", []):
                s[k] = json.dumps(n[k], ensure_ascii=False, sort_keys=True) if isinstance(n[k], (dict, list)) else str(n[k])
        out[str(i)] = s
    return out


SEEN_JA = {"type": "種類", "label": "名前", "lane": "レーン", "detail": "補足", "inferred": "推定の印", "link": "リンク先", "pages": "出典のページ", "edges": "線",
           "systems": "使うシステム", "direction": "帳票の向き", "attach_to": "付く箱", "calls": "呼び出し", "loop": "繰り返し", "folded": "畳んだ手続き"}


def figure_mark(fig):
    """図のデータ全体の印（中身から作る。1 か所でも直せば変わる）"""
    return hashlib.sha1(json.dumps(fig, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:12]


def file_mark(f):
    return hashlib.sha1(pathlib.Path(f).read_bytes()).hexdigest()[:12]


def read_record(work):
    """照合に出した版の記録: [{id, date, kind, figure, mark, nodes, reports?}]（記録の順）。無ければ []。崩れていれば ValueError"""
    p = pathlib.Path(work) / "review" / RECORD
    if not p.exists():
        return []
    try:
        d = load(p)
    except BadJson as e:
        raise ValueError(f"review/{RECORD}（照合に出した版の記録）が JSON として読めない（{str(e)[-60:]}）") from None
    errs = validate(d, "sheets")
    if errs:
        raise ValueError(f"review/{RECORD}（照合に出した版の記録。機械だけが書く）の形が崩れている（{'／'.join(errs[:3])}）")
    return d["sheets"]


def record(work, figs, kind, targets=None):
    """照合に出す版を控える。figs: {図の番号: 図}。→ {図の番号: シートの番号}。
    同じ図で、最後の控え（認めた版を除く）が同じ種類・同じ図の印・同じ箱なら、その番号を使う（書き出し直しでは増やさない）"""
    work = pathlib.Path(work)
    rec = read_record(work)
    nxt = 1 + max((int(e["id"][2:]) for e in rec), default=0)
    out, new = {}, []
    for F, fig in figs.items():
        now = summaries(fig)
        nodes = {i: now[i] for i in targets} if targets else now
        mark = figure_mark(fig)
        last = next((e for e in reversed(rec) if e["figure"] == F and e["kind"] != ASSUMED), None)
        if last and last["kind"] == kind and last["mark"] == mark and last["nodes"] == nodes:
            out[F] = last["id"]
            continue
        new.append({"id": f"S-{nxt}", "date": datetime.date.today().isoformat(), "kind": kind, "figure": F, "mark": mark, "nodes": nodes})
        out[F] = f"S-{nxt}"
    if new:
        save(work / "review" / RECORD, {"sheets": rec + new})
    return out


def mapped(fid, reps, rec):
    """報告ごとの、照合に出した版: [(ファイル, 中身, 控え)]（記録の順・名前の順）。どの報告にも sheet が無く、認めた版も無ければ None（前の版）。
    記録に無い番号・照合に出していない箱の判定・種類の違う名前は ValueError"""
    mine = [e for e in rec if e["figure"] == fid]
    assumed = [e for e in mine if e["kind"] == ASSUMED]
    has = lambda f: any((e.get("reports") or {}).get(f.name) == file_mark(f) for e in assumed)
    if not any(isinstance(d, dict) and d.get("sheet") for _f, d in reps) and not any(has(f) for f, _d in reps):
        return None
    at = {e["id"]: k for k, e in enumerate(rec)}
    out = []
    for f, d in reps:
        sid = d.get("sheet") if isinstance(d, dict) else None
        if not sid:
            e = next((e for e in reversed(assumed) if (e.get("reports") or {}).get(f.name) == file_mark(f)), None)
            if e is None:
                raise ValueError(f"reports/{f.name} に、照合したシートの番号（sheet）が無い（照合シートの見出しの［S-<n>］を写させる）")
            out.append((f, d, e))
            continue
        e = next((e for e in mine if e["id"] == sid), None)
        if e is None:
            raise ValueError(f"reports/{f.name} のシートの番号 {sid} が、{fid} を照合に出した版の記録（review/{RECORD}）に無い（シートの見出しの番号を写す。別の図の番号でないか）")
        if e["kind"] == ASSUMED or (e["kind"] == "差分") != bool(diff_no(f)):
            raise ValueError(f"reports/{f.name} のシート {sid} は「{e['kind']}」の照合（" + ("差分照合の報告は、名前を _差分<N> にする" if e["kind"] == "差分" else
                             "全体の照合の報告は、名前に _差分 を付けない") + "）")
        lost = sorted(str(x) for x in judged(d, fid) - set(e["nodes"]))
        if lost:
            raise ValueError(f"reports/{f.name} が、照合に出していない箱を判定している（{'・'.join(lost[:8])}。シート {sid} に無い。箱の番号を直す）")
        out.append((f, d, e))
    return sorted(out, key=lambda x: (at[x[2]["id"]], x[0].name))


def legacy(rv, maps):
    """前の版の採否表か: 報告に sheet が無く（認めた版も無く）、採否表があって sheets が無い"""
    return maps is None and isinstance(rv, dict) and "sheets" not in rv


def from_reports(work, fid, fig, rec=None, reps=None, maps=False):
    """報告から写した (判定, 指摘)。報告が無ければ None。形の崩れた報告・決まりと違う報告は ValueError（どのファイルの何が違うか）。
    指摘には、照合に出した版の番号（sheet）を付ける（前の版の報告には付けない）"""
    work = pathlib.Path(work)
    act = fid.startswith("A-")
    reps = reports(work, fid) if reps is None else reps
    if not reps:
        return None
    for f, d in reps:
        errs = validate(d, "report_code_review" if act else "report_review")
        if errs:
            raise ValueError(f"reports/{f.name} の形が崩れている（{'／'.join(errs[:3])}）")
    if maps is False:
        maps = mapped(fid, reps, read_record(work) if rec is None else rec)
    if not maps and act:                                             # 前の版の報告: 図に無い箱を指す指摘は、いまの図で見て止める
        for f, d in reps:
            lost = [str(x["node"]) for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node")
                    and x["node"] not in {n.get("id") for n in fig.get("nodes") or [] if isinstance(n, dict)}]
            if lost:
                raise ValueError(f"reports/{f.name} の指摘が、{fid} に無い箱を指す（{'・'.join(lost[:8])}。箱の番号を直す）")
    layers = [[(f, d) for f, d, _e in grp] for _e, grp in itertools.groupby(maps, key=lambda x: x[2]["id"])] if maps else [reps]
    sheet = {f.name: e["id"] for f, _d, e in maps} if maps else {}
    boxes = [n["id"] for n in fig.get("nodes") or [] if isinstance(n, dict) and n.get("type") not in ("start", "end")]
    verdict, findings = {}, []
    for layer in layers:                                             # 照合に出した順に重ねる（あとのシートの判定が、前の判定を置き換える）
        mine = {}
        if act:
            for f, d in layer:
                stray = [x for x in d.get("findings") or [] if x.get("activity") not in d["activities"]]
                if stray:
                    raise ValueError(f"reports/{f.name} の指摘が、報告の activities に無い図を指す（{'・'.join(sorted({str(x.get('activity')) for x in stray}))}。図の番号を直す）")
            hits = [(f, x) for f, d in layer for x in d.get("findings") or [] if x.get("activity") == fid]
            seen = {n for _f, d in layer for n in judged(d, fid)}
            for n in [b for b in boxes if b in seen] + sorted(seen - set(boxes), key=str):
                said = [x for _f, x in hits if x.get("node") == n]
                mine[n] = ({"node": n, "verdict": "不支持", "evidence": " ／ ".join(x["what"] for x in said)} if said
                           else {"node": n, "verdict": "支持", "evidence": NO_FINDING})
            findings += [dict({"kind": x["kind"], "target": x.get("node") or "図にしない手続き", "text": x["what"]}, **({"sheet": sheet[f.name]} if f.name in sheet else {}))
                         for f, x in hits]
        else:
            for f, d in layer:
                for v in d.get("verdicts") or []:
                    old = mine.get(v["node"])
                    if old is None or ORDER[v["verdict"]] > ORDER[old["verdict"]]:
                        mine[v["node"]] = {"node": v["node"], "verdict": v["verdict"], "evidence": v["evidence"]}
                findings += [dict({"kind": x["kind"], "target": x["target"], "text": x["text"]}, **({"sheet": sheet[f.name]} if f.name in sheet else {}))
                             for x in d.get("findings") or []]
        verdict.update(mine)
    return list(verdict.values()), findings


def unreviewed(fig, maps):
    """いまの図と、各箱をいちばん新しく照合に出した版（報告の届いた版）が違う箱: [(番号, 何が違うか)]"""
    seen = {}
    for _f, _d, e in maps:
        seen.update(e["nodes"])
    out = []
    for i, now in summaries(fig).items():
        was = seen.get(i)
        if not isinstance(was, dict):
            out.append((i, "照合のあとに足した"))
        elif was != now:
            out.append((i, "・".join(SEEN_JA.get(k, k) for k in sorted(set(now) | set(was)) if was.get(k) != now.get(k)) + "が変わった"))
    return out


def key(x):
    return (x.get("kind"), x.get("target"), x.get("text"))


def draft(work, fid, fig, rec=None):
    """採否表の下書き。(中身, 採否の書かれていない指摘の数)。報告が無ければ (None, 0)"""
    work = pathlib.Path(work)
    reps = reports(work, fid)
    if not reps:
        return None, 0
    path = work / "review" / f"{fid}.review.json"
    old = try_load(path)[0]
    rec = read_record(work) if rec is None else rec
    maps = mapped(fid, reps, rec)
    if maps is None and not legacy(old, maps):
        raise ValueError(f"報告（{'・'.join(f.name for f, _d in reps[:4])}）に、照合したシートの番号（sheet）が無い（照合シートの見出しの［S-<n>］を写させる）")
    verdicts, findings = from_reports(work, fid, fig, rec, reps, maps)
    kinds = {e["id"]: e["kind"] for e in rec}
    kept = {(x.get("sheet", ""), key(x)): x for x in (old.get("findings") or [] if isinstance(old, dict) else []) if isinstance(x, dict)}
    out = []
    for i, x in enumerate(findings, 1):
        was = kept.get((x.get("sheet", ""), key(x))) or (kept.get(("", key(x))) if kinds.get(x.get("sheet")) == ASSUMED else None) or {}
        out.append(dict(id=f"r{i}", **x, decision=was.get("decision", ""), action=was.get("action", "")))
    open_ = sum(1 for x in out if x["decision"] not in ("採用", "不採用", "保留") or not str(x["action"]).strip())
    rv = {"flow_id": fid, "verdicts": verdicts, "findings": out}
    if maps is not None:
        rv["sheets"] = {f.name: e["id"] for f, _d, e in maps}
    filled = [dict(x, decision=x["decision"] if x["decision"] in ("採用", "不採用", "保留") else "保留", action=str(x["action"]).strip() or "未記入") for x in out]
    errs = validate(dict(rv, findings=filled), "review")                 # 採否の空欄は、下書きでは誤りでない（人が書く。点検が「未記入」として止める）
    if errs:
        raise ValueError(f"採否表の下書きの形が崩れる（{'／'.join(errs[:3])}）")
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


def write_draft(work, fid, fig, rec=None):
    """下書きを作って、変わっていれば書く。→ (中身, 採否の書かれていない指摘の数, 書いたか)"""
    rv, open_ = draft(work, fid, fig, rec)
    if rv is None:
        return None, 0, False
    path = pathlib.Path(work) / "review" / f"{fid}.review.json"
    old = try_load(path)[0]
    if json.dumps(old, ensure_ascii=False, sort_keys=True) == json.dumps(rv, ensure_ascii=False, sort_keys=True):
        return rv, open_, False
    path.parent.mkdir(parents=True, exist_ok=True)
    save(path, rv)
    return rv, open_, True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft", "record", "snapshot"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="図の番号（カンマ区切り）")
    ap.add_argument("--nodes", help="record: 差分照合に出す箱の番号（カンマ区切り。--only に 1 枚だけ）")
    ap.add_argument("--assume-current", action="store_true", help="snapshot: いまの図を、照合済みと認める（利用者が認めたときだけ）")
    a = ap.parse_args()
    work = work_dir(a.work)
    only = [x.strip() for x in a.only.split(",") if x.strip()] if a.only else None
    figs = figures(work)
    if only and not set(only) & set(figs):
        raise SystemExit(f"--only の番号が、業務フローにもアクティビティ図にも無い: {', '.join(sorted(only))}")
    try:
        rec = read_record(work)
    except ValueError as e:
        raise SystemExit(str(e))
    if a.cmd == "record":
        targets = [x.strip() for x in a.nodes.split(",") if x.strip()] if a.nodes else None
        if not only or (targets is not None and (len(only) != 1 or not targets)):
            raise SystemExit("record は --only に図を指定して使う（--nodes を付けるときは、図を 1 枚だけ・箱の番号を 1 つ以上）")
        if targets:
            have = summaries(figs[only[0]]) if only[0] in figs else {}
            bad = [x for x in targets if x not in have]
            if bad:
                raise SystemExit(f"--nodes の番号が、{only[0]} の箱に無い（開始・終了は照合しない）: {'・'.join(bad)}")
        got = record(work, {F: figs[F] for F in only if F in figs}, "差分" if targets else "全体", targets)
        for F, sid in got.items():
            print(f"{F}: シート {sid}（照合の報告の sheet に「{sid}」と書かせる）")
        return
    if a.cmd == "snapshot":
        if not a.assume_current:
            raise SystemExit("snapshot は、いまの図を照合済みと認める。照合のあとに図を直していれば、照合していない箱まで済みになる。"
                             "利用者が認めたときだけ、--assume-current を付けて実行する")
        made = 0
        nxt = 1 + max((int(e["id"][2:]) for e in rec), default=0)
        for fid, fig in figs.items():
            if only and fid not in only:
                continue
            try:
                reps = reports(work, fid)
            except ValueError as e:
                print(f"{fid}: {e}。報告を直してから")
                continue
            if not reps:
                continue
            if any(isinstance(d, dict) and d.get("sheet") for _f, d in reps):
                print(f"{fid}: シートの番号（sheet）のある報告がある（照合に出した版の記録がある）。認めずに、照合したあとに変わった箱を差分照合で照合し直す")
                continue
            rec = rec + [{"id": f"S-{nxt}", "date": datetime.date.today().isoformat(), "kind": ASSUMED, "figure": fid, "mark": figure_mark(fig),
                          "nodes": summaries(fig), "reports": {f.name: file_mark(f) for f, _d in reps}}]
            print(f"{fid}: いまの図（箱 {len(rec[-1]['nodes'])}）を、照合済みと認めた版 S-{nxt} として控えた（報告 {len(reps)}）")
            nxt += 1
            made += 1
        if made:
            save(work / "review" / RECORD, {"sheets": rec})
            for fid in [e["figure"] for e in rec[-made:]]:
                try:
                    write_draft(work, fid, figs[fid], rec)
                except ValueError as e:
                    print(f"{fid}: {e}")
        print(f"照合済みと認めた図 {made}（review/{RECORD} に、認めた日と「{ASSUMED}」を残した。採否表も写し直した）")
        return
    made = bad = 0
    for fid, fig in figs.items():
        if only and fid not in only:
            continue
        try:
            rv, open_, wrote = write_draft(work, fid, fig, rec)
        except ValueError as e:
            print(f"{fid}: {e}。報告を直してから")
            bad += 1
            continue
        if rv is None:
            continue
        made += wrote
        n = len(rv["verdicts"])
        sup = sum(v["verdict"] == "支持" for v in rv["verdicts"])
        maps = mapped(fid, reports(work, fid), rec)
        gaps = unreviewed(fig, maps) if maps else None
        print(f"{fid}: 判定 {n}（支持 {sup}）・指摘 {len(rv['findings'])}" + (f"・採否が未記入 {open_}" if open_ else "")
              + ("・照合に出した版の記録なし（前の版の採否表）" if gaps is None else
                 f"・照合したあとに変わった・足した箱 {len(gaps)}（{'・'.join(i for i, _w in gaps[:8])}）" if gaps else ""))
    print(f"採否表を書いた・書き直した図 {made}。採否が未記入の指摘には、decision（採用／不採用／保留）と action（何を直したか・理由）を書く"
          "（references/source-review-prompt.md）")
    raise SystemExit(1 if bad else 0)


if __name__ == "__main__":
    main()
