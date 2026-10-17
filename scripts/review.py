#!/usr/bin/env python3
"""review.py — 照合に出す版を控え、照合の報告（reports/）から、採否表（review/<図の番号>.review.json）の下書きを作る。

  python3 scripts/review.py draft -w <作業フォルダ>                  # 報告のある図を全部
  python3 scripts/review.py draft -w <作業フォルダ> --only F03,A-01-2
  python3 scripts/review.py record -w <作業フォルダ> --only A-01-2 [--nodes P-003,P-007]   # 照合に出す版を控え、シートの番号（S-<n>）を出す
  python3 scripts/review.py snapshot -w <作業フォルダ> --only F03 --assume-current   # 前の版の採否表の図を、照合済みと認める（利用者が認めたときだけ）

照合に出した版の記録（review/sheets.json。機械だけが書く。手で直さない）: 照合に出すたびに、箱ごとの要約（種類・名前・レーン・補足・推定の印・
リンク先・使うシステム・帳票の向きと付く箱・出典のページと行・前後の線 ほか）と、図全体の印を、シートの番号（S-<n>）で残す。
業務フローは make_compare.py が照合シートを作るときに控える（シートの見出しに番号が出る）。アクティビティ図は、照合に出す前に record で控える
（1 回に出す図の組を --only に並べる。組は同じ番号）。どちらも、書き出した図（out/*.layout.json の印）が、いまのデータと違えば控えずに止まる。
組の全部の図で、その図の最後の控えが同じ番号・同じ中身・同じ種類なら、その番号を使う（書き出し直しでは増やさない）。
その図を初めて控えるときは、そのときもうある sheet の無い報告（前の版の報告）の印も控える。
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
形の崩れた報告・照合に出していない箱を判定した報告・記録に無い番号の報告・名前が決まりと違う報告は、報告を直すまで止まる。
点検（check_flow.py --stage review）は、報告が実際に判定した版のどれとも、いまの図が違う箱を「照合したあとに変わった・足した箱」として FAIL にする
（シートに載せても、報告が判定していなければ、照合済みにしない。差分照合で、その箱だけを照合し直す）。
採用した指摘は、その指摘を受けたシートの図と、いまの図が同じなら FAIL（まだ反映していない）。
sheet の無い前の版の報告だけの採否表は、点検は WARN。前の版の報告のまま、いまの図を照合済みと認めるときだけ、
利用者が snapshot --assume-current で「認めた」版を控える（判定は、どのシートの報告よりも先に重ねる）。
照合に出したあとに書いた報告に sheet が無ければ、止まる（前の報告に番号を書き足さない）。
採否の書かれていない指摘があるうちは、点検が FAIL にする。
"""
import argparse
import datetime
import hashlib
import itertools
import json
import pathlib
import re

from common import BadJson, load, save, site, try_load, work_dir
from validate import validate

ORDER = {"支持": 0, "一部支持": 1, "不支持": 2}
NO_FINDING = "指摘なし（照合の報告で、この箱を見て、指摘が無い）"
DIFF = re.compile(r"_差分([1-9][0-9]*)$")                            # 差分照合の報告（照合のあとに変わった箱だけを照合し直した）
RECORD = "sheets.json"                                               # review/ の下。照合に出した版の記録
FIG = "図の項目"                                                      # 箱でない、図全体の項目（開始・終了条件・対象外の項目と使わないページの理由）。要約では、箱と並べて 1 つの項目にする
ASSUMED = "認めた"                                                    # 利用者が、いまの図を照合済みと認めた版（前の版の採否表）


def diff_no(f):
    """差分照合の報告なら、その番号。全体の報告なら 0"""
    m = DIFF.search(pathlib.Path(f).stem)
    return int(m.group(1)) if m else 0


def read_report(f):
    try:
        return load(f)
    except BadJson as e:
        raise ValueError(site(f"reports/{f.name} が JSON として読めない（{str(e)[-60:]}）")) from None


def flow_reports(work, F):
    """業務フロー F の照合の報告（分冊・差分照合を含む）: [(ファイル, 中身)]。名前が決まりと違う報告があれば ValueError"""
    out = []
    for f in sorted((work / "reports").glob(f"review_{F}*.json")) if (work / "reports").exists() else []:
        rest = f.stem[len("review_"):]
        if rest != F and not rest.startswith(F + "_"):
            continue                                                 # 下位フロー（F03-1）の報告
        tail = rest[len(F) + 1:]
        if rest != F and not re.fullmatch(r"[1-9][0-9]*|差分[1-9][0-9]*", tail):
            raise ValueError(site(f"reports/{f.name}: 報告の名前が決まりと違う（review_{F}.json・review_{F}_<分冊の番号>.json・review_{F}_差分<N>.json。N は 1、2…）"))
        out.append((f, read_report(f)))
    return out


def code_reports(work, aid):
    """アクティビティ図 aid を含む照合の報告: [(ファイル, 中身)]"""
    out = []
    for f in sorted((work / "reports").glob("code_review_*.json")) if (work / "reports").exists() else []:
        if "差分" in f.stem and not DIFF.search(f.stem):
            raise ValueError(site(f"reports/{f.name}: 報告の名前が決まりと違う（code_review_<図の番号>_差分<N>.json。N は 1、2…）"))
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
    出典のページと行・前後の線（推定の印も）・アクティビティ図へのリンク・付いている帳票・アクティビティ図の呼び出し・繰り返し・畳んだ手続き}}。
    図全体の項目（開始・終了条件・対象外の項目と使わないページの理由。アクティビティ図は、箱にしない行・分岐の扱いと理由）は、FIG の 1 項目にする"""
    lanes = {x.get("id"): x.get("name") for x in fig.get("lanes") or [] if isinstance(x, dict)}
    edges = [e for e in fig.get("edges") or [] if isinstance(e, dict)]
    line = lambda e, end: f"{end}{(' ' + str(e.get('label'))) if e.get('label') else ''}{'（推定）' if e.get('inferred') else ''}"
    rows = lambda key: [x for x in fig.get(key) or [] if isinstance(x, dict)]
    files = {}
    for n in fig.get("nodes") or []:
        if isinstance(n, dict) and n.get("type") == "file" and n.get("attach_to"):
            files.setdefault(n["attach_to"], []).append(f"{n.get('label', '')}（{n.get('direction', '')}）")
    out = {FIG: {"type": FIG, "label": str(fig.get("name") or ""), "lane": "", "detail": str(fig.get("summary") or ""), "inferred": False, "link": "",
                 "pages": [], "edges": [], "start": str(fig.get("start_condition") or ""), "end": str(fig.get("end_condition") or ""),
                 "excluded": sorted([f"{x.get('doc_id')} {x.get('page')} {x.get('item')}: {x.get('reason', '')}" for x in rows("excluded_items")]
                                    + [f"{x.get('key')} {x.get('state')}: {x.get('reason', '')}" for x in rows("lines") + rows("branches")
                                       if x.get("state") not in ("box", "decision")]),
                 "unused": sorted(f"{x.get('doc_id')} {x.get('page')}: {x.get('reason', '')}" for x in rows("unused_pages"))}}
    for n in fig.get("nodes") or []:
        if not isinstance(n, dict) or n.get("type") in ("start", "end"):
            continue
        i = n.get("id")
        s = {"type": str(n.get("type") or ""), "label": str(n.get("label") or ""), "lane": str(lanes.get(n.get("lane"), n.get("lane")) or ""),
             "detail": str(n.get("detail") or ""), "inferred": bool(n.get("inferred")), "link": str(n.get("link_flow") or ""),
             "pages": sorted({f"{s.get('doc_id')} {s.get('page')}" + (f" {s['lines'][0]}-{s['lines'][-1]}" if isinstance(s.get("lines"), list) and s["lines"] else "")
                              for s in n.get("sources") or [] if isinstance(s, dict)}),
             "edges": sorted([line(e, f"→{e.get('to')}") for e in edges if e.get("from") == i] + [line(e, f"←{e.get('from')}") for e in edges if e.get("to") == i])}
        for k in ("systems", "direction", "attach_to", "calls", "loop", "folded", "activity"):   # 種類によって無い項目は、写しにも入れない
            if n.get(k) not in (None, "", []):
                s[k] = json.dumps(n[k], ensure_ascii=False, sort_keys=True) if isinstance(n[k], (dict, list)) else str(n[k])
        if files.get(i):
            s["files"] = sorted(files[i])                            # 付いている帳票（帳票の箱を消すと、付いていた箱が変わる）
        out[str(i)] = s
    return out


SEEN_JA = {"type": "種類", "label": "名前", "lane": "レーン", "detail": "補足", "inferred": "推定の印", "link": "リンク先", "pages": "出典のページ", "edges": "線",
           "systems": "使うシステム", "direction": "帳票の向き", "attach_to": "付く箱", "calls": "呼び出し", "loop": "繰り返し", "folded": "畳んだ手続き",
           "activity": "アクティビティ図へのリンク", "files": "付いている帳票", "start": "開始条件", "end": "終了条件", "excluded": "対象外の項目と理由",
           "unused": "使わないページと理由"}


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
        raise ValueError(site(f"review/{RECORD}（照合に出した版の記録）が JSON として読めない（{str(e)[-60:]}）")) from None
    errs = validate(d, "sheets")
    if errs:
        raise ValueError(site(f"review/{RECORD}（照合に出した版の記録。機械だけが書く）の形が崩れている（{'／'.join(errs[:3])}）"))
    return d["sheets"]


def save_record(work, rec):
    errs = validate({"sheets": rec}, "sheets")                         # 機械の不具合で崩れた記録を書かない（作り直せない）
    if errs:
        raise ValueError(f"review/{RECORD} に書く中身の形が崩れる（{'／'.join(errs[:3])}）")
    save(pathlib.Path(work) / "review" / RECORD, {"sheets": rec})


def record(work, figs, kind, targets=None):
    """照合に出す版を控える。figs: {図の番号: 図}（1 回に出す図の組）。→ {図の番号: シートの番号}（組の中は同じ番号）。
    組の全部の図で、その図の最後の控え（認めた版も含む）が同じ番号・同じ種類・同じ図の印・同じ箱なら、その番号を使う（書き出し直しでは増やさない）。
    1 枚でも違えば、組の全部に新しい番号を付ける（報告の sheet は 1 つ）。
    その図を初めて控えるときは、sheet の無い報告（前の版の報告）の印も控える（あとから書いた sheet の無い報告と見分ける）"""
    work = pathlib.Path(work)
    if not figs:
        return {}
    rec = read_record(work)
    sid = f"S-{1 + max((int(e['id'][2:]) for e in rec), default=0)}"
    new, same = [], set()
    for F, fig in figs.items():
        now = summaries(fig)
        nodes = {i: now[i] for i in targets} if targets else now
        mark = figure_mark(fig)
        last = next((e for e in reversed(rec) if e["figure"] == F), None)
        same.add(last["id"] if last and last["kind"] == kind and last["mark"] == mark and last["nodes"] == nodes else None)
        e = {"id": sid, "date": datetime.date.today().isoformat(), "kind": kind, "figure": F, "mark": mark, "nodes": nodes}
        if last is None:
            try:
                old = {f.name: file_mark(f) for f, d in reports(work, F) if isinstance(d, dict) and not d.get("sheet")}
            except ValueError:                                       # 名前の崩れた報告は、点検・下書きが止める
                old = {}
            if old:
                e["before"] = old
        new.append(e)
    if len(same) == 1 and None not in same:
        return {F: next(iter(same)) for F in figs}
    save_record(work, rec + new)
    return {F: sid for F in figs}


def next_report(work, fid):
    """差分照合の報告の置き場（reports/review_<フローID>_差分<N>.json・code_review_<図の番号>_差分<N>.json）。N は、まだ使っていない次の数
    （前の報告を上書きしない。親が数えない）"""
    stem = ("code_review_" if fid.startswith("A-") else "review_") + fid
    used = [int(m.group(1)) for p in (pathlib.Path(work) / "reports").glob(f"{stem}_差分*.json") for m in [re.fullmatch(rf"{re.escape(stem)}_差分(\d+)", p.stem)] if m]
    return f"reports/{stem}_差分{1 + max(used, default=0)}.json"


def drawn_mark(work, fid):
    """書き出した図（out/*.layout.json）の、そのページの図の印。無ければ None"""
    page = ("act-" if fid.startswith("A-") else "flow-") + fid
    for lp in sorted((pathlib.Path(work) / "out").glob("*.layout.json")):
        lay = try_load(lp)[0]
        for pg in (lay.get("pages") if isinstance(lay, dict) and isinstance(lay.get("pages"), list) else []):
            if isinstance(pg, dict) and pg.get("id") == page:
                return pg.get("mark")
    return None


def before_marks(rec, fid):
    """前の版の報告（sheet の無い報告）の印: その図を初めて控えたときにあったもの・認めた版に入れたもの"""
    mine = [e for e in rec if e["figure"] == fid]
    out = dict(mine[0].get("before") or {}) if mine else {}
    for e in mine:
        if e["kind"] == ASSUMED:
            out.update(e.get("reports") or {})
    return out


def is_before(rec, fid, f):
    """sheet の無い報告 f が、前の版の報告か（その図を一度も控えていない・控えたときにもうあった）"""
    return not any(e["figure"] == fid for e in rec) or before_marks(rec, fid).get(f.name) == file_mark(f)


def check_shapes(fid, reps):
    """報告の形（中を読む前に）。崩れていれば ValueError（どのファイルの何が違うか）"""
    for f, d in reps:
        errs = validate(d, "report_code_review" if fid.startswith("A-") else "report_review")
        if errs:
            raise ValueError(site(f"reports/{f.name} の形が崩れている（{'／'.join(errs[:3])}）"))


def no_sheet(rec, fid, f):
    """sheet の無い報告で止めるときの文。前の版の報告には、番号を書き足させない（書き足すと、照合していない直しまで照合済みになる）"""
    if is_before(rec, fid, f):
        return (f"reports/{f.name} は、前の版の報告（照合したシートの番号 sheet が無い）。前の報告に番号を書き足さない。"
                f"いまの図を照合済みと認めるなら、利用者が python3 scripts/review.py snapshot -w <作業フォルダ> --only {fid} --assume-current で認める。"
                "認めないなら、全体を照合し直す")
    return f"reports/{f.name} に、照合したシートの番号（sheet）が無い（照合シートの見出しの［S-<n>］を写させる）"


def mapped(fid, reps, rec):
    """報告ごとの、照合に出した版: [(ファイル, 中身, 控え)]（認めた版を先に、あとは記録の順・名前の順）。
    どの報告にも sheet が無く、認めた版も無ければ None（前の版の採否表か、sheet の書き忘れ。legacy で見分ける）。
    形の崩れた報告・記録に無い番号・照合に出していない箱の判定・種類の違う名前・認めていない前の版の報告は ValueError"""
    check_shapes(fid, reps)
    mine = [e for e in rec if e["figure"] == fid]
    assumed = [e for e in mine if e["kind"] == ASSUMED]
    in_assumed = lambda f: next((e for e in reversed(assumed) if (e.get("reports") or {}).get(f.name) == file_mark(f)), None)
    if not any(d.get("sheet") for _f, d in reps) and not any(in_assumed(f) for f, _d in reps):
        return None
    at = {e["id"]: k for k, e in enumerate(rec)}
    out = []
    for f, d in reps:
        sid = d.get("sheet")
        if not sid:
            e = in_assumed(f)
            if e is None:
                raise ValueError(site(no_sheet(rec, fid, f)))
            out.append((f, d, e))
            continue
        e = next((e for e in mine if e["id"] == sid), None)
        if e is None:
            raise ValueError(site(f"reports/{f.name} のシートの番号 {sid} が、{fid} を照合に出した版の記録（review/{RECORD}）に無い（シートの見出しの番号を写す。別の図の番号でないか）"))
        if e["kind"] == ASSUMED or (e["kind"] == "差分") != bool(diff_no(f)):
            raise ValueError(site(f"reports/{f.name} のシート {sid} は「{e['kind']}」の照合（" + ("差分照合の報告は、名前を _差分<N> にする" if e["kind"] == "差分" else
                             "全体の照合の報告は、名前に _差分 を付けない") + "）"))
        lost = sorted(str(x) for x in judged(d, fid) - set(e["nodes"]))
        if lost:
            raise ValueError(site(f"reports/{f.name} が、照合に出していない箱を判定している（{'・'.join(lost[:8])}。シート {sid} に無い。箱の番号を直す）"))
        out.append((f, d, e))
    return sorted(out, key=lambda x: (x[2]["kind"] != ASSUMED, at[x[2]["id"]], x[0].name))   # 認めた版は、前の版の報告（どのシートよりも前に照合した）


def legacy(rv, fid, reps, rec):
    """前の版の採否表か: 採否表があって sheets が無く、報告がどれも前の版の報告（sheet が無く、その図を初めて控えたときにもうあった）。
    mapped が None のときだけ呼ぶ"""
    return isinstance(rv, dict) and "sheets" not in rv and all(is_before(rec, fid, f) for f, _d in reps)


def from_reports(work, fid, fig, rec=None, reps=None, maps=False):
    """報告から写した (判定, 指摘)。報告が無ければ None。形の崩れた報告・決まりと違う報告は ValueError（どのファイルの何が違うか）。
    指摘には、照合に出した版の番号（sheet）を付ける（前の版の報告には付けない）"""
    work = pathlib.Path(work)
    act = fid.startswith("A-")
    reps = reports(work, fid) if reps is None else reps
    if not reps:
        return None
    check_shapes(fid, reps)
    if maps is False:
        maps = mapped(fid, reps, read_record(work) if rec is None else rec)
    if not maps and act:                                             # 前の版の報告: 図に無い箱を指す指摘は、いまの図で見て止める
        for f, d in reps:
            lost = [str(x["node"]) for x in d.get("findings") or [] if x.get("activity") == fid and x.get("node")
                    and x["node"] not in {n.get("id") for n in fig.get("nodes") or [] if isinstance(n, dict)}]
            if lost:
                raise ValueError(site(f"reports/{f.name} の指摘が、{fid} に無い箱を指す（{'・'.join(lost[:8])}。箱の番号を直す）"))
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
                    raise ValueError(site(f"reports/{f.name} の指摘が、報告の activities に無い図を指す（{'・'.join(sorted({str(x.get('activity')) for x in stray}))}。図の番号を直す）"))
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


def unreviewed(fid, fig, maps):
    """照合していない箱: [(番号, 何が違うか)]。箱は、報告が実際に判定した版（シートに載せただけでは数えない。認めた版は全部の箱）の
    どれかと、いまの図が同じなら照合済み。どの報告も判定していない箱は、最初の全体の照合のときから変わっていなければ、
    照合し直さない（判定なしとして、支持率で数える）"""
    looked = {}                                                      # 箱 → 報告が判定した版の要約（照合に出した順）
    for _f, d, e in maps:
        for b in (set(e["nodes"]) if e["kind"] == ASSUMED else (judged(d, fid) | {FIG}) & set(e["nodes"])):   # 図の項目は、判定の欄が無い。シートに載せて報告が届けば済み
            looked.setdefault(b, []).append(e["nodes"][b])
    first = next((e for _f, _d, e in maps if e["kind"] != "差分"), None)
    last = {}                                                        # 箱 → いちばん新しく載せたシート
    for _f, _d, e in maps:
        for b in e["nodes"]:
            last[b] = e
    out = []
    for i, now in summaries(fig).items():
        seen = looked.get(i) or []
        if now in seen:
            continue
        was = seen[-1] if seen else (first or {}).get("nodes", {}).get(i)
        if not seen and was == now:
            continue
        if i in last and last[i]["nodes"][i] == now:
            out.append((i, f"シート {last[i]['id']} に載せたが、報告が判定していない"))
        elif not isinstance(was, dict):
            out.append((i, "照合のあとに足した"))
        else:
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
    if maps is None and not legacy(old, fid, reps, rec):
        raise ValueError(no_sheet(rec, fid, next((f for f, _d in reps if not is_before(rec, fid, f)), reps[0][0])))
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
        stale = [F for F in only if F in figs and F.startswith("A-") and drawn_mark(work, F) != figure_mark(figs[F])]
        if stale:
            raise SystemExit(f"図（out/アクティビティ図.drawio）が、activity.json と違う: {'・'.join(stale)}（書き出したあとに直した）。"
                             "python3 scripts/build_activity.py → check_layout.py で書き出し直してから、もう一度（照合に出す図と、控える版をそろえる）")
        try:
            got = record(work, {F: figs[F] for F in only if F in figs}, "差分" if targets else "全体", targets)
        except ValueError as e:
            raise SystemExit(str(e))
        for F, sid in got.items():
            print(f"{F}: シート {sid}（照合の報告の sheet に「{sid}」と書かせる）" + (f"。報告の置き場は {next_report(work, F)}" if targets else ""))
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
            try:
                check_shapes(fid, reps)
            except ValueError as e:
                print(f"{fid}: {e}。報告を直してから")
                continue
            old = [(f, d) for f, d in reps if not d.get("sheet")]
            new = [f.name for f, _d in old if not is_before(rec, fid, f)]
            if new:
                print(f"{fid}: {'・'.join(new)} は、照合に出したあとに書いた報告なのに、シートの番号（sheet）が無い。認めずに、報告に、照合したシートの番号を書かせる")
                continue
            if not old:
                print(f"{fid}: 前の版の報告（sheet の無い報告）が無い。認めずに、照合したあとに変わった箱を差分照合で照合し直す")
                continue
            rec = rec + [{"id": f"S-{nxt}", "date": datetime.date.today().isoformat(), "kind": ASSUMED, "figure": fid, "mark": figure_mark(fig),
                          "nodes": summaries(fig), "reports": {f.name: file_mark(f) for f, _d in old}}]
            print(f"{fid}: いまの図（箱 {len(rec[-1]['nodes'])}）を、照合済みと認めた版 S-{nxt} として控えた（前の版の報告 {len(old)}。"
                  f"判定は、どのシートの報告よりも先に重ねる）")
            nxt += 1
            made += 1
        if made:
            try:
                save_record(work, rec)
            except ValueError as e:
                raise SystemExit(str(e))
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
        gaps = unreviewed(fid, fig, maps) if maps else None
        print(f"{fid}: 判定 {n}（支持 {sup}）・指摘 {len(rv['findings'])}" + (f"・採否が未記入 {open_}" if open_ else "")
              + ("・照合に出した版の記録なし（前の版の採否表）" if gaps is None else
                 f"・照合したあとに変わった・足した箱 {len(gaps)}（{'・'.join(i for i, _w in gaps[:8])}）" if gaps else ""))
    print(f"採否表を書いた・書き直した図 {made}。採否が未記入の指摘には、decision（採用／不採用／保留）と action（何を直したか・理由）を書く"
          "（references/source-review-prompt.md）")
    raise SystemExit(1 if bad else 0)


if __name__ == "__main__":
    main()
