#!/usr/bin/env python3
"""sc_match.py — 行の中のタスクの対応（rows/<行>.match.json）の材料・下書き・記載なしの候補（工程 5）。

  python3 scripts/sc_match.py material -w <比較の作業フォルダ> --row B01     # 行の材料: 拠点ごとの箱（鍵・担当・名前・補足・札・帳票・分岐の行き先）
  python3 scripts/sc_match.py draft -w <比較の作業フォルダ> --row B01        # 下書き（基準の拠点の作業の箱を 1 つずつタスクに。無いときだけ）
  python3 scripts/sc_match.py candidates -w <比較の作業フォルダ> --row B01   # 記載なし（箱の無い拠点）ごとの、機械の候補（absent.checked に全部書く）

候補: その拠点の、ほかのタスク・ほかの行の作業の箱と、フローから外した項目のうち、タスクの名前・ほかの拠点の箱の名前と、字の組（2 字ずつ）が
似ているもの（似ている度合い CAND_MIN 以上・多い順に CAND_MAX まで）。記載なしと書くには、候補を全部見て、どれも同じ作業でないことを確かめる。
"""
import argparse
import json
import sys

from sc_common import STEP_TYPES, likeness, match_path, save, sites_of, work_dir
from sc_rows import Model

CAND_MIN, CAND_MAX = 0.34, 6


def node_line(mdl, sid, key, edges):
    n = mdl.node(sid, key)
    if not n:
        return f"    {key}（写しに無い）"
    t = {"process": "作業", "decision": "判断", "system": "システム", "subprocess": "別のフローへ", "start": "開始", "end": "終了"}.get(n["type"], n["type"])
    parts = [f"    {key} [{t}] {mdl.role_name(sid, key)}｜{n['label']}"]
    if n.get("detail"):
        parts.append(f"（{n['detail']}）")
    if n.get("systems"):
        parts.append(" 札 " + "・".join(n["systems"]))
    if n.get("files"):
        parts.append(" 帳票 " + "・".join(f"{f['name']}{'（作る）' if f['direction'] == 'out' else ''}" for f in n["files"]))
    if n.get("inferred"):
        parts.append(" ※推定")
    outs = [e for e in edges if e[0] == key]
    if n["type"] == "decision":
        parts.append(" → " + "／".join(f"{e[2] or '—'}: {e[1]}" for e in outs))
    if n.get("src"):
        parts.append(f"｛{n['src'][0]}｝")
    return "".join(parts)


def material(mdl, rid):
    r = mdl.row_of.get(rid)
    if not r:
        sys.exit(f"行 {rid} が axis.json に無い")
    out = [f"== {rid} {r['name']}（基準: {mdl.names.get(mdl.baseline, 'なし')}）"]
    for sid in mdl.order:
        sp = mdl.site_part(rid, sid) or {}
        out.append(f"-- {mdl.names[sid]}: {sp.get('relation', '—')}" + (f"（{sp['reason']}）" if sp.get("reason") else ""))
        nodes, edges = mdl.row_nodes(rid, sid)
        for k in nodes:
            if (mdl.node(sid, k) or {}).get("type") in ("start", "end"):
                continue
            out.append(node_line(mdl, sid, k, edges))
        nxt = [f"{a}→{b}" for a, b, *_ in edges if (mdl.node(sid, a) or {}).get("type") not in ("decision", "start")
               and nodes.index(b) != nodes.index(a) + 1 and (mdl.node(sid, b) or {}).get("type") != "end"]
        if nxt:
            out.append("    並びと違う線: " + "、".join(nxt))
    shared = [x["id"] for x in mdl.rows() if x["id"] != rid and any(
        set(p["flow"] for p in (mdl.site_part(x["id"], s) or {}).get("parts") or []) & set(p["flow"] for p in (mdl.site_part(rid, s) or {}).get("parts") or [])
        for s in mdl.order)]
    if shared:
        out.append(f"同じフローを分け合う行: {', '.join(shared)}（まとめて対応を決める。箱はどれか 1 つの行のタスクにだけ入れる）")
    return out


def draft(mdl, rid):
    base = next((s for s in mdl.order if s == mdl.baseline and (mdl.site_part(rid, s) or {}).get("parts")), None) or \
        next((s for s in mdl.order if (mdl.site_part(rid, s) or {}).get("parts")), None)
    tasks = []
    for k in mdl.step_nodes(rid, base) if base else []:
        n = mdl.node(base, k)
        name = n["label"] if n["type"] != "decision" else f"{n['label'].rstrip('？?')}を判断する"
        tasks.append({"id": f"t{len(tasks) + 1}", "name": name, "members": {s: ([k] if s == base else []) for s in mdl.with_parts(rid)},
                      "absent": {}, "joined": "", "content": []})
    return {"row": rid, "tasks": tasks, "questions": []}


def task_text(mdl, rid, t, sid):
    """候補を探す手がかりの文: タスクの名前と、ほかの拠点の箱の名前"""
    words = [t.get("name", "")]
    for s, ks in (t.get("members") or {}).items():
        if s != sid:
            words += [(mdl.node(s, k) or {}).get("label", "") for k in ks or []]
    return words


def candidates(mdl, rid, t, sid):
    """記載なしの候補: [(鍵, 似ている度合い, 文, どこ)]"""
    words = task_text(mdl, rid, t, sid)
    s = mdl.snaps.get(sid) or {}
    mine = set((t.get("members") or {}).get(sid) or [])
    pool = []
    for k, n in (s.get("nodes") or {}).items():
        if n["type"] in STEP_TYPES and k not in mine:
            pool.append((k, n["label"], f"{n['flow']} {s['flows'].get(n['flow'], {}).get('name', '')}"))
    for x in s.get("excluded") or []:
        pool.append((x["key"], x["text"], f"フローから外した項目（{x['reason'][:20]}）"))
    out = []
    for k, text, where in pool:
        sc = max((likeness(w, text) for w in words if w), default=0.0)
        if sc >= CAND_MIN:
            out.append((k, round(sc, 2), text, where))
    out.sort(key=lambda x: (-x[1], x[0]))
    return out[:CAND_MAX]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["material", "draft", "candidates"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--row", required=True)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    mdl = Model(cmp, sites_of(cmp))
    if a.row not in mdl.row_of:
        sys.exit(f"行 {a.row} が axis.json に無い")
    if a.cmd == "material":
        print("\n".join(material(mdl, a.row)))
    elif a.cmd == "draft":
        p = match_path(cmp, a.row)
        if p.exists() and not a.force:
            sys.exit(f"{p.name} はもうある（作り直すときは --force。LLM が書いた対応は消える）")
        p.parent.mkdir(exist_ok=True)
        d = draft(mdl, a.row)
        save(p, d)
        print(f"タスク {len(d['tasks'])} の下書き → {p}（ほかの拠点の箱を入れ、粗い拠点に合わせてまとめる。references/match-guide.md）")
    else:
        if not mdl.match(a.row):
            sys.exit(f"{a.row}.match.json が無い")
        n = 0
        for t in mdl.tasks(a.row):
            for sid in mdl.with_parts(a.row):
                if (t.get("members") or {}).get(sid):
                    continue
                cs = candidates(mdl, a.row, t, sid)
                n += 1
                print(f"{t['id']} {t['name']} — {mdl.names[sid]} に箱が無い。checked に書く候補 {len(cs)}:")
                for k, sc, text, where in cs:
                    print(f"    {k}  {sc:.2f}  {text}（{where}）")
                print(f"    checked: {json.dumps([c[0] for c in cs], ensure_ascii=False)}")
        if not n:
            print("記載なしのタスクは無い")


if __name__ == "__main__":
    main()
