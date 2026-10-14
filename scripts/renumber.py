#!/usr/bin/env python3
"""renumber.py — フローのノード番号を、時系列の順に機械が振り直す（P-001・D-001・S-001・F-001・SUB-001）。

  python3 scripts/renumber.py -w <作業フォルダ>                 # 全フロー
  python3 scripts/renumber.py -w <作業フォルダ> --only F03      # フローを選ぶ
  python3 scripts/renumber.py -w <作業フォルダ> --dry-run       # 何が変わるかだけ見る
  python3 scripts/renumber.py -w <作業フォルダ> --lanes         # JSON をまるごと書いたフローのレーンも、決まり 4.2 で並べ直す（付けたときだけ）

番号は LLM に数えさせない（飛び・重なり・順不同が出て、実行のたびに変わる）。フローを書いたあと・ノードを足し引きしたあとに実行する。
あわせて、作業のノードの推定の印（inferred）を出典から付け直す: カードの項目（item）を1つも引かないノードが推定（原典に無いノード）。
札（systems: その作業で使うシステムの名前）も、出典に引いたカードの手順の system から付け直す。
決まりで決まる所も当てる（common.settle_flow）: 判断の名前の「？」→「か」・書かれていない側のラベル（反対の語）・空の終了条件・確認事項の重要度（決まり 5.4）。
プログラムの起動点の箱（S-。activity つき）も、システム台帳の紐付けから置き直す（名前は、アクティビティ図の名前）。
順番: 開始から線をたどった深さ（戻りの線は数えない）の浅い順。同じ深さでは nodes に書いた順。ファイルは、付け先のノードの順。
書き換えるもの: nodes[].id・attach_to、edges[].from / to、questions[].node / nodes / edge。
変わらないもの: start・end、確認事項の id と通し番号（Q-…）、出典。
人が入れた処理時間（times.json。業務マトリクス一覧から取り込んだ値）の番号も、一緒に振り直す。
原典照合レビュー（review/*.review.json）を始めたあとは振り直さない（採否表のノード番号とずれる）。
"""
import argparse
import re

from common import CardItems, Times, flow_heads, infer_nodes, order_lanes, order_of, save, settle_flow, system_names, tag_systems, try_load, work_dir
from validate import blocking, validate

PREFIX = {"process": "P", "decision": "D", "system": "S", "file": "F", "subprocess": "SUB"}
# 振り直しに関わる形の違反（番号・種別・線の端・確認事項の指す先、必須の項目の欠け、型の違い）。余分な項目などは、振り直しを止めない
RISKY = re.compile(r"\.(id|type|attach_to|from|to|node|edge)(\[\d+\])?: |必須の項目が無い|型が違う")


def relane(fl):
    """JSON をまるごと書いたフローのレーンを、決まり 4.2 で並べ直す（開始の担当が左端・システムが右端・線がまたぐレーンの数が最小）。
    変えたら (前の並び, 後の並び)、変わらなければ None"""
    lanes = [x for x in fl.get("lanes") or [] if isinstance(x, dict)]
    name = {x.get("id"): x.get("name") for x in lanes}
    at = {n.get("id"): name.get(n.get("lane")) for n in fl.get("nodes") or [] if isinstance(n, dict)}
    names = [x.get("name") for x in lanes]
    systems = [x.get("name") for x in lanes if x.get("kind") == "system"]
    start = at.get("start")
    new = order_lanes(names, [(at.get(e.get("from")), at.get(e.get("to"))) for e in fl.get("edges") or [] if isinstance(e, dict)],
                      keep=[] if start in systems or start is None else [start], right=systems)
    if new == names:
        return None
    fl["lanes"] = sorted(lanes, key=lambda x: new.index(x.get("name")))
    return names, new


def renumber(fl):
    depth = order_of(fl)
    idx = {n["id"]: k for k, n in enumerate(fl["nodes"])}
    key = lambda n: (depth.get(n.get("attach_to"), 10 ** 6) if n.get("type") == "file" else depth[n["id"]], idx[n["id"]])
    ren, count = {}, {}
    for n in sorted(fl["nodes"], key=key):
        pre = PREFIX.get(n.get("type"))
        if not pre:
            continue
        count[pre] = count.get(pre, 0) + 1
        ren[n["id"]] = f"{pre}-{count[pre]:03d}"
    ren = {a: b for a, b in ren.items() if a != b}
    if not ren:
        return {}
    if set(ren.values()) & ({n["id"] for n in fl["nodes"]} - set(ren)):
        raise ValueError("振り直す番号が、振り直さないノードの番号と重なる")   # 線・確認事項が、どちらのノードを指すのか分からなくなる
    apply_ids(fl, ren)
    return ren


def apply_ids(fl, ren):
    """番号の付け替え（ren: 前の番号 → 新しい番号）を、ノード・帳票の付け先・線・確認事項に当てる"""
    r = lambda x: ren.get(x, x)
    for n in fl["nodes"]:
        n["id"] = r(n["id"])
        if n.get("attach_to"):
            n["attach_to"] = r(n["attach_to"])
    for e in fl.get("edges") or []:
        e["from"], e["to"] = r(e.get("from")), r(e.get("to"))
    for q in fl.get("questions") or []:
        if q.get("node"):
            q["node"] = r(q["node"])
        if isinstance(q.get("nodes"), list):
            q["nodes"] = [r(x) for x in q["nodes"]]
        if q.get("edge"):
            q["edge"] = [r(x) for x in q["edge"]]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--only", help="フロー ID（カンマ区切り）")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--lanes", action="store_true", help="JSON をまるごと書いたフローのレーンも、決まり 4.2 で並べ直す（直しの一覧のフローは apply_edits.py が並べる）")
    a = ap.parse_args()
    work = work_dir(a.work)
    blocking(work)
    only = set(a.only.split(",")) if a.only else None
    changed = 0
    items_of, names = CardItems(work), system_names(work)
    from systems import Ledger, has_ledger, place_entries, placed
    led = Ledger(work) if has_ledger(work) else None
    times, moved = Times(work), 0
    tax = try_load(work / "taxonomy.json")[0]
    cats = {c.get("id"): c for c in (tax.get("categories") or [] if isinstance(tax, dict) else []) if isinstance(c, dict)}
    for f in sorted((work / "flows").glob("*.flow.json")):
        fl, err = try_load(f)
        if err or not isinstance(fl, dict):
            print(f"{f.name}: 読めないので振り直さない（{err or '全体がオブジェクトになっていない'}）")
            continue
        if only and fl.get("flow_id") not in only:
            continue
        errs = [e for e in validate(fl, "flow") if RISKY.search(e)]  # 形が崩れたまま振り直すと、番号が重なって元に戻せなくなる（type の書き違いなど）
        if errs:
            print(f"{f.name}: 形の違反が {len(errs)} 件あるので振り直さない（先に check_flow.py --stage flow を通す。例: {errs[0][:60]}）")
            continue
        cat = cats.get(fl.get("flow_id"))
        heads = {k: x for k, x in flow_heads(cat).items() if fl.get(k) != x} if cat else {}
        if heads:                                                    # 名前・番号・親・段・準拠は、業務分類の写し（分類を直したら、ここで付いていく）
            print(f"{f.name}: 見出しを業務分類から{'写す' if a.dry_run else '写した'}（" + "・".join(f"{k}「{fl.get(k)}」→「{x}」" for k, x in heads.items()) + "）")
            fl.update(heads)
        lanes = a.lanes and not (work / "flows" / f"{fl.get('flow_id')}.edits.json").exists() and relane(fl)
        if lanes:
            print(f"{f.name}: レーンを{'並べ直す' if a.dry_run else '並べ直した'}（{'・'.join(lanes[0])} → {'・'.join(lanes[1])}）")
        marks = infer_nodes(fl)                                      # 推定の印は、採否表があっても付け直す（番号は変わらない）
        if marks:
            print(f"{f.name}: 作業のノードの推定の印を{'付け直す' if a.dry_run else '付け直した'}（"
                  + "、".join(f"{i} {'推定にする' if on else '推定を外す'}" for i, on in marks[:8]) + f"{' …' if len(marks) > 8 else ''}）"
                  + ("。推定にしたノードには、それを指す確認事項が要る" if any(on for _i, on in marks) else ""))
        tags = tag_systems(fl, items_of, names)                      # 札も、採否表があっても付け直す
        if tags:
            print(f"{f.name}: 札（使うシステム）を{'付け直す' if a.dry_run else '付け直した'}（"
                  + "、".join(f"{i} {'・'.join(x) or '札なし'}" for i, x in tags[:8]) + f"{' …' if len(tags) > 8 else ''}）")
        boxes = led is not None and place_entries(fl, led)           # 起動点の箱（番号は、置いた中で振る。採否表があれば、空いている番号を付ける）
        if boxes:
            print(f"{f.name}: プログラムの起動点の箱を{'置き直す' if a.dry_run else '置き直した'}（"
                  + ("、".join(f"{x['id']} ▶ {x['activity']}" for x in placed(fl)) or "箱なし") + "）")
        ren = {}
        if (work / "review" / f"{fl.get('flow_id')}.review.json").exists():
            print(f"{f.name}: 原典照合レビューの採否表があるので振り直さない")
        else:
            try:
                ren = renumber(fl)
            except ValueError as e:
                print(f"{f.name}: {e}ので振り直さない（先に check_flow.py --stage flow を通す）")
                ren = None
        if ren:
            show = "、".join(f"{a_}→{b_}" for a_, b_ in sorted(ren.items(), key=lambda x: (re.sub(r"\d", "", x[0]), x[0]))[:8])
            print(f"{f.name}: {len(ren)} 個の番号を振り直{'す' if a.dry_run else 'した'}（{show}{' …' if len(ren) > 8 else ''}）")
        rules = settle_flow(fl)                                      # 決まりで決まる所（判断の名前・分岐のラベル・終了条件・確認事項の重要度）。採否表があっても当てる（番号は変わらない）
        for x in rules:
            print(f"{f.name}: {x}")
        if ren or marks or tags or boxes or heads or rules or lanes:                # 振り直しが止まっても、ほかの当てはめは保存する（renumber は番号を書き換える前に止まる）
            changed += 1
            if not a.dry_run:
                save(f, fl)
                moved += times.rename(fl.get("flow_id") or f.name.split(".")[0], ren or {})
    if moved:
        times.save()
        print(f"times.json: 人が入れた処理時間 {moved} 個の番号も、振り直した")
    print(f"番号・推定の印・札・見出し・決まりで決まる所が変わ{'る' if a.dry_run else 'った'}フロー {changed} 本")


if __name__ == "__main__":
    main()
