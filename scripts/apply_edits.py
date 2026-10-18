#!/usr/bin/env python3
"""apply_edits.py — 下書き（draft_flow.py）に、直しの一覧（flows/<ID>.edits.json）を当てて、flows/<ID>.flow.json を作る。

  python3 scripts/apply_edits.py -w <作業フォルダ> --flow F03
  python3 scripts/apply_edits.py -w <作業フォルダ> --flow F03 --no-edits     # 直しの一覧なしで、下書きのままのフローを作る（確かめ用）

LLM が書くのは、直しの一覧だけ（形は references/schemas.md の「直しの一覧」）。ここが、下書きに直しを当て、線・帳票・開始と終了・
決まった形の確認事項・番号・レーンの並び（決まり 4.2。直しの一覧の lane_order で指定できる）を機械で付けて、flow.json にする。flow.json から先（点検・図・業務マトリクス一覧・原典照合）は、いままでと同じ。
  直しの種類（ops[].op）   merge まとめる／unmerge まとめた候補を外す／exclude フローに入れない（理由）／lane レーンを決める／label ラベルを短くする／
                           detail 補足・参考を書く／split 担当の違う作業に分ける／decide 手順の条件を判断にする／branch 書かれていない側の行き先／
                           add 原典に無い作業・判断を足す（推定）／move 並べ替える／connect・disconnect 線を足す・外す／link ほかの業務へ続く／drop_file 帳票を付けない／
                           time 処理時間の動作の種類・量
  機械が作る確認事項       担当の書かれていない作業（1件にまとめる）／順の書かれていない箇条書き／補足の中の作業の位置／資料をまたぐ順／
                           分岐の書かれていない側／並べ替え・足した線（推定）／原典に無い判断／開始条件・終了条件が空
直しの一覧に誤りがあれば（無いキー・決まっていない種類・根拠の無い分け方など）、何も書かずに、誤りを全部並べて終わる（終了コード 1）。
"""
import argparse
import copy
import pathlib
import re
import sys

from common import (ACTIONS, EFFECTS, ELSE, ENDS_ASKED, LABEL_LONG, SEV_LABEL, VOLUMES, CardItems, ask_form, flow_heads, infer_nodes, load, load_page, norm, opposite, order_lanes, page_json, page_units, same_text, site,
                    save, settle_flow, severity_of, span_text, system_names, tag_systems, try_load, work_dir)
from draft_flow import PAGE_TYPE_JA, Draft
from renumber import PREFIX, apply_ids, renumber
from validate import blocking, validate

TIME_FIELDS = ("action", "volume", "volume_reason")              # 処理時間の仮の値を引く鍵（人の作業の箱だけ）
OPS = ("merge", "unmerge", "exclude", "lane", "label", "detail", "time", "split", "decide", "branch", "add", "move", "connect", "disconnect", "link", "drop_file")
RETURN = re.compile(r"差し?戻す|差戻す|返却する|突き返す")              # 差し戻す作業（決まり 3.7: 分岐の先にあれば、そのあとは推定の終了＋確認事項）
RETURNED = re.compile(r"差し?戻(し|され)|差戻(し|され)")               # 差し戻しを受ける側の作業（足さない）
LANE_KIND = {"actor": "role", "system": "system"}                    # 用語の種類 → レーンの種類（lane_kinds で直せる）


def short_label(label, limit=LABEL_LONG):
    """LABEL_LONG（30 字）を超え、かっこ書きがあるラベルは、かっこの前までにする。(ラベル, 補足に回す文字)"""
    head, sep, rest = label.partition("（")
    if len(label) <= limit or not sep or len(head.strip()) < 8:
        return label, ""
    return head.strip(), rest.rstrip("）。 ").strip()


def branch_label(text):
    """分岐の線のラベル: 「もし」「〜の場合は」「〜とき」を外す（「不備がある場合」→「不備がある」）"""
    return re.sub(r"(の?場合|とき|時)(に|は|には)?$", "", re.sub(r"^もし、?", "", str(text or "").strip())).strip() or str(text or "").strip()


class Builder:
    def __init__(self, draft, edits):
        self.d, self.e, self.err = draft, edits or {}, []
        self.nodes = copy.deepcopy(draft.nodes)
        self.seq = list(draft.seq)
        self.merged = {a: b for a, (b, _why) in draft.merged.items()}
        self.excluded, self.moved, self.links, self.else_of, self.orig = {}, {}, [], {}, {}
        self.clash = []                                              # まとめた項目どうしで担当が違うもの（まとめ先のキー, まとめ先の担当, まとめた側の担当）

    # ---- 直しを当てる ----

    def bad(self, i, op, msg):
        self.err.append(site(f"ops[{i}]（{op.get('op')}）: {msg}", 2))

    def res(self, key):
        """キー → いまのノードのキー（まとめた項目は、まとめ先）。無ければ None"""
        seen = set()
        while key in self.merged and key not in seen:
            seen.add(key)
            key = self.merged[key]
        return key if key in self.nodes and key not in self.excluded else None

    def need(self, i, op, name):
        k = self.res(op.get(name))
        if k is None:
            self.bad(i, op, f"{name}「{op.get(name)}」という項目が、下書きに無い（対象外にした項目も指せない）")
        return k

    def lane_name(self, i, op, text):
        nm = self.d.terms.name(text, ("actor", "system"))
        if not nm:
            self.bad(i, op, f"レーン「{text}」が、用語の正本（glossary.json）に無い")
        return nm

    def apply(self):
        ops = [x for x in self.e.get("ops") or [] if isinstance(x, dict)]
        for i, op in enumerate(ops):
            if op.get("op") not in OPS:
                self.bad(i, op, f"決まっていない種類。使えるのは {'・'.join(OPS)}")
        for phase in (("unmerge",), ("merge", "exclude"), ("lane", "label", "detail", "time", "drop_file", "split", "decide", "branch"), ("add", "move", "link")):
            for i, op in enumerate(ops):
                if op.get("op") in phase:
                    getattr(self, "op_" + op["op"])(i, op)
            if phase == ("merge", "exclude"):
                self.fold()
        return ops

    def op_unmerge(self, i, op):
        if op.get("item") not in self.merged:
            return self.bad(i, op, f"item「{op.get('item')}」は、まとめた候補に無い")
        del self.merged[op["item"]]

    def op_merge(self, i, op):
        into = self.need(i, op, "into")
        for k in op.get("items") or []:
            if k not in self.nodes or self.res(k) is None:
                self.bad(i, op, f"items の「{k}」という項目が、下書きに無い")
            elif into and k != into:
                if self.nodes[k]["kind"] != self.nodes[into]["kind"]:
                    self.bad(i, op, f"「{k}」と「{into}」は種類が違う（手順は手順に、分岐は分岐にまとめる）")
                self.merged[k] = into
        vals = op.get("values") if isinstance(op.get("values"), dict) else {}
        if vals and into:                                            # 資料どうしで値だけが違う（決まり 3.12）: 補足を決まった形で書く「まとめ先の値。〈資料〉では、〈値〉」
            if into not in vals or any(k != into and k not in (op.get("items") or []) for k in vals):
                return self.bad(i, op, "values には、まとめ先（into）と、まとめる項目（items）の値を、項目のキーで書く")
            title = self.d.lab.title
            self.nodes[into]["detail"] = "。".join([str(vals[into]).strip()] + [f"{title(self.nodes[k]['doc'])}では、{str(v).strip()}" for k, v in vals.items() if k != into])

    def op_exclude(self, i, op):
        k = op.get("item")
        if k not in self.nodes or k in self.merged:
            return self.bad(i, op, f"item「{k}」が下書きに無いか、ほかの項目にまとめてある（まとめた項目は、まとめ先を指すか、先に unmerge する）")
        same, other = op.get("same_as"), op.get("flow")
        if same is not None and (same not in self.nodes or same == k):
            return self.bad(i, op, f"same_as「{same}」が、下書きに無い（同じ作業として入れる、ほかの項目のキーを書く）")
        if other is not None and other not in self.d.cats:
            return self.bad(i, op, f"flow「{other}」が、業務分類に無い")
        fixed = (f"「{self.nodes[same]['label'][:40]}」と同じ作業（そちらをフローに入れた）" if same is not None else
                 f"「{self.d.cats[other].get('name')}」（{other}）のフローで描く" if other is not None else "")   # 決まった形の理由
        why = "。".join(x for x in (fixed, str(op.get("reason", "")).strip()) if x)
        if not why:
            return self.bad(i, op, "reason（入れない理由）が空（同じ作業なら same_as、ほかのフローで描くなら flow を書く）")
        if any(b == k for b in self.merged.values()):
            return self.bad(i, op, f"「{k}」には、ほかの項目がまとめてある。先に unmerge する")
        self.excluded[k] = why

    def fold(self):
        """まとめた項目を、まとめ先に畳む（出典・帳票・担当を引き継ぐ）"""
        for a in [k for k in self.seq if k in self.merged]:
            b = self.res(a)
            if b is None:                                            # まとめ先が対象外になった
                self.err.append(site(f"「{a}」のまとめ先が無い（対象外にした）。unmerge するか、まとめ先を変える"))
                continue
            A, B = self.nodes[a], self.nodes[b]
            B["sources"] += A["sources"]
            B.setdefault("via", {}).update({x: a for x in A["inputs"] + A["outputs"] if x not in B["inputs"] + B["outputs"]})   # その帳票が書かれている項目
            B["inputs"] += [x for x in A["inputs"] if x not in B["inputs"]]
            B["outputs"] += [x for x in A["outputs"] if x not in B["outputs"]]
            if not B["actor"] and A["actor"]:
                B["actor"], B["lane"] = A["actor"], A["lane"]
            elif A["lane"] and B["lane"] and A["lane"] != B["lane"] and (b, B["lane"], A["lane"]) not in self.clash:
                self.clash.append((b, B["lane"], A["lane"]))             # 担当が違う作業をまとめた: 黙って捨てず、確認事項にする（決まり 3.12）
        self.seq = [k for k in self.seq if k not in self.merged and k not in self.excluded]
        self.orig = dict(zip(self.seq, self.seq[1:]))                # 下書きの並びで、すぐあとにあった項目（並べ替えたあとも、もとの続きは推定にしない）

    def op_lane(self, i, op):
        k, nm = self.need(i, op, "item"), self.lane_name(i, op, op.get("lane"))
        if k and nm:
            self.nodes[k]["lane"], self.nodes[k]["lane_set"] = nm, True

    def op_label(self, i, op):
        k = self.need(i, op, "item")
        if k and not str(op.get("label", "")).strip():
            return self.bad(i, op, "label が空")
        if k:
            self.nodes[k]["label"], self.nodes[k]["relabeled"] = str(op["label"]).strip(), True
            names = op.get("branches")
            if names is not None:                                    # 分岐の線のラベルも短くする（カードの branches の順に）
                if self.nodes[k]["kind"] != "decision" or not isinstance(names, list) or len(names) != len(self.nodes[k]["branches"]):
                    return self.bad(i, op, f"branches は、分岐の行き先（{len(self.nodes[k]['branches'])} 本）と同じ数だけ、その順に書く")
                for b, nm in zip(self.nodes[k]["branches"], names):
                    b["label"] = str(nm).strip()

    def op_detail(self, i, op):
        if "detail" not in op and "reference" not in op:
            self.bad(i, op, "detail か reference を書く")
            return
        k = self.need(i, op, "item")
        if k:
            for f in ("detail", "reference"):
                if f in op:
                    self.nodes[k][f] = str(op[f]).strip()

    def timing(self, i, op, src):
        """動作の種類・量（処理時間の仮の値を引く鍵）。決まった語だけ（動作の種類は、2 つの動きのある箱だけ、違う語を 2 つまで）。量を変えるなら、理由が要る"""
        act, vol, why = src.get("action"), src.get("volume"), str(op.get("reason") or "").strip()
        acts = act if isinstance(act, list) else [act]
        if act is not None and (not 1 <= len(acts) <= 2 or any(a not in ACTIONS for a in acts) or len(set(map(str, acts))) < len(acts)):
            return self.bad(i, op, f"action「{'・'.join(map(str, acts))}」は、決まった語でないか、2 語を超える・同じ語が 2 回（{'・'.join(ACTIONS)}。2 つの動きのある箱だけ、違う語を 2 つ）")
        act = acts[0] if len(acts) == 1 else act
        if vol is not None and vol not in VOLUMES:
            return self.bad(i, op, f"volume「{vol}」は、決まった語でない（{'・'.join(VOLUMES)}）")
        if vol is not None and not why:
            return self.bad(i, op, "volume（量）を書くときは、reason（機械の決め方と変える理由）も書く")
        return {k: v for k, v in (("action", act), ("volume", vol), ("volume_reason", why if vol else None)) if v}

    def op_time(self, i, op):
        k = self.need(i, op, "item")
        if k and self.nodes[k]["kind"] != "step":
            return self.bad(i, op, f"「{k}」は手順でない（動作の種類を付けるのは、人の作業の箱だけ。判断は、機械が決める）")
        if k and "action" not in op and "volume" not in op:
            return self.bad(i, op, "action（動作の種類）か volume（量）を書く")
        if k:
            self.nodes[k].update(self.timing(i, op, op) or {})

    def op_drop_file(self, i, op):
        k = self.need(i, op, "item")
        if k:
            n, name = self.nodes[k], op.get("name")
            if name not in n["inputs"] + n["outputs"]:
                return self.bad(i, op, f"「{k}」に、帳票「{name}」は付いていない（付いているのは {'・'.join(n['inputs'] + n['outputs']) or 'なし'}）")
            n["inputs"], n["outputs"] = [x for x in n["inputs"] if x != name], [x for x in n["outputs"] if x != name]

    def op_split(self, i, op):
        k, parts = self.need(i, op, "item"), [p for p in op.get("parts") or [] if isinstance(p, dict)]
        if not k:
            return
        if self.nodes[k]["kind"] != "step" or len(parts) < 2:
            return self.bad(i, op, "分けられるのは手順だけ。parts は2つ以上")
        out = []
        for p in parts:
            nm = self.lane_name(i, op, p.get("lane"))
            if not str(p.get("label", "")).strip():
                self.bad(i, op, "parts の label が空")
            out.append(dict({"label": str(p.get("label", "")).strip(), "lane": nm, "inferred": bool(p.get("inferred")),
                             "detail": str(p.get("detail") or "").strip()},                  # 部分ごとの補足（書かなければ、手順の補足は最初の部分に付く）
                            **(self.timing(i, dict(op, reason=op.get("reason") or "部分ごとに指定"), p) or {})))
        if any(a["lane"] == b["lane"] for a, b in zip(out, out[1:])):
            return self.bad(i, op, "続く2つが同じレーン。分けるのは、文の途中で担当が変わるときだけ（同じ担当の作業は、1つのノードのままにする）")
        self.nodes[k]["parts"] = out

    def op_decide(self, i, op):
        k = self.need(i, op, "item")
        if not k:
            return
        if self.nodes[k]["kind"] != "step" or not self.nodes[k]["condition"]:
            return self.bad(i, op, f"「{k}」は、条件（condition）の書かれた手順でない。判断にできるのは、カードに条件が書かれた手順だけ")
        if not all(str(op.get(x, "")).strip() for x in ("question", "yes")):
            return self.bad(i, op, "question（何で分かれるか）・yes（その手順を行う側のラベル）を書く")
        self.nodes[k]["gate"] = {x: str(op.get(x) or "").strip() for x in ("question", "yes", "no")}
        self.nodes[k]["gate"]["no"] = self.nodes[k]["gate"]["no"] or opposite(self.nodes[k]["gate"]["yes"]) or ELSE   # 行わない側は、反対の語（決まり 3.4）
        self.nodes[k]["gate"]["no_to"] = op.get("no_to", "next")
        self.nodes[k]["gate"]["lane"] = self.lane_name(i, op, op.get("lane")) if op.get("lane") else ""   # 判断する人（書かなければ、その先の手順のレーン）

    def op_branch(self, i, op):
        k = self.need(i, op, "item")
        if k and self.nodes[k]["kind"] != "decision":
            return self.bad(i, op, f"「{k}」は分岐でない")
        if k:
            self.else_of[k] = (str(op.get("label") or "").strip(), op.get("to", "next"))   # ラベルを書かなければ、書かれた側の反対の語

    def op_add(self, i, op):
        """原典に手順として書かれていない作業・判断を足す（推定のノードになり、確認事項が付く）"""
        src, kind, label = (op.get("source") if isinstance(op.get("source"), dict) else {}), op.get("type"), str(op.get("label") or "").strip()
        key = str(op.get("key") or f"x{sum(1 for n in self.nodes.values() if n.get('added')) + 1}")
        if kind not in ("process", "decision") or not label:
            return self.bad(i, op, "type（process か decision）と label を書く")
        if key in self.nodes or "." in key:
            return self.bad(i, op, f"key「{key}」が、ほかの項目と重なっているか、「.」を含む（x1・x2 のように付ける）")
        if not (src.get("doc_id") and isinstance(src.get("page"), int) and (src.get("lines") or str(src.get("quote", "")).strip())):
            return self.bad(i, op, "source（そう読める場所: doc_id・page と、lines か quote）を書く")
        ref = self.need(i, op, "before" if kind == "decision" else "after")
        lane = self.lane_name(i, op, op.get("lane")) if op.get("lane") else ""
        if kind == "process" and not lane:
            return self.bad(i, op, "足す作業の lane（レーン）を書く")
        if kind == "decision" and not str(op.get("yes", "")).strip():
            return self.bad(i, op, "足す判断の yes（before の項目へ進む側のラベル）を書く（no を書かなければ、yes の反対の語）")
        if kind == "process" and RETURNED.search(label):
            return self.bad(i, op, "差し戻しを受ける側の箱は足さない（決まり 3.7: 差し戻したあとは、機械が推定の終了と確認事項にする。"
                                   "原典に、差し戻したあとの手順が書かれていれば、connect でその手順へつなぐ）")
        if not ref:
            return
        self.nodes[key] = {"key": key, "doc": src["doc_id"], "page": src["page"], "item": None, "kind": "step" if kind == "process" else "decision", "overview": False,
                           "label": label, "relabeled": True, "actor": lane, "lane": lane, "lane_set": bool(lane), "system": "", "condition": "", "inputs": [], "outputs": [],
                           "origin": "body", "applies_to": None, "num": None, "title": "", "line": None, "mark": "", "group": 0, "branches": [], "sources": [], "files": [],
                           "added": src, "fork": {"yes": str(op.get("yes", "")).strip(), "no": str(op.get("no") or "").strip() or opposite(op.get("yes")) or ELSE} if kind == "decision" else None}
        if kind == "process":
            self.nodes[key].update(self.timing(i, op, op) or {})
        if kind == "decision":
            self.nodes[key]["fork"]["no_to"] = op.get("no_to", "next")
        self.seq.insert(self.seq.index(ref) + (0 if kind == "decision" else 1), key)

    def op_move(self, i, op):
        k = self.need(i, op, "item")
        after = op.get("after")
        a = after if after == "start" else self.need(i, op, "after")
        if k and a and k != a:
            self.seq.remove(k)
            self.seq.insert(0 if a == "start" else self.seq.index(a) + 1, k)
            self.moved[k] = op.get("inferred", True) is not False

    def op_link(self, i, op):
        to = op.get("flow")
        if not (self.d.cats.get(to) or {}).get("make_flow") or to == self.d.cid:
            return self.bad(i, op, f"flow「{to}」という、フローを描く業務分類が無い")
        src = op.get("source") if isinstance(op.get("source"), dict) else {}
        if not (src.get("doc_id") and isinstance(src.get("page"), int) and (src.get("lines") or str(src.get("quote", "")).strip())):
            return self.bad(i, op, "source（続くことが書かれた場所: doc_id・page と、lines か quote）を書く")
        after = self.need(i, op, "after") if op.get("after") else None
        lane = self.lane_name(i, op, op.get("lane")) if op.get("lane") else ""   # 別のフローへ進む箱の担当（書かなければ、前の箱と同じ）
        self.links.append({"flow": to, "after": after, "source": src, "label": str(op.get("label") or "").strip(), "lane": lane})

    # ---- フローに組み立てる ----

    def card_item(self, doc, page, item):
        c = self.d.cards.get((doc, page)) or {}
        it = next((x for place in ("steps", "decisions") for x in c.get(place) or [] if x.get("id") == item), None)
        if it is None:                                               # 図の箱（図だけのページ）: 引用は箱の字
            g = next((x for x in (c.get("diagram") or {}).get("nodes") or [] if x.get("id") == item), None)
            return {"quote": str(g.get("label", "")), "from_image": False} if g else {}
        return it

    def source(self, doc, page, item):
        it = self.card_item(doc, page, item)
        return {"doc_id": doc, "page": page, "item": item, "quote": str(it.get("quote", "")), "from_image": bool(it.get("from_image"))}

    def plain(self, src):
        """カードの項目を引かない出典（足した作業・判断、リンク）。lines があれば、その行の文字を引用にする"""
        out = {"doc_id": src["doc_id"], "page": src["page"], "quote": str(src.get("quote", "")), "from_image": bool(src.get("from_image"))}
        if src.get("lines"):
            text = span_text(page_units(load_page(page_json(self.d.work, src["doc_id"], src["page"]))), src["lines"])
            if text is None:
                self.err.append(site(f"出典の行の番号 {src['lines']} が、{src['doc_id']} p{src['page']:04d} の範囲の外"))
            out.update(lines=src["lines"], quote=text or "")
        return out

    def expand(self):
        """並びを、フローのノード（仮の id）にする。判断にした手順は「判断＋手順」、分けた手順は部分ごと"""
        out = []
        for k in self.seq:
            n = self.nodes[k]
            base = {"key": k, "doc": n["doc"], "page": n["page"], "origin": n["origin"], "mark": n["mark"], "group": n["group"], "actor": n["actor"],
                    "sources": [self.source(*s) for s in n["sources"]] + ([self.plain(n["added"])] if n.get("added") else []), "fork": n.get("fork")}
            if n.get("gate"):
                out.append(dict(base, id=k + "?", type="decision", label=n["gate"]["question"], lane=n["gate"].get("lane", ""), gate=n["gate"], follows=k))
            label, aside = short_label(n["label"]) if not n.get("relabeled") else (n["label"], "")
            detail = n.get("detail", "・".join(x for x in (n["condition"] if not n.get("gate") and not n.get("target") else "", aside) if x))   # 使うシステムは、札（systems）に出る
            if n.get("parts"):
                for j, p in enumerate(n["parts"], 1):
                    out.append(dict(base, id=f"{k}#{j}", type="process", label=p["label"], lane=p["lane"], lane_set=True, part=j, soft=p["inferred"],
                                    detail=p.get("detail") or (detail if j == 1 and not any(q.get("detail") for q in n["parts"]) else ""),
                                    reference=n.get("reference", "") if j == 1 else "", files=j == 1,
                                    **{f: p.get(f) or n.get(f) for f in TIME_FIELDS}))
            else:
                out.append(dict(base, id=k, type="decision" if n["kind"] == "decision" else "process", label=label, lane=n["lane"],
                                lane_set=n.get("lane_set", False), detail=detail if n["kind"] == "step" else "", reference=n.get("reference", ""),
                                files=n["kind"] == "step", branches=n["branches"], **{f: n.get(f) for f in TIME_FIELDS if n["kind"] == "step"}))
        return out

    def mark_targets(self):
        """分岐の行き先に当たる手順（条件は分岐が表すので、補足に重ねて書かない）"""
        for i, k in enumerate(self.seq):
            n = self.nodes[k]
            if n["kind"] == "decision":
                for t in self.targets(i):
                    if t is not None:
                        self.nodes[self.seq[t]]["target"] = True

    def targets(self, i):
        """並びの i 番目の分岐について、行き先ごとに、当たる手順の位置（無ければ None）。行き先が1つで当たる手順が無ければ、すぐあとの手順"""
        n, later, out = self.nodes[self.seq[i]], [(j, self.nodes[k]) for j, k in enumerate(self.seq) if j > i and self.nodes[k]["kind"] == "step"], []
        for b in n["branches"]:
            to = str(b.get("to") or "")
            hit = next((j for j, x in later if to and x["page"] == n["page"] and (same_text(to, x["label"]) or any(
                same_text(to, str(self.card_item(*s).get("action", ""))) for s in x["sources"])) and j not in out), None)
            out.append(hit)
        if n["branches"] and out[0] is None and later and later[0][0] == i + 1:
            out[0] = i + 1
        return out

    def build(self):
        e, d = self.e, self.d
        self.mark_targets()
        order = self.expand()
        if not order:
            self.err.append(site("フローに入れる手順が1つも無い"))
            return None
        for n in order:                                              # 判断の名前は「〜か」（「？」で終わる名前を直す）
            if n["type"] == "decision":
                n["label"] = ask_form(n["label"])
        by = {n["id"]: n for n in order}
        first = {}                                                   # キー → そのキーの最初のノードの id（判断にした手順は判断、分けた手順は1つ目）
        for n in order:
            first.setdefault(n["key"], n["id"])
        last = {n["key"]: n["id"] for n in order}
        # レーン
        default = d.terms.name(e.get("default_lane", ""), ("actor", "system")) if e.get("default_lane") else ""
        if e.get("default_lane") and not default:
            self.err.append(site(f"default_lane「{e['default_lane']}」が、用語の正本（glossary.json）に無い"))
        used = [n["lane"] for n in order if n["lane"]]
        fallback = default or (max(used, key=lambda x: (used.count(x), -used.index(x))) if used else (d.actors[0] if d.actors else ""))
        if not fallback:
            self.err.append(site("レーンが1つも決まらない（カードに担当が無い）。default_lane を書く"))
        for i, n in enumerate(order):
            if n["type"] == "decision" and not n["lane"]:            # 判断は、その先の手順のレーン
                nxt = next((x for x in order[i + 1:] if x["type"] != "decision"), None)
                n["lane"] = (nxt["lane"] if nxt and nxt["lane"] else "") or fallback
            n["lane"] = n["lane"] or fallback
            if n["type"] == "process" and d.terms.kind(n["lane"]) == "system":
                n["type"] = "system"
        # 線
        edges, soft_why = [], {}

        def add(a, b, label="", why=None):
            edges.append(dict({"from": a, "to": b}, **({"label": label} if label else {}), **({"inferred": True} if why else {})))
            if why:
                soft_why[(a, b)] = why

        def step_after(i):                                           # i 番目のあとのノード（無ければ終了）
            return order[i + 1]["id"] if i + 1 < len(order) else "end"

        def soft(a, b):
            """a のすぐあとに b を置いた線が、推定か（理由を返す）"""
            if b.get("part", 1) > 1:
                return "part" if b.get("soft") else None
            if self.moved.get(b["key"]) or (self.moved.get(a["key"]) and self.orig.get(a["key"]) != b["key"]):
                return "move"
            if "note" in (a["origin"], b["origin"]) and a["key"] != b["key"]:
                return "note"
            if a["doc"] != b["doc"]:
                return "doc"
            if a["mark"] == b["mark"] == "bullet" and a["page"] == b["page"] and a["group"] != b["group"]:
                return "list"
            return None
        pos = {k: i for i, k in enumerate(self.seq)}
        jump = {}                                                    # 分岐の行き先どうし（どちらか一方を行う）は、最後の行き先のあとで合流する
        add("start", order[0]["id"])
        for i, n in enumerate(order):
            if n.get("gate"):
                g = n["gate"]
                add(n["id"], order[i + 1]["id"], g["yes"])
                to = self.exit(g["no_to"], last[n["key"]], order, first)
                add(n["id"], to, g["no"], "else")
            elif n.get("fork"):                                      # 足した判断: 進む側は次の項目、進まない側は no_to（どちらも線は推定にしない。ノードが推定）
                g = n["fork"]
                if i + 1 >= len(order):
                    self.err.append(site(f"足した判断「{n['label'][:16]}」のあとに項目が無い"))
                    continue
                add(n["id"], order[i + 1]["id"], g["yes"])
                add(n["id"], self.exit(g["no_to"], last[order[i + 1]["key"]], order, first), g["no"])
            elif n["type"] == "decision":
                ts = self.targets(pos[n["key"]])
                hits = [(b, first[self.seq[t]]) for b, t in zip(n["branches"], ts) if t is not None]
                for b, to in hits:
                    add(n["id"], to, branch_label(b.get("label")))
                run = sorted(pos[by[to]["key"]] for _b, to in hits)
                after = last[self.seq[run[-1]]] if run else n["id"]
                nxt = step_after(next(j for j, x in enumerate(order) if x["id"] == after))
                if len(hits) >= 2 and run == list(range(run[0], run[-1] + 1)):
                    for t in run[:-1]:
                        jump[last[self.seq[t]]] = nxt
                if len(hits) < 2 or len(hits) < len(n["branches"]):   # 書かれていない側（行き先が1つ、または行き先の手順が見つからない）
                    label, to = self.else_of.get(n["key"], ("", "next"))
                    said = [branch_label(b.get("label")) for b in n["branches"] if str(b.get("label") or "").strip()]
                    label = label or (opposite(said[0]) if len(said) == 1 else None) or ELSE
                    add(n["id"], nxt if to == "next" else self.exit(to, after, order, first), label, "else")
            elif RETURN.search(n["label"]) and n["type"] == "process" and any(x["to"] == n["id"] and by[x["from"]]["type"] == "decision" for x in edges if x["from"] in by):
                add(n["id"], "end2", "", "dead")                     # 分岐の先の差し戻し: 次の作業へは進まない（受け手の箱は足さない。続きは確認事項で聞く）
            else:
                to = jump.get(n["id"], step_after(i))
                add(n["id"], to, "", soft(n, by[to]) if to in by else None)
        self.edge_ops(edges, soft_why, by, first, last)
        ends = {x["to"] for x in edges if x["to"] in ("end", "end2")}
        for n in order:                                              # 行き止まりは、推定の終了へ
            if not any(x["from"] == n["id"] for x in edges):
                add(n["id"], "end2", "", "dead")
                ends.add("end2")
        seen, todo = {"start"}, ["start"]
        while todo:
            u = todo.pop()
            for x in edges:
                if x["from"] == u and x["to"] not in seen:
                    seen.add(x["to"])
                    todo.append(x["to"])
        lost = [n for n in order if n["id"] not in seen]
        if lost:
            self.err.append(site("開始からたどり着けない項目がある（move・connect・disconnect で線が切れた）: " + "、".join(f"{n['key']}「{n['label'][:16]}」" for n in lost[:6])))
        for x in edges:
            if x["from"] in by and by[x["from"]]["type"] == "decision" and not x.get("label"):
                self.err.append(site(f"分岐 {by[x['from']]['key']}「{by[x['from']]['label'][:16]}」から出る線に、ラベルの無いものがある（カードの行き先にラベルが無い。label の branches で書く）"))
        # ほかの業務へ続く
        seqs = [x for x in (load(d.work / "structure" / "hierarchy.json", {}) or {}).get("sequence") or [] if isinstance(x, dict)]
        hid = {c.get("hierarchy_id"): c["id"] for c in d.cats.values() if c.get("make_flow")}
        auto = [{"flow": hid[x["to"]], "after": None, "label": "", "source": dict((x.get("sources") or [{}])[0], quote=str(x.get("evidence", "")))}
                for x in seqs if x.get("from") == d.cat.get("hierarchy_id") and x.get("to") in hid and (x.get("sources") or [{}])[0].get("doc_id")]
        subs = []
        for k, ln in enumerate(self.links + [x for x in auto if x["flow"] not in {y["flow"] for y in self.links}], 1):
            sid = f"link{k}"
            at = last[ln["after"]] if ln["after"] else next((x["from"] for x in edges if x["to"] == "end" and x["from"] != "start"), order[-1]["id"])
            src = ln["source"]
            subs.append({"id": sid, "type": "subprocess", "lane": ln.get("lane") or by[at]["lane"], "label": ln["label"] or f"{d.cats[ln['flow']].get('name')}に進む", "link_flow": ln["flow"],
                         "sources": [self.plain(src)]})
            for x in [x for x in edges if x["from"] == at and by[at]["type"] != "decision"]:      # at → 次 を、at → リンク → 次 にする
                x["from"] = sid
            edges.append({"from": at, "to": sid})
        # レーン・ノード・帳票
        names = []
        for n in order:
            if n["lane"] not in names:
                names.append(n["lane"])
        lane_of = {n["id"]: n["lane"] for n in order + subs}
        lane_of.update(start=order[0]["lane"], end=order[-1]["lane"], end2=order[-1]["lane"])
        kinds = e.get("lane_kinds") if isinstance(e.get("lane_kinds"), dict) else {}
        kind_of = {nm: kinds.get(nm) or LANE_KIND.get(d.terms.kind(nm), "role") for nm in names}
        systems = [x for x in names if kind_of[x] == "system"]
        want = [str(x).strip() for x in e.get("lane_order") or [] if str(x).strip()]
        lost = [x for x in want if x not in names]
        if lost:
            self.err.append(site(f"lane_order の「{'・'.join(lost)}」は、このフローのレーンに無い（いまのレーン: {'・'.join(names)}）"))
        twice = [x for x in dict.fromkeys(want) if want.count(x) > 1]
        if twice:
            self.err.append(site(f"lane_order に、同じレーンを2回以上書いている: {'・'.join(twice)}（並べたい順に1回ずつ書く）"))
        if [x for x in want if x in systems]:
            d.notes.append(f"lane_order に書いたシステムのレーン（{'・'.join(x for x in want if x in systems)}）は、書いた位置に置いた（決まり 4.2 では右端）")
        keep = want or ([] if order[0]["lane"] in systems else [order[0]["lane"]])
        names = order_lanes(names, [(lane_of.get(x["from"]), lane_of.get(x["to"])) for x in edges], keep=keep, right=[x for x in systems if x not in want])   # 決まり 4.2
        lid = {nm: f"L{i}" for i, nm in enumerate(names, 1)}
        nodes = [{"id": "start", "type": "start", "lane": lid[order[0]["lane"]], "label": "開始"}]
        for n in order:
            x = {"id": n["id"], "type": n["type"], "lane": lid[n["lane"]], "label": n["label"]}
            x.update({f: n[f] for f in ("detail", "reference") + (TIME_FIELDS if n["type"] == "process" else ()) if n.get(f)})
            x["sources"] = n["sources"]
            nodes.append(x)
            if n.get("files"):
                src = self.nodes[n["key"]]
                for name, way in [(f, "in") for f in src["inputs"]] + [(f, "out") for f in src["outputs"]]:
                    doc = d.terms.name(name, ("document",))
                    if not doc:
                        d.notes.append(f"{n['key']}: 帳票「{name}」が用語の正本に無いので、付けていない")
                        continue
                    owner = self.nodes[src.get("via", {}).get(name, n["key"])]
                    nodes.append({"id": f"{n['id']}/{way}/{doc}", "type": "file", "lane": lid[n["lane"]], "label": doc, "attach_to": n["id"], "direction": way,
                                  "sources": [self.source(owner["doc"], owner["page"], owner["item"])]})
        for s in subs:
            nodes.append(dict(s, lane=lid[s["lane"]]))
        end_lane = lid[order[-1]["lane"]]
        nodes.append({"id": "end", "type": "end", "lane": end_lane, "label": "終了"})
        if "end2" in ends:
            nodes.append({"id": "end2", "type": "end", "lane": end_lane, "label": "終了"})
        fl = {"flow_id": d.cid, **flow_heads(d.cat), "summary": str(e.get("summary") or d.cat.get("name") or "").strip(),
              "start_condition": str(e.get("start_condition") or "").strip(), "end_condition": str(e.get("end_condition") or "").strip(),
              "lanes": [{"id": lid[nm], "name": nm, "kind": kind_of[nm]} for nm in names],
              "nodes": nodes, "edges": edges, "unused_pages": self.unused(order), "excluded_items": [
                  {"doc_id": self.nodes[k]["doc"], "page": self.nodes[k]["page"], "item": self.nodes[k]["item"], "reason": r} for k, r in self.excluded.items()],
              "questions": []}
        infer_nodes(fl)
        tag_systems(fl, CardItems(d.work), system_names(d.work))
        d.notes.extend(settle_flow(fl, ask=False))                              # 決まりで決まる所（終了条件・分岐のラベル・判断の名前）。確認事項より先に（終了条件が入れば、問いにしない）
        fl["questions"] = self.questions(fl, order, by, first, soft_why)
        return fl

    def exit(self, to, after_id, order, first):
        """行き先の書き方（"next" すぐあと／"end" 終了／キー）→ ノードの id"""
        if to == "end":
            return "end2"
        if to == "next" or to is None:
            i = next(j for j, x in enumerate(order) if x["id"] == after_id)
            return order[i + 1]["id"] if i + 1 < len(order) else "end"
        k = self.res(to)
        if k is None or k not in first:
            self.err.append(site(f"行き先「{to}」という項目が無い（\"next\"・\"end\"・項目のキーのどれか）"))
            return "end"
        return first[k]

    def edge_ops(self, edges, soft_why, by, first, last):
        for i, op in enumerate(x for x in self.e.get("ops") or [] if isinstance(x, dict)):
            if op.get("op") not in ("connect", "disconnect"):
                continue
            a = self.need(i, op, "from")
            b = "end" if op.get("to") == "end" else self.need(i, op, "to")
            if not a or not b:
                continue
            src, dst = last[a], (b if b == "end" else first[b])
            if op["op"] == "disconnect":
                hit = [x for x in edges if x["from"] == src and x["to"] == dst]
                if not hit:
                    self.bad(i, op, f"線 {a} → {b} は、下書きに無い")
                for x in hit:
                    edges.remove(x)
                continue
            label = str(op.get("label") or "").strip()
            if by[src]["type"] == "decision" and not label:
                self.bad(i, op, "分岐から出る線には label が要る")
            if by[src]["type"] != "decision":                        # 分岐でないノードの線は1本。先を付け替える
                for x in [x for x in edges if x["from"] == src]:
                    edges.remove(x)
            x = dict({"from": src, "to": dst}, **({"label": label} if label else {}))
            if op.get("inferred", True) is not False:
                x["inferred"] = True
                soft_why[(src, dst)] = "connect"
            elif not isinstance(op.get("evidence"), dict):
                self.bad(i, op, "推定でない線（inferred: false）には、順が書かれた場所（evidence: doc_id・page・lines）を書く")
            edges.append(x)

    def unused(self, order):
        cited = {(s["doc_id"], s["page"]) for n in order for s in n["sources"]}
        given = {(x.get("doc_id"), x.get("page")): str(x.get("reason", "")).strip() for x in self.e.get("unused_pages") or [] if isinstance(x, dict)}
        out = []
        for doc, page in self.d.pages:
            if (doc, page) in cited:
                continue
            c = self.d.cards.get((doc, page)) or {}
            has = bool((c.get("steps") or []) + (c.get("decisions") or []))
            out.append({"doc_id": doc, "page": page, "reason": given.get((doc, page)) or (
                "このページの手順・分岐は、どれもフローに入れなかった（理由は excluded_items）" if has
                else f"手順・分岐の記載が無いページ（{PAGE_TYPE_JA.get(c.get('page_type'), '種別なし')}）")})
        return out

    # ---- 確認事項 ----

    def questions(self, fl, order, by, first, soft_why):
        """機械が作る確認事項（答えで、図のどこが変わるか＝M_EFFECT）と、LLM が書いた確認事項"""
        d, e, out = self.d, self.e, []
        lane = {l["id"]: l["name"] for l in fl["lanes"]}
        node = {n["id"]: n for n in fl["nodes"]}
        twice = {n["label"] for n in fl["nodes"] if sum(1 for x in fl["nodes"] if x["label"] == n["label"] and x["type"] == n["type"]) > 1 and n["type"] != "end"}
        name = lambda i: f"「{node[i]['label'][:40]}」" + (f"（{lane[node[i]['lane']]}）" if node[i]["label"] in twice else "")   # 同じラベルが2つあれば、担当を添える
        pages = lambda ids: [dict(t) for t in dict.fromkeys(tuple(sorted({"doc_id": s["doc_id"], "page": s["page"]}.items()))
                                                           for i in ids for s in node[i].get("sources") or [])]

        def q(qid, cat, text, assume, ids=(), edge=None):
            x = {"id": qid, "severity": "", "category": cat, "effect": M_EFFECT[re.sub(r"-\d+$", "", qid)]}
            if edge:
                x["edge"] = list(edge)
            elif len(ids) == 1:
                x["node"] = ids[0]
            elif ids:
                x["nodes"] = list(ids)
            ref = list(edge) if edge else list(ids)
            x["severity"] = severity_of(x)                           # 重要度は決まり 5.4 で（答えで変わる所と、指す箱の数）
            x.update(question=text, assumption=assume, sources=pages([i for i in ref if i in node]) or [{"doc_id": d.pages[0][0], "page": d.pages[0][1]}],
                     ask=lane.get(node[ref[0]]["lane"], "") if ref and ref[0] in node else fl["lanes"][0]["name"])
            out.append(x)
        # 担当の書かれていない作業（1件にまとめる）
        idle = [n["id"] for n in order if n["type"] in ("process", "decision") and not n["actor"]
                and not any(norm(lane[node[n["id"]]["lane"]]) in norm(s["quote"]) for s in n["sources"])
                and any(self.card_item(s["doc_id"], s["page"], s.get("item")).get("actor") is not None for s in n["sources"])]
        if idle:
            where = "・".join(dict.fromkeys(lane[node[i]["lane"]] for i in idle))
            q("m-lane", "記載なし", f"次の作業は、原典に担当が書かれていません。{where}が行っていますか、別の担当ですか: " + "、".join(name(i) for i in idle),
              f"担当の書かれていない作業を、{where}のレーンに置いた", idle)
        # まとめた作業で、資料によって担当が違うもの（まとめ先の担当に置いた）
        for j, (key, mine, other) in enumerate([c for c in self.clash if c[0] in first], 1):
            q(f"m-actor-{j}", "資料間の矛盾", f"{name(first[key])}は、資料によって担当が違います（{mine}・{other}）。どちらが行いますか。",
              f"{lane[node[first[key]]['lane']]}のレーンに置いた", [first[key]])
        # 推定の線
        runs, k = [], 0
        soft = [(a, b, soft_why.get((a, b))) for x in fl["edges"] if x.get("inferred") for a, b in [(x["from"], x["to"])]]
        for a, b, why in soft:
            k += 1
            if why == "list":
                run = next((r for r in runs if r[-1] == a), None)
                if run is None:
                    runs.append([a, b])
                else:
                    run.append(b)
            elif why == "else":
                to = "終了" if b in ("end", "end2") else name(b)
                lab = next((x.get("label", "") for x in fl["edges"] if x["from"] == a and x["to"] == b), "")
                q(f"m-else-{k}", "記載なし", f"{name(a)}で、原典に書かれた場合に当たらないとき（「{lab}」）は、{to}に進みますか。",
                  f"書かれていない側は、{to}に進むと仮に置いた", edge=(a, b))
            elif why == "note":
                which = b if by.get(b, {}).get("origin") == "note" else a
                if which == b:                                       # 補足の作業が線の先: 「a のあと、b の次の箱の前」
                    nxt = next((x["to"] for x in fl["edges"] if x["from"] == b), None)
                    after, before = name(a), "終了" if nxt is None or nxt not in node or node[nxt]["type"] == "end" else name(nxt)
                else:                                                # 補足の作業が線の元: 「a の前の箱のあと、b の前」
                    prv = next((x["from"] for x in fl["edges"] if x["to"] == a), None)
                    after, before = ("開始" if prv is None or prv not in node or node[prv]["type"] == "start" else name(prv)), ("終了" if b not in node or node[b]["type"] == "end" else name(b))
                q(f"m-note-{k}", "記載なし", f"{name(which)}は、補足（注記・囲み）に書かれた作業です。{after}のあと、{before}の前に行いますか。",
                  "補足が掛かる手順のあと（分からなければ、書かれた順）に置いた", edge=(a, b))
            elif why == "doc":
                q(f"m-doc-{k}", "資料間の順序", f"{name(a)}のあとに、別の資料に書かれた{name(b)}を行いますか。", "資料の並びの順に置いた", edge=(a, b))
            elif why == "dead":
                q(f"m-dead-{k}", "記載なし", f"{name(a)}のあとは、この業務は終わりですか、続く作業がありますか。", "続きが書かれていないので、終了を仮に置いた", edge=(a, b))
            else:                                                    # 並べ替え・足した線・分けた作業の順
                q(f"m-seq-{k}", "記載なし", f"{name(a)}のあとは、{'終了' if b not in node or node[b]['type'] == 'end' else name(b)}に進みますか。", "原典に順が書かれていないので、この順に仮に置いた", edge=(a, b))
        for j, run in enumerate(runs, 1):
            q(f"m-list-{j}", "記載なし", "次の作業は、原典では箇条書きで、順が書かれていません。書かれた順に行いますか、順は決まっていませんか: " + "、".join(name(i) for i in run),
              "原典に書かれた順に並べた", run)
        for n in fl["nodes"]:                                        # 原典に無いノード（機械が推定の印を付けたもの）
            if n.get("inferred"):
                q(f"m-node-{len(out) + 1}", "記載なし", f"{name(n['id'])}は、原典に手順として書かれていません。実際に行っていますか。", "流れをつなぐために仮に置いた", [n["id"]])
        for end, (qid, key, text, assume) in ENDS_ASKED.items():
            if not fl[key]:
                q(qid, "記載なし", text, assume, [end])
        # LLM が書いた確認事項（対象はキーで書かれている）
        seen = {x["id"] for x in out}
        for i, x in enumerate(y for y in e.get("questions") or [] if isinstance(y, dict)):
            qid = str(x.get("id") or "")
            if not qid or qid in seen or qid.startswith("m-"):
                self.err.append(site(f"questions[{i}]: id「{qid}」が空か、重なっている（m- で始まる名前は、機械が作る確認事項のもの）"))
                continue
            seen.add(qid)
            if x.get("effect") not in EFFECTS:
                self.err.append(site(f"questions[{i}]（{qid}）: effect（答えで、図のどこが変わるか）が無い・決まった語でない（{'／'.join(EFFECTS)}）。"
                                "どれにも当たらない問い（値だけ・いまも同じかの確かめ・名前だけ）は、確認事項にしない"))
                continue
            y = {"id": qid, "severity": "", "category": x.get("category", "記載なし"), "effect": x["effect"]}
            ids = [self.res(k) for k in x.get("items") or []]
            edge = [self.res(k) for k in x.get("edge") or []]
            if None in ids or None in edge or (edge and len(edge) != 2):
                self.err.append(site(f"questions[{i}]（{qid}）: items・edge が指す項目が、下書きに無い"))
                continue
            parts = {}                                               # 分けた手順を指す問いは、分けた部分の全部を指す（担当を仮に決めたのは、どの部分も同じ）
            for n in order:
                if n.get("part"):
                    parts.setdefault(n["key"], []).append(n["id"])
            ids = list(dict.fromkeys(j for k in ids for j in (parts.get(k) or [first[k]])))
            if edge:
                y["edge"] = [first[edge[0]], first[edge[1]]]
            elif len(ids) == 1:
                y["node"] = ids[0]
            elif ids:
                y["nodes"] = ids
            y["severity"] = severity_of(y)
            if x.get("severity") not in (None, y["severity"]):
                d.notes.append(f"確認事項 {qid}: 重要度は、決まり 5.4 で{SEV_LABEL[y['severity']]}にした（直しの一覧の「{SEV_LABEL.get(x['severity'], x['severity'])}」は使わない）")
            y.update(question=str(x.get("question", "")).strip(), assumption=str(x.get("assumption", "")).strip(),
                     sources=[s for s in x.get("sources") or [] if isinstance(s, dict)] or pages(y.get("edge") or ids) or [{"doc_id": d.pages[0][0], "page": d.pages[0][1]}],
                     ask=str(x.get("ask") or "").strip() or (lane.get(node[ids[0]]["lane"], "") if ids else fl["lanes"][0]["name"]))
            out.append(y)
        return out


# 機械が作る確認事項の、答えで変わる所（id の頭 → effect）
M_EFFECT = {"m-lane": "担当", "m-actor": "担当", "m-else": "分岐", "m-note": "順序", "m-doc": "順序", "m-seq": "順序", "m-list": "順序", "m-dead": "箱", "m-node": "箱",
            "m-start": "フローの範囲", "m-end": "フローの範囲"}


def keep_numbers(work, fl):
    """原典照合のあと（採否表 review/<ID>.review.json がある）は、前の flow.json の番号を引き継ぐ（採否表・確認事項一覧・処理時間の番号とずらさない）。
    同じ箱の見分け: 種別と、出典に引いたカードの項目が同じもの。項目を引かない箱は、種別と名前が同じもの。
    前に無かった箱には、その種類の、前の番号のどれよりも大きい番号を付ける（消えた箱の番号は使い回さない）。返すのは、新しい番号を付けた箱の数"""
    old = try_load(pathlib.Path(work) / "flows" / f"{fl['flow_id']}.flow.json")[0] if (pathlib.Path(work) / "review" / f"{fl['flow_id']}.review.json").exists() else None
    if not isinstance(old, dict):
        return None
    numbered = lambda x: [n for n in x.get("nodes") or [] if isinstance(n, dict) and isinstance(n.get("id"), str) and n.get("type") in PREFIX]
    items = lambda n: tuple(sorted((str(s.get("doc_id")), str(s.get("page")), str(s.get("item"))) for s in n.get("sources") or [] if isinstance(s, dict) and s.get("item")))
    olds, news, ren, taken = numbered(old), numbered(fl), {}, set()
    for key in (lambda n: (n.get("type"), items(n)) if items(n) else None, lambda n: (n.get("type"), str(n.get("label") or ""))):
        pool = {}
        for n in olds:
            if key(n) is not None:
                pool.setdefault(key(n), []).append(n["id"])
        for n in news:
            if n["id"] not in ren and key(n) is not None:
                free = [i for i in pool.get(key(n), []) if i not in taken]
                if free:
                    ren[n["id"]] = free[0]
                    taken.add(free[0])
    used, fresh = {n["id"] for n in olds}, 0
    for n in news:
        if n["id"] not in ren:
            pre = PREFIX[n["type"]]
            top = max((int(i.rsplit("-", 1)[1]) for i in used | set(ren.values()) if i.rsplit("-", 1)[0] == pre and i.rsplit("-", 1)[1].isdigit()), default=0)
            ren[n["id"]] = f"{pre}-{top + 1:03d}"
            fresh += 1
    apply_ids(fl, {a: b for a, b in ren.items() if a != b})
    return fresh


def make(work, cid, edits):
    """(flow か None, 誤りの並び, 下書き)"""
    d = Draft(work, cid)
    b = Builder(d, edits)
    b.apply()
    fl = b.build() if not b.err else None
    if fl is not None and not b.err:
        renumber(fl)
        from systems import Ledger, has_ledger, place_entries
        if has_ledger(work):                                         # 紐付いたプログラムの起動点の箱（S-）を、その手順のすぐあとに置く
            place_entries(fl, Ledger(work))
        d.kept = keep_numbers(work, fl)
    return (fl if not b.err else None), b.err, d


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--flow", required=True, help="業務分類の id（F03）")
    ap.add_argument("--no-edits", action="store_true", help="直しの一覧を読まずに、下書きのままのフローを作る")
    a = ap.parse_args()
    work = pathlib.Path(work_dir(a.work))
    blocking(work, ("taxonomy", "hierarchy", "glossary", "systems"))       # flow.json は、ここで作り直すので見ない
    ef = work / "flows" / f"{a.flow}.edits.json"
    edits = {} if a.no_edits else load(ef)
    if edits is None:
        sys.exit(f"直しの一覧 {ef} が無い（下書きは python3 scripts/draft_flow.py -w {work} --flow {a.flow} で見る）")
    if not isinstance(edits, dict) or (edits and edits.get("flow_id") not in (None, a.flow)):
        sys.exit(f"{ef.name}: 全体がオブジェクトでないか、flow_id が {a.flow} でない")
    shape = [] if a.no_edits else validate(edits, "edits")
    if shape:
        print(f"{ef.name}: 形の違反 {len(shape)} 件（references/schemas.md の「直しの一覧」）。flow.json は書いていない:")
        for x in shape:
            print("  ★ " + x)
        sys.exit(1)
    fl, err, d = make(work, a.flow, edits)
    if err:
        print(f"直しの一覧に誤り {len(err)} 件。flow.json は書いていない:")
        for x in err:
            print("  ★ " + x)
        sys.exit(1)
    save(work / "flows" / f"{a.flow}.flow.json", fl)
    n = sum(1 for x in fl["nodes"] if x["type"] in ("process", "decision", "system", "subprocess"))
    mech = sum(1 for q in fl["questions"] if q["id"].startswith("m-"))
    print(f"flows/{a.flow}.flow.json を書いた: 作業のノード {n}・線 {len(fl['edges'])}・確認事項 {len(fl['questions'])} 件（うち機械が作ったもの {mech}）"
          + ("" if d.kept is None else f"。原典照合のあとなので、前の番号を引き継いだ（新しい番号を付けた箱 {d.kept} 個）"))
    for x in d.notes:
        print("  注意: " + x)
    lane = {x["id"]: x["name"] for x in fl["lanes"]}
    print("\nできたフロー（番号・レーン・ラベル → 行き先）:")
    for x in fl["nodes"]:
        if x["type"] != "file":
            outs = " ／ ".join((f"「{e['label']}」" if e.get("label") else "") + ("（推定）" if e.get("inferred") else "") + "→ " + e["to"] for e in fl["edges"] if e["from"] == x["id"])
            print(f"  {x['id']:8} [{lane[x['lane']]}] {x['label'][:36]}  {outs}")
    print("確認事項: " + "、".join(f"{q['id']}（{q.get('node') or '・'.join(q.get('nodes') or q.get('edge') or [])}）" for q in fl["questions"]))


if __name__ == "__main__":
    main()
