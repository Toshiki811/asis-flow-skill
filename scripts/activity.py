#!/usr/bin/env python3
"""activity.py — プログラムのアクティビティ図のデータを作る。機械が骨組み（下書き）を作り、LLM は「直しの一覧」だけを書く。

  python3 scripts/activity.py list  -w <作業フォルダ>                 # 図にする単位（A-02-1 …）と、進み具合（呼ばれる側から先の順）
  python3 scripts/activity.py list  -w <作業フォルダ> --group 4       # サブエージェント・照合の担当に渡す組（4 枚ずつ。呼ばれる側から先）
  python3 scripts/activity.py draft -w <作業フォルダ> --id A-02-1     # 下書きを、ソースの行つきで表示する
  python3 scripts/activity.py apply -w <作業フォルダ> --id A-02-1     # 下書き＋直しの一覧（activities/A-02-1.edits.json）→ activities/A-02-1.activity.json
  python3 scripts/activity.py apply -w <作業フォルダ> --all           # 直しの一覧がある単位を、全部作る（呼ばれる側から先に）
  python3 scripts/activity.py apply -w <作業フォルダ> --id A-02-1 --no-edits   # 直しの一覧なしで、下書きのままの図を作る（確かめ用）
  python3 scripts/activity.py terms -w <作業フォルダ> [--id A-02-1]   # プログラムの中の名前の候補と、用語の対応（structure/code_terms.json）への登録の状況

下書きの決め方（同じソースからは、いつも同じ下書き）:
  単位      起動点ごとに1つ。呼ばれる手続きは、呼ぶ所が1か所だけで箱が 8 個以下なら呼ぶ側に入れ、ほかは別の単位（呼び出しの箱からリンクする）
  項目      構造（code/*.struct.json）の木のノードが、そのまま項目になる。キーは「手続き.先頭の行」（Main.23）
  箱        続く文のまとまり → 処理の箱／分岐 → 判断／繰り返し → 中身が箱1つなら「繰り返す箱」1つ、そうでなければ判断と戻りの線／
            操作者とのやり取り（メッセージ・入力・ファイルの選択）→ 操作者のレーンの箱／別の単位の呼び出し → 呼び出しの箱
  対象外    機械が技術的と分かる文（画面の更新の制御・変数の準備・ログ・後片付け・エラーの扱いの設定）は、はじめから対象外（include で戻せる）
  エラー    On Error GoTo の所に「このあとの処理でエラーが起きたか」の判断を置き、起きた側にエラー処理の節をつなぐ
  まとめの案  箱が 25 個を超える図だけ: 同じ共通の処理を 3 回以上続けて呼ぶ並びを、ソースのコメントの見出しで切って、merge の items の案を出す
            （呼び出しが 3 回に満たない見出しは、前の見出しと合わせる。先頭なら後ろと。案を見せるだけで、使うか・名前は LLM）
直しの種類（ops[].op）: name 名付ける／merge まとめる／split 分ける／fold 畳む（技術的な分岐・繰り返しを1つの箱に）／exclude 対象外にする／include 図に戻す
直しの一覧に誤りがあれば、何も書かずに、誤りを全部並べて終わる（終了コード 1）。
"""
import argparse
import copy
import json
import re
import sys

import code_struct as cs
import terms as T
from common import EFFECTS, SEV_LABEL, load, load_page, page_json, page_units, save, severity_of, site, span_text, work_dir
from renumber import renumber
from systems import Ledger, has_ledger
from validate import blocking, validate

OPS = ("name", "merge", "split", "fold", "exclude", "include")
FOLD_REASONS = ("保存先・ファイルの有無の確認", "型や空の確認", "エラー番号の振り分け", "書式・表示の調整", "再試行・待ち合わせ", "配列・文字の加工")
EXCLUDE_REASONS = ("画面の更新の制御", "変数の準備", "ログ", "後片付け", "エラーの扱いの設定", "書式・表示の調整", "開発・試験用")
MAX_BOXES = 25
CALL_RUN = 3                                                         # 同じ共通の処理を、これ以上続けて呼ぶ並びは、1つの箱にまとめられる
OPERATOR, EXTERNAL = "操作者", "外部"
ERR_LABEL = "このあとの処理でエラーが起きたか"
ENDLESS = re.compile(r"(?i)^\(?\s*(true|1|-1)?\s*\)?$|^\s*(while\s*\(\s*(true|1)\s*\)|for\s*\(\s*;\s*;\s*\)|loop|perform\s+forever)\s*\{?\s*$")
ENDLESS_WHY = "条件の無い繰り返し（戻る線で表す。終わるのは、中の「抜ける」）"
FRAMEWORK_NOTE = "  ※枠組みが、この起動点の手続きの前に呼ぶ（木の起動点の before）"
QUITS_WHY = "プログラムを止める手続きを呼ぶ（「終了（プログラムを止める）」で表す）"


def endless(n):
    """構造の繰り返しのノードが、条件の無い繰り返しか（Do … Loop・While True・while (true)・for (;;)）。回数や集まりを回す繰り返しは、当たらない"""
    form = str(n.get("form") or "").lower()
    cond = re.sub(r"(?i)^(while|until)\s+", "", str(n.get("cond") or "").strip()) if form.startswith(("do", "while")) else str(n.get("cond") or "").strip()
    return not form.startswith("for") and bool(ENDLESS.match(cond))


# やり取りの箱に、機械が付ける名前（動詞で終える）: 見せる文字が無いとき／あるとき（{} に、その文字）
INTERACT_LABEL = {"msgbox": ("メッセージを出す", "「{}」と、メッセージを出す"), "inputbox": ("入力を求める", "「{}」と、入力を求める"),
                  "filedialog": ("ファイルを選ばせる", "「{}」の画面で、ファイルを選ばせる"), "form": ("フォームを開く", "フォーム「{}」を開く")}
INTERACT_OTHER = ("操作者に知らせる", "「{}」と、操作者に知らせる")
# 外部のレーンの箱に、機械が付ける名前（動詞で終える。相手の名前は、かっこ書き）: (種類, 向き) → 名前
EXTERNAL_LABEL = {("メール", "out"): "メールを送る", ("メール", "in"): "メールを受け取る", ("外部プログラム", "out"): "外部プログラムを動かす", ("外部プログラム", "in"): "外部プログラムを動かす",
                  ("外部システム", "out"): "外部システムに渡す", ("外部システム", "in"): "外部システムから受け取る"}
EXTERNAL_IO = ("メール", "外部プログラム", "外部システム")                      # 外部のレーンの箱にする入出力（ほかは、帳票と同じ付け方）


def placeholder(module, a, b):
    return f"（{module} {a} 行）" if a == b else f"（{module} {a}〜{b} 行）"


class Draft:
    """図にする単位1つの下書き: 木（items）と、キー → 項目（by）"""

    def __init__(self, led, unit):
        self.led, self.unit = led, unit
        self.st = led.structs[unit["project"]]
        self.procs = {p["id"]: p for p in self.st["procedures"]}
        self.inline = set(self.st.get("inline") or [])
        self.blank = set(self.st.get("blank") or [])                 # 図にする文が無い手続き（呼び出しを対象外にする）
        self.folded = {k.split(".", 1)[1]: why for k, why in led.folded.items() if k.startswith(unit["project"] + ".")}   # 図にしない手続き（補助）→ 理由
        self.by, self.notes = {}, []
        root = self.procs[unit["procedure"]]
        self.scope = [root["id"]]                                    # この図に入る手続き（起動点の手続きと、呼ぶ側に入れた手続き）
        self.handlers = {}                                           # 手続き → エラー処理の節（呼ぶ側の図に入れた手続きの分も）
        self.items = self.convert(root, root["body"])
        self.handlers[root["id"]] = [(h["label"], self.convert(root, h["body"])) for h in root.get("handlers") or []]

    def key(self, proc, line):
        base, k = f"{proc['name']}.{line}", 1
        key = base
        while key in self.by:
            k += 1
            key = f"{base}.{k}"
        return key

    def convert(self, proc, nodes):
        out = []
        for n in nodes:
            it = {"key": self.key(proc, n["lines"][0]), "kind": n["kind"], "proc": proc["id"], "module": proc["module"], "lines": list(n["lines"]),
                  "exec": list(n.get("exec") or []), "io": list(n.get("io") or []), "state": "box", "reason": "", "label": "", "detail": ""}
            self.by[it["key"]] = it
            out.append(it)
            k = n["kind"]
            if k == "seq":
                it["hint"] = n.get("hint", "")
                if n.get("tech"):
                    it["state"], it["reason"] = "excluded", n["tech"]
            elif k in ("if", "select"):
                it.update(cond=n.get("cond", ""), head=n.get("head"), single=bool(n.get("single")), interact=n.get("interact"),
                          branches=[{"label": b["label"], "cond": b.get("cond", ""), "name": "", "body": self.convert(proc, b["body"])} for b in n["branches"]])
            elif k == "loop":
                it.update(cond=f"{n.get('form', '')} {n.get('cond', '')}".strip(), head=n.get("head"), body=self.convert(proc, n["body"]), names=["", ""], endless=endless(n))
            elif k == "call":
                if n.get("to") in self.folded and n["to"] not in self.blank:   # 図にしない手続き（補助）の呼び出しは、ふつうのまとまりにする（名前を付ける。となりとまとめられる）
                    it.update(kind="seq", hint="", helper=n["to"])
                    if self.folded[n["to"]] == "ログの表示":           # ログの表示だけをする補助を呼ぶ行は、ほかのログの文と同じく、機械が対象外にする
                        it["state"], it["reason"] = "excluded", "ログ"
                    continue
                it["to"], it["framework"] = n.get("to"), bool(n.get("framework"))
                if n.get("to") in self.blank:
                    it["quits"] = bool(self.procs[n["to"]].get("ends"))   # 呼び先が、プログラムごと終える（そのあとは、途中の終了）
                    it["state"], it["reason"] = "excluded", QUITS_WHY if it["quits"] else "呼び先に、図にする文が無い"
                if n.get("to") in self.inline and n["to"] not in self.scope:
                    callee = self.procs[n["to"]]
                    self.scope.append(callee["id"])
                    it["inline"] = self.convert(callee, callee["body"])
                    self.handlers[callee["id"]] = [(h["label"], self.convert(callee, h["body"])) for h in callee.get("handlers") or []]
            elif k == "interact":
                it["interact"] = n.get("interact") or {}
                plain, shown = INTERACT_LABEL.get(it["interact"].get("how"), INTERACT_OTHER)
                it["label"] = shown.format(it["interact"]["text"][:60]) if it["interact"].get("text") else plain
            elif k == "onerror":
                it["mode"], it["target"] = n.get("mode"), n.get("label")
                if n.get("mode") != "goto":
                    it["state"], it["reason"] = "excluded", "エラーの扱いの設定"
            elif k == "exit":
                it["what"], it["target"] = n.get("what"), n.get("label")
            elif k == "label":
                it["name"] = n.get("name")
        return out

    def walk(self, items=None):
        for it in self.items if items is None else items:
            yield it
            for b in it.get("branches") or []:
                yield from self.walk(b["body"])
            for key in ("body", "inline"):
                if it.get(key):
                    yield from self.walk(it[key])
        if items is None:
            for hs in self.handlers.values():
                for _lab, body in hs:
                    yield from self.walk(body)

    def boxes(self):
        """下書きのままの箱の数（呼ぶ側の図に入れた呼び出しは、中身で数える。対象外は数えない）"""
        return sum(1 for it in self.walk() if it["state"] == "box"
                   and (it["kind"] in ("seq", "interact", "if", "select", "loop") or (it["kind"] == "call" and it.get("inline") is None)))

    def heading(self, it, prev):
        """項目のすぐ上（前の項目 prev のあと）の、コメントだけの行の文字（ソースの見出し）。間に文の行があれば、そこまで。無ければ空"""
        a, b = (prev["lines"][1] + 1) if prev else max(1, it["lines"][0] - 20), it["lines"][0] - 1
        rows = [t for _ln, t in self.source(it["module"], a, b)] if a <= b else []
        if self.st.get("language") in cs.VBA_FAMILY:
            kinds = [("note", x) if not code.strip() and x else ("code", "") if code.strip() else ("", "") for code, x in map(cs.strip_comment, rows)]
        else:
            import code_lang
            keep, notes = code_lang.plain_lines(self.st.get("language"), rows)
            kinds = [("note", notes[k]) if notes.get(k) else ("code", "") if k in keep else ("", "") for k in range(1, len(rows) + 1)]
        for kind, text in reversed(kinds):                           # 下から: 空の行・かっこだけの行は飛ばし、文の行で止まる
            if kind:
                return text if kind == "note" else ""
        return ""

    def call_runs(self):
        """まとめの案（箱が MAX_BOXES 個を超える図だけ）: [{to, times, heads, keys}]。references/code-guide.md の「同じ共通の処理を続けて呼ぶ並び」の決まりで、機械が切る:
        並びは、分岐・繰り返し・やり取り・抜ける所・まとめられない呼び出し・呼び出しの無い見出しで切れる。並びの中は、ソースのコメントの見出しで切る。
        見出しの呼び先は、いちばん多く呼ぶ共通の処理。呼ぶ回数が CALL_RUN に満たない見出しは、同じ呼び先の前の見出しと合わせる（前に無ければ、後ろと）。
        CALL_RUN 回以上になった並びが案。items は、はじめの呼び出しの前の「まとまり」（値を取る）から、終わりの呼び出しまで（前の対象外の文と、あとの文は入れない）"""
        if self.boxes() <= MAX_BOXES:
            return []
        proj = self.unit["project"]
        ok = lambda it: (it["kind"] == "call" and it["state"] == "box" and it.get("inline") is None and not it.get("framework")
                         and not (self.procs.get(it["to"]) or {}).get("ends") and bool(self.led.activity.get((proj, it["to"]))))
        holders = [self.items] + [b["body"] for x in self.walk() for b in x.get("branches") or []] \
            + [x[k] for x in self.walk() for k in ("body", "inline") if x.get(k)] + [body for hs in self.handlers.values() for _l, body in hs]
        out = []
        for holder in holders:
            secs = []                                                # 見出しの節 {head, items}。None は並びの切れ目
            for k, it in enumerate(holder):
                if not (it["kind"] == "seq" or ok(it)):
                    secs.append(None)
                    continue
                head = self.heading(it, holder[k - 1] if k else None)
                if not secs or secs[-1] is None or head:
                    secs.append({"head": head, "items": []})
                secs[-1]["items"].append(it)
            groups = []                                              # 呼び出しのある、続く見出しの組
            for s in secs + [None]:
                tos = [it["to"] for it in s["items"] if ok(it)] if s else []
                if not tos:
                    groups.append(None)
                    continue
                to = max(dict.fromkeys(tos), key=tos.count)
                groups.append({"to": to, "times": tos.count(to), "secs": [s]})
            k = 0
            while k < len(groups):                                   # 呼び出しが足りない見出しは、前（同じ呼び先）と。前に無ければ、後ろと
                g = groups[k]
                if g is None or g["times"] >= CALL_RUN:
                    k += 1
                    continue
                prev = groups[k - 1] if k and groups[k - 1] and groups[k - 1]["to"] == g["to"] else None
                nxt = groups[k + 1] if k + 1 < len(groups) and groups[k + 1] and groups[k + 1]["to"] == g["to"] else None
                into = prev or nxt
                if into is None:
                    k += 1
                    continue
                into["secs"] = into["secs"] + g["secs"] if into is prev else g["secs"] + into["secs"]
                into["times"] += g["times"]
                del groups[k]
                k = max(k - 1, 0)
            for g in groups:
                if g is None or g["times"] < CALL_RUN:
                    continue
                its = [it for s in g["secs"] for it in s["items"]]
                last = max(i for i, it in enumerate(its) if ok(it))
                first = next(i for i, it in enumerate(its) if it["state"] == "box")
                out.append({"to": g["to"], "times": g["times"], "heads": [s["head"] for s in g["secs"]], "keys": [it["key"] for it in its[first:last + 1]]})
        return out

    def exec_lines(self):
        """この図が受け持つ実行文の行: {(モジュール, 行)}"""
        return {(self.procs[p]["module"], ln) for p in self.scope for ln in self.procs[p]["exec"]}

    # ---- 見せる ----

    def source(self, module, a, b):
        out = []
        for c in cs.cite(self.st, module, a, b):
            units = page_units(load_page(page_json(self.led.work, c["doc_id"], c["page"])))
            at = next(pg["line_from"] for m in self.st["modules"] if m["name"] == module for pg in m["pages"]
                      if pg["page"] == c["page"] and (pg.get("doc_id") or m["doc_id"]) == c["doc_id"])
            out += [(at + k - 1, units[k - 1]) for k in range(c["lines"][0], c["lines"][1] + 1)]
        return out

    def show(self, out=print):
        u, e = self.unit, None
        if u["kind"] == "entry":
            e = self.led.entry(f"{u['project']}.{u['entry']}")
        root = self.procs[u["procedure"]]
        out(f"== 下書き {u['activity']}  {u['system']['name']}  " + (f"起動点 {e['key']}「{e['label']}」" if e else "共通の処理（ほかの単位から呼ばれる）")
            + f"  → {root['id']}（{root['module']} {root['lines'][0]}〜{root['lines'][1]} 行）")
        if len(self.scope) > 1:
            out("呼ぶ側に入れた手続き: " + "・".join(self.scope[1:]))
        out("項目（キー・種類・行・機械が拾ったこと）。行の頭の数字は、モジュールの行の番号:")

        shown = set()                                                # 中身を見せた、補助の手続き（1つの下書きに、1回だけ見せる）

        def dump(items, depth):
            pad = "  " * depth
            for it in items:
                k, (a, b) = it["kind"], it["lines"]
                io = "、".join(f"{x['kind']}「{x['name']}」{'に書く' if x['way'] == 'out' else 'を読む'}" for x in it.get("io") or [])
                off = f"  〈対象外: {it['reason']}〉" if it["state"] == "excluded" else ""
                if k == "seq":
                    out(f"{pad}{it['key']}  まとまり  {a}〜{b} 行{'  コメント「' + it['hint'] + '」' if it.get('hint') else ''}{'  ' + io if io else ''}{off}")
                    for ln, text in self.source(it["module"], a, b):
                        out(f"{pad}    {ln:>4}| {text}")
                    for via in {(x["via"]["module"], *x["via"]["lines"]) for x in it.get("io") or [] if x.get("via")}:   # 動かすクエリの定義（SQL）
                        out(f"{pad}    （クエリの定義 {via[0]} {via[1]}〜{via[2]} 行）")
                        for ln, text in self.source(*via):
                            out(f"{pad}    {ln:>4}: {text}")
                    if it.get("helper"):                             # 補助の手続きを呼ぶ行: その手続きは図にしないので、この箱に、していることを名前で書く
                        hp = self.procs[it["helper"]]
                        out(f"{pad}    （補助の手続き {it['helper']} を呼ぶ。図にしない: {self.folded.get(it['helper'], '')}。この箱の名前に、何のために呼ぶかを書く"
                            + ("。中身は、上に出した）" if it["helper"] in shown else "。下が、その中身。図には出ない）"))
                        if it["helper"] not in shown:
                            shown.add(it["helper"])
                            body = self.source(hp["module"], *hp["lines"])
                            for ln, text in body[:40]:
                                out(f"{pad}      {ln:>4}: {text}")
                            if len(body) > 40:
                                out(f"{pad}      （ほか {len(body) - 40} 行）")
                elif k in ("if", "select"):
                    out(f"{pad}{it['key']}  分岐  {a}〜{b} 行  条件: {it['cond']}" + ("  ※操作者に聞く" if it.get("interact") else "") + (f"  {io}" if io else ""))
                    for j, br in enumerate(it["branches"], 1):
                        out(f"{pad}  ・出口{j}（{br['label']}{' ' + br['cond'] if br['cond'] and br['label'] != 'then' else ''}）")
                        dump(br["body"], depth + 2)
                elif k == "loop":
                    out(f"{pad}{it['key']}  繰り返し  {a}〜{b} 行  {it['cond']}" + (f"  {io}" if io else "")
                        + ("  ※条件が無い: 判断にしない（中の最後から頭へ戻る線で表す。終わるのは、中の「抜ける」だけ）" if it.get("endless") else ""))
                    dump(it["body"], depth + 1)
                elif k == "call":
                    if it.get("inline") is not None:
                        out(f"{pad}{it['key']}  呼び出し  {a} 行  → {it['to']}（この図に入れる。下が、その中身）" + (FRAMEWORK_NOTE if it.get("framework") else ""))
                        dump(it["inline"], depth + 1)
                    else:
                        where = self.led.activity.get((u["project"], it["to"]))
                        out(f"{pad}{it['key']}  呼び出し  {a}{'' if a == b else f'〜{b}'} 行  → {it['to']}（{'別の図 ' + where if where else '図にする文の無い手続き' if it['to'] in self.blank else '別の図 ？'}）"
                            f"{'  ' + io if io else ''}" + (FRAMEWORK_NOTE if it.get("framework") else ""))
                        for ln, text in [] if it.get("framework") else self.source(it["module"], a, b):
                            out(f"{pad}    {ln:>4}| {text}")
                elif k == "interact":
                    out(f"{pad}{it['key']}  やり取り  {a} 行  {it['label']}")
                elif k == "exit":
                    out(f"{pad}{it['key']}  抜ける  {a} 行  " + {"proc": "この処理を終える", "loop": "繰り返しを抜ける", "continue": "繰り返しの次の回へ進む", "select": "振り分けを抜ける（次へ進む）", "program": "プログラムごと終える（呼んだ側に戻らない）", "goto": f"ラベル {it.get('target')} へ飛ぶ", "resume": "処理に戻る"}[it["what"]])
                elif k == "onerror":
                    out(f"{pad}{it['key']}  エラーの設定  {a} 行  " + (f"エラーが起きたら {it['target']} へ" if it["mode"] == "goto" else "エラーが起きても続ける／設定を解く") + off)
                elif k == "label":
                    out(f"{pad}（ラベル {it['name']}  {a} 行）")
        dump(self.items, 0)
        for lab, body in self.handlers.get(self.unit["procedure"]) or []:
            out(f"エラーが起きたとき（{lab}）:")
            dump(body, 1)
        runs = self.call_runs()
        if runs:
            out(f"まとめの案（箱が {self.boxes()} 個で、{MAX_BOXES} 個を超える。同じ共通の処理を {CALL_RUN} 回以上続けて呼ぶ並びを、ソースの見出しで切った。"
                f"呼び出しが {CALL_RUN} 回に満たない見出しは、前の見出しと合わせた。使うなら merge の items に写し、何を・何項目かを名前にする。references/code-guide.md）:")
            for k, r in enumerate(runs, 1):
                where = self.led.activity.get((u["project"], r["to"]))
                out(f"  案{k}  → {where} {r['to']}（{r['times']} 回）  見出し: {'・'.join('「' + h + '」' for h in r['heads'] if h) or '無し'}")
                out(f"        \"items\": {json.dumps(r['keys'], ensure_ascii=False)}")
        for n in self.notes:
            out("注意: " + n)


# ---------- 直しを当てる ----------

class Builder:
    def __init__(self, draft, edits):
        self.d, self.e, self.err = draft, edits or {}, []
        self.by = draft.by

    def bad(self, i, op, msg):
        self.err.append(site(f"ops[{i}]（{op.get('op')}）: {msg}", 2))

    def need(self, i, op, key, kinds=None):
        it = self.by.get(key)
        if it is None:
            self.bad(i, op, f"「{key}」という項目が、下書きに無い")
        elif kinds and it["kind"] not in kinds:
            self.bad(i, op, f"「{key}」は{'・'.join(kinds)}でない（{it['kind']}）")
            it = None
        return it

    def siblings(self, it):
        for holder in [self.d.items] + [b["body"] for x in self.d.walk() for b in x.get("branches") or []] \
                + [x[k] for x in self.d.walk() for k in ("body", "inline") if x.get(k)] + [body for hs in self.d.handlers.values() for _l, body in hs]:
            if any(x is it for x in holder):
                return holder
        return None

    def apply(self):
        ops = [x for x in self.e.get("ops") or [] if isinstance(x, dict)]
        for phase in (("split",), ("merge",), ("fold", "exclude", "include"), ("name",)):
            for i, op in enumerate(ops):
                if op.get("op") not in OPS:
                    if phase == ("split",):
                        self.bad(i, op, f"決まっていない種類。使えるのは {'・'.join(OPS)}")
                elif op["op"] in phase:
                    getattr(self, "op_" + op["op"])(i, op)

    def op_split(self, i, op):
        it = self.need(i, op, op.get("item"), ("seq",))
        cuts = sorted({x for x in op.get("at") or [] if isinstance(x, int)})
        if it is None:
            return
        if not cuts or any(not (it["lines"][0] < c <= it["lines"][1]) or c not in it["exec"] for c in cuts):
            return self.bad(i, op, f"at（分ける行）は、「{it['key']}」の2つ目以降の実行文の行で書く（{it['lines'][0] + 1}〜{it['lines'][1]} 行のうち、実行文のある行）")
        holder, proc = self.siblings(it), self.d.procs[it["proc"]]
        at, parts, bounds = holder.index(it), [], [it["lines"][0]] + cuts + [it["lines"][1] + 1]
        ex0, io0, first = list(it["exec"]), list(it["io"]), it["lines"][0]
        for a, b in zip(bounds, bounds[1:]):
            ex = [x for x in ex0 if a <= x < b]
            part = it if a == first else dict(copy.deepcopy(it), key=f"{proc['name']}.{a}", hint="")
            part.update(lines=[a, max(ex)], exec=ex, io=[])
            for x in io0:                                            # 同じ先を読む・書く行が、この部分にあれば付ける（はじめの行だけで決めない）
                at_lines = [ln for ln in [x["line"]] + (x.get("also") or []) if a <= ln < b]
                if at_lines:
                    rest = dict(x, line=at_lines[0])
                    rest.pop("also", None)
                    part["io"].append(dict(rest, also=at_lines[1:]) if at_lines[1:] else rest)
            self.by[part["key"]] = part
            parts.append(part)
        holder[at:at + 1] = parts

    def op_merge(self, i, op):
        its = [self.need(i, op, k, ("seq", "call")) for k in op.get("items") or []]
        if None in its or len(its) < 2:
            return self.bad(i, op, "items に、まとめる「まとまり」を2つ以上、並んでいる順に書く") if None not in its else None
        calls = [it for it in its if it["kind"] == "call"]
        if calls:                                                    # 同じ共通の処理（別の図）を続けて呼ぶ並び
            if any(it.get("inline") is not None or it["state"] != "box" or it.get("framework") or (self.d.procs.get(it["to"]) or {}).get("ends") for it in calls):
                return self.bad(i, op, "まとめられる呼び出しは、別の図になる共通の処理の呼び出しだけ（呼ぶ側の図に入れた手続き・図にする文の無い手続き・プログラムを止める手続きの呼び出しは、まとめられない）")
            tos = [it["to"] for it in calls]
            if max(tos.count(t) for t in tos) < CALL_RUN:
                return self.bad(i, op, f"呼び出しをまとめられるのは、同じ共通の処理を {CALL_RUN} 回以上続けて呼ぶ並びだけ（値を取る → 共通の処理に渡す、の繰り返し）")
        holder = self.siblings(its[0])
        idx = [next((k for k, x in enumerate(holder) if x is it), None) for it in its]
        if None in idx or idx != list(range(idx[0], idx[0] + len(idx))):
            between = "（間に、分岐・繰り返し・呼び出し・やり取りがある）" if None not in idx else "（同じ並びの中に無い）"
            return self.bad(i, op, f"まとめられるのは、同じ並びで隣り合う「まとまり」だけ{between}。機械が対象外にした文をはさむときは、それも items に入れる")
        first = its[0]
        if calls:
            first["calls"] = [it["to"] for it in calls]
            first["kind"] = "seq"
        for it in its[1:]:
            first["exec"] = sorted(set(first["exec"] + it["exec"]))
            first["lines"][1] = max(first["lines"][1], it["lines"][1])
            first["io"] += [x for x in it["io"] if (x["kind"], x["name"], x["way"]) not in [(y["kind"], y["name"], y["way"]) for y in first["io"]]]
            del self.by[it["key"]]
        first["state"], first["reason"] = "box", ""
        holder[idx[0]:idx[0] + len(idx)] = [first]

    def op_fold(self, i, op):
        it = self.need(i, op, op.get("item"), ("if", "select", "loop"))
        if it is None:
            return
        if op.get("reason") not in FOLD_REASONS:
            return self.bad(i, op, f"reason は、決まった語から選ぶ（{'／'.join(FOLD_REASONS)}）。当てはまらなければ、畳まずに判断として描く")
        if it.get("interact") or any(x["kind"] in ("interact", "call") and x["state"] == "box" for x in self.d.walk([it]) if x is not it):
            return self.bad(i, op, "中に、操作者とのやり取り・別の処理の呼び出しがある分岐・繰り返しは、畳めない")
        if any(x["kind"] == "exit" and x.get("what") in ("proc", "program", "goto") for x in self.d.walk([it]) if x is not it):   # 決まり 6.11: 抜ける所は、図の終了の丸になる
            return self.bad(i, op, "中に、抜ける所（Exit Sub・Exit Function・End など）がある分岐・繰り返しは、畳めない（判断として描き、抜ける側を終了にする）")
        it["state"], it["reason"] = "folded", op["reason"]

    def op_exclude(self, i, op):
        it = self.need(i, op, op.get("item"), ("seq", "onerror"))
        if it is None:
            return
        if op.get("reason") not in EXCLUDE_REASONS:
            return self.bad(i, op, f"reason は、決まった語から選ぶ（{'／'.join(EXCLUDE_REASONS)}）。当てはまらなければ、箱として描く")
        if any(x["way"] == "out" for x in it.get("io") or []):
            return self.bad(i, op, "シート・ファイル・テーブルに書く文は、対象外にできない（箱として描く）")
        it["state"], it["reason"] = "excluded", op["reason"]

    def op_include(self, i, op):
        it = self.need(i, op, op.get("item"), ("seq",))
        if it is not None:
            it["state"], it["reason"] = "box", ""

    def op_name(self, i, op):
        it = self.need(i, op, op.get("item"))
        if it is None:
            return
        if it["kind"] in ("exit", "label") or (it["kind"] == "call" and it.get("inline") is not None):
            return self.bad(i, op, f"「{it['key']}」は箱にならない項目（{it['kind']}）")
        if not str(op.get("label", "")).strip():
            return self.bad(i, op, "label が空")
        it["label"], it["named"] = str(op["label"]).strip(), True
        if "detail" in op:
            it["detail"] = str(op["detail"]).strip()
        names = op.get("branches")
        if names is not None:
            want = len(it["branches"]) + (0 if any(b["label"] == "else" for b in it["branches"]) else 1) if it["kind"] in ("if", "select") else 2 if it["kind"] in ("loop", "onerror") else 0
            if not isinstance(names, list) or len(names) != want or not all(str(x).strip() for x in names):
                return self.bad(i, op, f"branches は、出口の数（{want}）だけ、下書きの出口の順に書く"
                                + ("（最後は、どの条件にも当たらない側）" if it["kind"] in ("if", "select") and want > len(it["branches"]) else "")
                                + ("（繰り返す側・終わる側の順）" if it["kind"] == "loop" else "") + ("（起きない側・起きた側の順）" if it["kind"] == "onerror" else ""))
            it["names"] = [str(x).strip() for x in names]

    # ---- 図のデータにする ----

    def build(self):
        d, u, led = self.d, self.d.unit, self.d.led
        self.nodes, self.edges, self.files, self.owned, self.branch_log = {}, [], [], [], []
        self.handled = set()                                         # つないだエラー処理の節: (手続き, ラベル)
        self.labels, self.ends = {}, set()
        root = d.procs[u["procedure"]]
        self.lane_of = {}
        entry = self.seq(d.items, "end", {"ret": "end2", "loop": None}, top=True)
        for a in self.edges:                                         # GoTo の行き先（ラベル）を、ノードに直す
            if a["to"].startswith("@"):
                a["to"] = self.labels.get(a["to"][1:].lower(), "end")
        entry = self.labels.get(entry[1:].lower(), "end") if entry.startswith("@") else entry
        edges = [{"from": "start", "to": entry}] + self.edges
        merged = {}
        for e in edges:                                              # 同じ2点を結ぶ線は1本にする（ラベルは「／」でつなぐ）
            k = (e["from"], e["to"])
            if k in merged:
                if e.get("label") and e["label"] not in str(merged[k].get("label", "")).split("／"):
                    merged[k]["label"] = "／".join(x for x in (merged[k].get("label"), e["label"]) if x)
            else:
                merged[k] = dict(e)
        edges = list(merged.values())
        used = {x for e in edges for x in (e["from"], e["to"])}
        names = [operator_name(self.e, led, u), u["system"]["name"], EXTERNAL]
        kinds = {names[0]: "role", names[1]: "system", names[2]: "external"}
        lanes_used = [nm for nm in names if any(n["lane"] == nm for i, n in self.nodes.items() if i in used)] or [names[1]]
        lid = {nm: f"L{k}" for k, nm in enumerate(lanes_used, 1)}
        prog = lid.get(names[1]) or lid[lanes_used[0]]
        nodes = [{"id": "start", "type": "start", "lane": prog, "label": "開始"}]
        made = list(self.nodes)[::-1]                                # 後ろから作ったので、逆にすると、ほぼソースの順
        for sid in [i for i in made if i.endswith("/x")]:            # 外部の箱は、渡す側の箱のすぐあと
            made.remove(sid)
            made.insert(made.index(sid[:-2]) + 1, sid)
        at = {i: k for k, i in enumerate(made)}
        edges.sort(key=lambda e: at.get(e["from"], -1))              # 線も、出どころの箱の順に（同じ箱から出る線の順は、そのまま）
        n_end = {}
        for e in edges:                                              # 途中で終わる所は、抜ける所ごとに、終了の丸を置く（1つの丸へ、長い線を集めない）
            if e["to"] in ("end2", "end3"):
                n_end[e["to"]] = k = n_end.get(e["to"], 0) + 1
                e["to"] += f"-{k}" if k > 1 else ""
        for i in made:
            n = self.nodes[i]
            if i in used:
                nodes.append(dict(n, lane=lid[n["lane"]]))
                nodes += [dict(f, lane=lid[n["lane"]]) for f in self.files if f["attach_to"] == i]
        for e_id, lab in (("end", "終了"), ("end2", "終了（途中で終わる）"), ("end3", "終了（プログラムを止める）")):
            for x in [e_id] + [f"{e_id}-{k}" for k in range(2, n_end.get(e_id, 0) + 1)]:
                src = [e["from"] for e in edges if e["to"] == x]
                if src:                                              # 終了の丸は、そこへ進む箱が1つなら、その箱と同じレーンに置く（箱のすぐ下に来る）
                    one = self.nodes.get(src[0]) if len(src) == 1 else None
                    nodes.append({"id": x, "type": "end", "lane": lid[one["lane"]] if one else prog, "label": lab})
        reach, todo = {"start"}, ["start"]
        while todo:
            x = todo.pop()
            for e in edges:
                if e["from"] == x and e["to"] not in reach:
                    reach.add(e["to"])
                    todo.append(e["to"])
        lost = [n["id"] for n in nodes if n["type"] != "file" and n["id"] not in reach]
        nodes = [n for n in nodes if n["id"] in reach or (n["type"] == "file" and n["attach_to"] in reach)]
        edges = [e for e in edges if e["from"] in reach]
        if lost:
            d.notes.append("開始からたどり着けない箱を外した（ソースで、Exit や GoTo のあとに書かれた文）: " + "・".join(lost))
        ent = led.entry(f"{u['project']}.{u['entry']}") if u["kind"] == "entry" else None
        act = {"flow_id": u["activity"], "name": str(self.e.get("name") or "").strip() or root["name"], "number": "", "kind": u["kind"],
               "system": {"id": u["system"].get("id"), "name": u["system"]["name"]}, "project": u["project"], "entry": ent["key"] if ent else None,
               "trigger": ent["label"] if ent else "ほかの処理から呼ばれる", "procedure": root["id"], "scope": list(d.scope), "tier": ("parsed" if root.get("read") == "axl" else "heuristic") if root.get("kind") == "macro" else d.st["tier"],
               "summary": str(self.e.get("summary") or "").strip(), "lanes": [{"id": lid[nm], "name": nm, "kind": kinds[nm]} for nm in lanes_used],
               "nodes": nodes, "edges": edges, "lines": self.owned, "branches": self.branch_log, "questions": []}
        ren = renumber(act)
        for o in self.owned + self.branch_log:
            if o.get("node"):
                o["node"] = ren.get(o["node"], o["node"])
        act["questions"] = self.questions(act, ren)
        return act

    def cite(self, it, a=None, b=None):
        a, b = (it["lines"] if a is None else (a, b))
        out = []
        for c in cs.cite(self.d.st, it["module"], a, b):
            text = span_text(page_units(load_page(page_json(self.d.led.work, c["doc_id"], c["page"]))), [c["lines"][0], c["lines"][0]]) or ""
            out.append(dict(c, quote=text.strip(), from_image=False))
        return out

    def own(self, it, node, state, lines=None):
        self.owned.append({"node": node, "key": it["key"], "module": it["module"], "lines": sorted(set(it["exec"] if lines is None else lines)), "state": state,
                           **({"reason": it["reason"]} if it.get("reason") else {})})

    def all_exec(self, it):
        return sorted({ln for x in self.d.walk([it]) for ln in x["exec"]})

    def add_node(self, it, ntype, lane, label, follow, **extra):
        nid = it["key"]
        named = it.get("named") or it["kind"] in ("interact", "call", "onerror")
        n = {"id": nid, "type": ntype, "lane": lane, "label": label or placeholder(it["module"], *it["lines"]), "sources": self.cite(it, *extra.pop("span", it["lines"]))}
        if it.get("detail"):
            n["detail"] = it["detail"]
        if not named and not label:
            n["unnamed"] = True
        n.update(extra)
        self.nodes[nid] = n
        if follow is not None:
            self.edges.append({"from": nid, "to": follow})
        for x in it.get("io") or [] if ntype != "decision" else []:
            if x["kind"] in EXTERNAL_IO:
                continue
            fid = f"{nid}/{x['way']}/{x['kind']}/{x['name']}"
            if fid not in [f["id"] for f in self.files]:
                self.files.append({"id": fid, "type": "file", "lane": lane, "label": f"{x['kind']} {x['name']}", "attach_to": nid, "direction": x["way"],
                                   "sources": self.cite(it, x["line"], x["line"]) + (self.cite({"module": x["via"]["module"]}, *x["via"]["lines"]) if x.get("via") else [])})
        return nid

    def external(self, it, nid, follow):
        """メール・外部プログラムに渡す文のあとに、外部のレーンの箱を置く"""
        ext = [x for x in it.get("io") or [] if x["kind"] in EXTERNAL_IO]
        if not ext:
            return
        x = ext[0]
        sid = f"{nid}/x"
        self.nodes[sid] = {"id": sid, "type": "system", "lane": EXTERNAL, "label": EXTERNAL_LABEL.get((x["kind"], x.get("way")), f"{x['kind']}を使う") + (f"（{x['name']}）" if x["name"] != x["kind"] else ""),
                           "sources": self.cite(it, x["line"], x["line"])}
        for e in self.edges:
            if e["from"] == nid and e["to"] == follow:
                e["to"], e["kind"] = sid, "system"
        self.edges.append({"from": sid, "to": follow, "kind": "system"})

    def seq(self, items, follow, cx, top=False):
        """項目の並びを、後ろから順にノードにする。戻り値は、この並びの入口（ノードの id。箱が1つも無ければ follow）"""
        d, prog, op = self.d, self.d.unit["system"]["name"], operator_name(self.e, self.d.led, self.d.unit)
        for at in range(len(items) - 1, -1, -1):
            it = items[at]
            k = it["kind"]
            last = top and all(x["kind"] == "label" for x in items[at + 1:])     # 本体の最後の文（そこでの Exit は、ふつうの終了）
            if k == "label":
                self.labels[it["name"].lower()] = follow
            elif it["state"] == "excluded":
                self.own(it, None, "excluded", self.all_exec(it))
                if it.get("quits"):
                    follow = "end" if last else "end3"
            elif it["state"] == "folded":
                nid = self.add_node(it, "process", prog, it["label"], follow, folded=it["reason"])
                self.own(it, nid, "folded", self.all_exec(it))
                self.branch_log += [{"key": x["key"], "state": "folded", "node": nid, "reason": it["reason"]} for x in d.walk([it]) if x["kind"] in ("if", "select", "loop")]
                follow = nid
            elif k == "seq":
                tos = [d.led.activity.get((d.unit["project"], to), "") for to in it.get("calls") or []]
                run = [{"to": a, "times": tos.count(a), "label": d.led.title(a, "")} for a in dict.fromkeys(tos)]
                nid = self.add_node(it, "process", prog, it["label"], follow, **({"calls": run} if run else {}))
                self.external(it, nid, follow)
                self.own(it, nid, "box")
                follow = nid
            elif k == "interact":
                nid = self.add_node(it, "process", op, it["label"], follow)
                self.own(it, nid, "box")
                follow = nid
            elif k == "call":
                if it.get("inline") is not None:
                    inner = self.seq(it["inline"], follow, dict(cx, ret=follow))
                    if it["exec"]:
                        self.owned.append({"node": None, "key": it["key"], "module": it["module"], "lines": it["exec"], "state": "inlined"})
                    follow = inner
                else:
                    aid = d.led.activity.get((d.unit["project"], it["to"]), "")
                    if (d.procs.get(it["to"]) or {}).get("ends") and not last:   # 呼び先が、いつもプログラムごと終える（異常終了の手続きなど）。呼んだ側には戻らない
                        follow = "end3"
                    nid = self.add_node(it, "subprocess", prog, it["label"] or d.led.title(aid, str(it["to"]).split(".")[-1]), follow, link_flow=aid)
                    self.own(it, nid, "box")
                    follow = nid
            elif k == "exit":
                self.own(it, None, "exit")
                follow = ("end" if last else "end3") if it["what"] == "program" else ("end" if last and it["what"] == "proc" else cx["ret"]) if it["what"] in ("proc", "resume") else \
                    (cx["loop"] or follow) if it["what"] == "loop" else (cx.get("again") or follow) if it["what"] == "continue" else \
                    follow if it["what"] == "select" else f"@{it.get('target') or ''}"
            elif k == "onerror":
                hs = d.handlers.get(it["proc"]) or []
                body = next((b for lab, b in hs if lab.lower() == str(it.get("target") or "").lower()), None)
                seen = (it["proc"], str(it.get("target") or "").lower())
                if seen in self.handled:                             # 同じ節への設定が、前にもある（節をつなぐのは、最初の1つだけ）
                    body = None
                self.handled.add(seen)
                inner = self.seq(body, "end2", dict(cx, ret="end2")) if body else "end2"
                if inner == "end2":                                  # エラー処理の節に、図にする文が無い
                    it["reason"] = it["reason"] or "エラーの扱いの設定"
                    self.own(it, None, "excluded")
                    continue
                names = it.get("names") or ["起きない", "起きた"]
                nid = self.add_node(it, "decision", prog, it["label"] or ERR_LABEL, None)
                self.edges += [{"from": nid, "to": follow, "label": names[0]}, {"from": nid, "to": inner, "label": names[1]}]
                self.own(it, nid, "box")
                self.branch_log.append({"key": it["key"], "state": "decision", "node": nid})
                follow = nid
            elif k in ("if", "select"):
                follow = self.branch(it, follow, cx, op if it.get("interact") else prog)
            elif k == "loop":
                follow = self.loop(it, follow, cx, prog)
        return follow

    def branch(self, it, follow, cx, lane):
        nid = it["key"]
        outs = [self.seq(b["body"], follow, cx) for b in reversed(it["branches"])][::-1]   # 後ろの出口から作る（できた箱の並びが、ソースの順になる）
        if all(o == follow for o in outs):                           # どの出口の先にも、図にする文が無い
            self.own(it, None, "excluded")
            it["reason"] = "分岐の先に、図にする文が無い"
            self.owned[-1]["reason"] = it["reason"]
            self.branch_log.append({"key": it["key"], "state": "skipped", "node": None, "reason": it["reason"]})
            return follow
        has_else = any(b["label"] == "else" for b in it["branches"])
        names = it.get("names") or []
        default = lambda j, b: {"then": "はい", "else": "いいえ" if it["kind"] == "if" and len(it["branches"]) == 2 else "その他"}.get(b["label"], b["cond"][:24])
        self.add_node(it, "decision", lane, it["label"] or it["cond"][:40], None, span=it.get("head") or it["lines"], **({} if it.get("named") else {"unnamed": True}))
        for j, (b, o) in enumerate(zip(it["branches"], outs)):
            self.edges.append({"from": nid, "to": o, "label": names[j] if j < len(names) else default(j, b)})
        if not has_else:
            self.edges.append({"from": nid, "to": follow, "label": names[len(it["branches"])] if len(names) > len(it["branches"]) else "いいえ" if len(it["branches"]) == 1 else "その他"})
        self.own(it, nid, "box")
        self.branch_log.append({"key": it["key"], "state": "decision", "node": nid})
        return nid

    def loop(self, it, follow, cx, lane):
        nid = it["key"]
        live = [x for x in it["body"] if x["state"] != "excluded" and x["kind"] != "label"]
        simple = len(live) == 1 and live[0]["kind"] == "seq" and live[0]["state"] == "box"
        if simple or not live:                                       # 中身が箱1つ（か、図にする文が無い）: 繰り返す箱1つ
            if not live:
                self.own(it, None, "excluded", self.all_exec(it))
                self.owned[-1]["reason"] = "繰り返しの中に、図にする文が無い"
                self.branch_log.append({"key": it["key"], "state": "skipped", "node": None, "reason": "繰り返しの中に、図にする文が無い"})
                return follow
            body = live[0]
            it["io"] = it["io"] + [x for x in body["io"] if (x["kind"], x["name"], x["way"]) not in [(y["kind"], y["name"], y["way"]) for y in it["io"]]]
            self.add_node(it, "process", lane, it["label"] or body["label"], follow, loop=it["cond"][:40],
                          **({} if it.get("named") or body.get("named") else {"unnamed": True}))
            self.own(it, nid, "box", self.all_exec(it))
            self.branch_log.append({"key": it["key"], "state": "loop_box", "node": nid})
            return nid
        names = it.get("names") if any(it.get("names") or []) else ["繰り返す", "終わり"]
        inner = self.seq(it["body"], nid, dict(cx, loop=follow, again=nid))
        if it.get("endless"):                                        # 条件の無い繰り返し（Do … Loop・while True）: 頭で抜ける道は無いので、判断にしない。中の最後から、中の頭へ戻る
            entry = follow if inner == nid else inner
            for e in self.edges:
                if e["to"] == nid:
                    e["to"] = entry
            it["reason"] = ENDLESS_WHY if entry is inner else "繰り返しの中に、図にする文が無い"
            self.own(it, None, "excluded")
            self.branch_log.append({"key": it["key"], "state": "endless" if entry is inner else "skipped", "node": None, "reason": it["reason"]})
            return entry
        self.add_node(it, "decision", lane, it["label"] or f"繰り返すか（{it['cond'][:30]}）", None, span=it.get("head") or it["lines"],
                      **({} if it.get("named") else {"unnamed": True}))
        self.edges += [{"from": nid, "to": inner, "label": names[0]}, {"from": nid, "to": follow, "label": names[1]}]
        self.own(it, nid, "box")
        self.branch_log.append({"key": it["key"], "state": "decision", "node": nid})
        return nid

    def questions(self, act, ren):
        out, seen = [], set()
        ids = {n["id"] for n in act["nodes"]}
        for i, x in enumerate(y for y in self.e.get("questions") or [] if isinstance(y, dict)):
            qid = str(x.get("id") or "")
            if not qid or qid in seen:
                self.err.append(site(f"questions[{i}]: id「{qid}」が空か、重なっている"))
                continue
            seen.add(qid)
            keys = [k for k in x.get("items") or []]
            nodes = [ren.get(k, k) for k in keys]
            if any(n not in ids for n in nodes):
                self.err.append(site(f"questions[{i}]（{qid}）: items が指す項目が、図の箱になっていない（{'・'.join(k for k, n in zip(keys, nodes) if n not in ids)}）"))
                continue
            if x.get("effect") not in EFFECTS:
                self.err.append(site(f"questions[{i}]（{qid}）: effect（答えで、図のどこが変わるか）が無い・決まった語でない（{'／'.join(EFFECTS)}）。"
                                "どれにも当たらない問い（値だけ・いまも同じかの確かめ・名前だけ）は、確認事項にしない"))
                continue
            q = {"id": qid, "severity": "", "category": x.get("category", "プログラム"), "effect": x["effect"]}
            if len(nodes) == 1:
                q["node"] = nodes[0]
            elif nodes:
                q["nodes"] = nodes
            q["severity"] = severity_of(q)                           # 重要度は決まり 5.4 で（答えで変わる所と、指す箱の数）
            if x.get("severity") not in (None, q["severity"]):
                self.d.notes.append(f"確認事項 {qid}: 重要度は、決まり 5.4 で{SEV_LABEL[q['severity']]}にした（直しの一覧の「{SEV_LABEL.get(x['severity'], x['severity'])}」は使わない）")
            src = [{"doc_id": s["doc_id"], "page": s["page"]} for n in act["nodes"] if n["id"] in nodes for s in n.get("sources") or []][:2]
            q.update(question=str(x.get("question", "")).strip(), assumption=str(x.get("assumption", "")).strip(),
                     sources=src or [{"doc_id": act["project"]}], ask=str(x.get("ask") or "").strip())
            out.append(q)
        return out


def operator_name(edits, led, unit):
    """操作者のレーンの名前: 直しの一覧の operator。無ければ、紐付けた手順の担当（1人に決まるとき）。それも無ければ「操作者」"""
    if str(edits.get("operator") or "").strip():
        return str(edits["operator"]).strip()
    if unit["kind"] == "entry":
        actors = {str(st[3].get("actor") or "").strip() for x in led.good_links() if x.get("entry") == f"{unit['project']}.{unit['entry']}"
                  for st in [led.step(x["step"])] if st}
        actors.discard("")
        if len(actors) == 1:
            return actors.pop()
    return OPERATOR


def make(led, unit, edits):
    """(図のデータ か None, 誤りの並び, 下書き)"""
    d = Draft(led, unit)
    b = Builder(d, edits)
    b.apply()
    act = b.build() if not b.err else None
    return (act if not b.err else None), b.err, d


def show_terms(led, tm, only=None):
    """プログラムの中の名前の候補（箱が受け持つ行に出てくるもの）と、用語の対応への登録の状況を、まとまりごとに出す"""
    found = {}
    for u in ordered(led):
        if (only and u["activity"] != only) or led.off(u) is not None:
            continue
        for tok, n, row, where in T.candidates(tm, Draft(led, u)):
            x = found.setdefault(u["project"], {}).setdefault(tok, [0, row, where, []])
            x[0] += n
            x[3].append(u["activity"])
    if not found:
        print("名前の候補は無い" + (f"（{only}）" if only else ""))
    for proj, rows in found.items():
        s = led.sys_of_proj.get(proj) or {}
        done = sum(1 for x in rows.values() if x[1])
        print(f"== プログラムの中の名前  {s.get('name', '')}（{proj}）  候補 {len(rows)}・登録済み {done}。業務の言葉を、{T.FILE} に書く"
              "（担当の目に見える名前＝シート・ファイル・メッセージは、載せない）")
        for tok, (n, row, where, acts) in sorted(rows.items(), key=lambda kv: (kv[1][1] is None and kv[1][2] is None, -kv[1][0], kv[0])):
            if not row and not where:
                continue
            state = ("＝（不明）" if row["basis"] == "unknown" else f"＝{row.get('word', '')}〔{T.BASIS.get(row.get('basis'), '')}〕") if row else "未登録"
            print(f"  {tok}  {n} 行（{'・'.join(acts[:4])}{' ほか' if len(acts) > 4 else ''}）  {state}" + (f"  定義: {where[0]} {where[1]}| {where[2][:80]}" if where and not row else ""))
        rest = [tok for tok, x in rows.items() if not x[1] and not x[2]]
        if rest:
            print(f"  このソースに定義の無い名前 {len(rest)} 個（部品・枠組みの名前か、読んでいない定義。業務の言葉にするものだけ載せる）: {'、'.join(sorted(rest)[:60])}" + (" ほか" if len(rest) > 60 else ""))
    for msg in tm.problems():
        print("  ★ " + msg)


def ordered(led):
    """図にする単位を、呼ばれる側から先に（呼び出しの箱の名前に、呼ばれる側の図の名前を使う）。共通の処理どうしも、呼ばれる側が先"""
    units = sorted(led.units, key=lambda u: (u["kind"] != "shared", u["activity"]))
    by = {u["activity"]: u for u in units}

    def callees(proj, pid, seen):
        """その手続きから呼ぶ、別の図（呼ぶ側の図に入る手続きは、その先までたどる）"""
        out = []

        def walk(nodes):
            for n in nodes or []:
                for to in ([n.get("to")] if n.get("kind") == "call" else []) + list(n.get("calls") or []):
                    if led.activity.get((proj, to)):
                        out.append(led.activity[(proj, to)])
                    elif to and to not in seen:
                        seen.add(to)
                        out.extend(callees(proj, to, seen))
                for b in n.get("branches") or []:
                    walk(b.get("body"))
                walk(n.get("body"))
        p = next((x for x in (led.structs.get(proj) or {}).get("procedures") or [] if x.get("id") == pid), None) or {}
        walk(p.get("body"))
        for h in p.get("handlers") or []:
            walk(h.get("body"))
        return out
    deps = {u["activity"]: sorted(set(callees(u["project"], u["procedure"], {u["procedure"]})) - {u["activity"]}) for u in units}
    done, out = set(), []

    def visit(u, path=()):
        if u["activity"] in done or u["activity"] in path:           # たがいに呼び合う手続きは、番号の順のまま
            return
        for d in deps[u["activity"]]:
            if d in by:
                visit(by[d], path + (u["activity"],))
        done.add(u["activity"])
        out.append(u)
    for u in units:
        visit(u)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["list", "draft", "apply", "terms"])
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--id", help="アクティビティ図の番号（A-02-1）")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--no-edits", action="store_true")
    ap.add_argument("--group", type=int, default=0, help="list: サブエージェント・照合の担当に渡す組を、この枚数ずつ、呼ばれる側から先の順で出す")
    a = ap.parse_args()
    work = work_dir(a.work)
    blocking(work, ("glossary", "systems"))
    if not has_ledger(work):
        sys.exit("structure/systems.json が無い。先に python3 scripts/systems.py draft -w <作業フォルダ> を実行する")
    led = Ledger(work)
    adir = work / "activities"
    if a.cmd == "list":
        for u in ordered(led):
            key = f"{u['project']}.{u['entry']}" if u["kind"] == "entry" else ""
            st = led.structs[u["project"]]
            by = {x["id"]: x for x in st["procedures"]}
            p = by[u["procedure"]]

            def boxes(pid, seen):                                    # 呼ぶ側の図に入れる手続きの分も数える
                more = {c["to"] for c in by[pid]["calls"] if c.get("to") in (st.get("inline") or []) and c["to"] not in seen}
                return by[pid]["boxes"] + sum(boxes(x, seen | more | {pid}) - 1 for x in more)
            off = led.off(u)
            state = f"図にしない（{off}）" if off is not None else "図あり" if (adir / f"{u['activity']}.activity.json").exists() else \
                "直しの一覧あり" if (adir / f"{u['activity']}.edits.json").exists() else "未"
            print(f"{u['activity']}  {u['system']['name']}  {'起動点 ' + key if key else '共通の処理'}  {u['procedure']}（箱の見込み {boxes(p['id'], set())}・実行文 {len(p['exec'])} 行）  {state}")
        todo = [u["activity"] for u in ordered(led) if led.off(u) is None]
        if a.group > 0 and todo:                                     # 組は、呼ばれる側から先の順に切る（前の組の図の名前を、呼び出しの箱に使える）
            print(f"\n組（{a.group} 枚ずつ。上の組から順に渡す）:")
            for k in range(0, len(todo), a.group):
                print(f"  組 {k // a.group + 1}: {'・'.join(todo[k:k + a.group])}")
        return
    tm = T.Terms(work, led.structs)
    bad_terms = validate(load(work / T.FILE), "code_terms") if tm.exists() else []
    if bad_terms:                                                    # 根拠を入れて保存する前に（崩れた行を、黙って落とさない）
        sys.exit(f"{T.FILE} の形が崩れている（{'／'.join(bad_terms[:5])}）。直してから（references/schemas.md の「用語の対応」）")
    for x in tm.fill():                                              # 根拠と場所は、定義の行のコメント・定数の定義から機械が入れる
        print(f"{T.FILE}: {x}")
    if a.cmd == "terms":
        show_terms(led, tm, a.id)
        return
    units = [u for u in ordered(led) if (a.all or u["activity"] == a.id) and led.off(u) is None]
    if not units:
        sys.exit(f"図にする単位「{a.id}」が無い（python3 scripts/activity.py list で一覧を見る。図にしない起動点・図にしない手続きは作らない）" if not a.all else "図にする単位が無い")
    adir.mkdir(exist_ok=True)
    bad = 0
    for u in units:
        aid = u["activity"]
        if a.cmd == "draft":
            d = Draft(led, u)
            d.show()
            cand = T.candidates(tm, d)
            if cand:
                print("プログラムの中の名前（この図の行に出てくるもの）。箱の名前には書かず、業務の言葉で書く。言い換えは、structure/code_terms.json に書いて、どの図でも同じ言葉を使う:")
                for tok, _n, row, where in cand:
                    if row or where:
                        print(f"  {tok} ＝ " + (("（不明: もとの名前のまま「」で書く）" if row["basis"] == "unknown" else f"{row.get('word', '')}〔{T.BASIS.get(row.get('basis'), '')}〕") if row
                                              else f"※未登録  定義: {where[0]} {where[1]}| {where[2][:70]}"))
                rest = [tok for tok, _n, row, where in cand if not row and not where]
                if rest:
                    print(f"  このソースに定義の無い名前（部品・枠組みの名前か、読んでいない定義。これも、箱の名前には書かない）: {'、'.join(rest[:40])}" + (f" ほか {len(rest) - 40}" if len(rest) > 40 else ""))
            continue
        ef = adir / f"{aid}.edits.json"
        edits = {} if a.no_edits else load(ef)
        if edits is None:
            if a.all:
                print(f"{aid}: 直しの一覧（activities/{aid}.edits.json）がまだ無い")
                continue
            sys.exit(f"直しの一覧 {ef} が無い（下書きは python3 scripts/activity.py draft -w {work} --id {aid} で見る）")
        shape = [] if a.no_edits else validate(edits, "activity_edits")
        if not shape and edits and edits.get("activity") not in (None, aid):
            shape = [f"activity が {aid} でない"]
        act, err, d = (None, [], None) if shape else make(led, u, edits)
        if shape or err:
            bad += 1
            print(f"{aid}: 直しの一覧に誤り {len(shape or err)} 件。activity.json は書いていない:")
            for x in shape or err:
                print("  ★ " + x)
            continue
        save(adir / f"{aid}.activity.json", act)
        n = sum(1 for x in act["nodes"] if x["type"] in ("process", "decision", "system", "subprocess"))
        left = sum(1 for x in act["nodes"] if x.get("unnamed"))
        print(f"activities/{aid}.activity.json を書いた:「{act['name']}」箱 {n}・線 {len(act['edges'])}・対象外にした文のまとまり {sum(o['state'] == 'excluded' for o in act['lines'])}"
              + (f"・名前の付いていない箱 {left}" if left else "") + (f"  ※ 箱が {MAX_BOXES} 個を超える。fold・merge で減らす（同じ共通の処理を続けて呼ぶ並びは、draft にまとめの案が出る）" if n > MAX_BOXES else ""))
        for x in d.notes:
            print("  注意: " + x)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
