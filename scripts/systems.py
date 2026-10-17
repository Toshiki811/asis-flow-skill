#!/usr/bin/env python3
"""systems.py — システム台帳（structure/systems.json）と、業務フローとプログラムの紐付け。

  python3 scripts/systems.py draft -w <作業フォルダ>      # 台帳の下書きを書く（用語の正本のシステム＋プログラムのまとまり）。あれば、足りない分だけ足す
  python3 scripts/systems.py entries -w <作業フォルダ>    # 起動点の一覧（アクティビティ図の番号・紐付け・対象外）
  python3 scripts/systems.py links -w <作業フォルダ>      # 紐付けの候補（起動点 × マニュアルの手順。両方に出てくる名前つき）

台帳（LLM が書くのは、ここだけ。形は references/schemas.md の「システム台帳」）:
  systems            システム: 番号（SYS-01…）・名前（用語の正本の名前）・種類・プログラムの資料
  links              紐付け: 起動点（M05.e2）⇔ マニュアルの手順（M02.3.s4）。根拠は3種類だけ
                       named     名指し（evidence の文字が、マニュアルの手順の文と、プログラムの側の名前の両方にある）
                       io        入出力の一致（evidence の帳票名が、手順の文・帳票と、プログラムの読む・作るものの両方にある）
                       inferred  推定（根拠の文が無い。確認事項は機械が作る）
  excluded_entries   図にしない起動点と理由（中身が空・画面の表示だけ など）
  folded_procedures  図にしない手続き（補助）と理由。機械が出す候補（入出力・やり取り・プログラムの終了が無い、共通の処理）から選ぶ。
                     呼び出しは、呼ぶ側の図の、ふつうの箱になる
機械が決めるもの: 起動点（code_struct.py）、アクティビティ図の番号（A-<システムの番号>-<連番>）、紐付いた起動点の箱（S-）を置く場所、
決まった形の確認事項（業務フローに出てこない起動点・推定の紐付け・ボタンの割り当てが見つからないマクロ・たどれない手続き・実行時に決まる名前）。
"""
import argparse
import json
import pathlib
import re

import code_struct as cs
from common import (SYSTEM_KIND, all_pages, schema_enum, file_stem, labels, flow_key, load, load_card, load_flows, load_page, load_systems, manifest, norm, page_json, page_units, save, site, span_text,
                    system_names, try_load, work_dir)

SYS_ID = re.compile(r"^SYS-(\d{2,})$")
BASIS = {"named": "名指し", "io": "入出力の一致", "inferred": "推定"}
GENERIC = {"sheet1", "sheet2", "sheet3", "module1", "module2", "thisworkbook", "thisdocument", "実行", "開始", "処理", "ボタン", "メニュー", "データ", "マクロ", "ツール",
           "macro", "vba", "src", "source", "code", "scripts", "java", "cobol", "cbl", "python"}
ACCESS_OBJECT = {"form": "フォーム", "report": "レポート", "macro": "マクロ", "module": "モジュール"}
KIND_OF_DOC = {"xlsx": "euc_excel", "macrobook": "euc_excel", "legacy": "euc_excel", "access": "euc_access"}
KIND_OF_LANG = {"python": "script", "batch": "script", "powershell": "script", "shell": "script", "sql": "script", "vbscript": "script", "javascript": "script"}


# 図にしない手続き（補助）の理由（決まった語）。業務の計算・判定（金額・税・件数・期限・区分）をする手続きと、画面に見える動きをする手続きは、補助にしない
FOLD_REASONS = tuple(schema_enum("systems", "$defs", "folded", "properties", "reason"))


class Ledger:
    """台帳・構造・紐付けを、引きやすい形にしたもの"""

    def page_lines(self, doc, page):
        """ソースのページの単位（1 行が 1 単位）。読んだページは覚えておく（下書き・図を組むたびに、同じページを読み直さない）"""
        if (doc, page) not in self.page_memo:
            self.page_memo[(doc, page)] = page_units(load_page(page_json(self.work, doc, page)))
        return self.page_memo[(doc, page)]

    def __init__(self, work, m=None):
        self.work = pathlib.Path(work)
        self.m = m or manifest(self.work)
        self.raw = load_systems(self.work)
        self.systems = [x for x in self.raw.get("systems") or [] if isinstance(x, dict) and x.get("name")]
        self.projects = cs.projects(self.m)
        self.structs = cs.load_structs(self.work)
        self.proj_of_doc = {d["doc_id"]: p["id"] for p in self.projects for d in p["docs"]}
        self.sys_of_proj = {}
        for s in self.systems:
            for d in s.get("docs") or []:
                if d in self.proj_of_doc:
                    self.sys_of_proj.setdefault(self.proj_of_doc[d], s)
        # 図にしない手続き（補助）: 「まとまり.手続き」→ 理由。番号は取っておく（ほかの図の番号が変わらない）が、呼び出しの箱の行き先にはしない
        self.folded = {str(x.get("procedure")): str(x.get("reason") or "") for x in self.raw.get("folded_procedures") or [] if isinstance(x, dict)}
        self.units, self.activity = [], {}                           # 図にする単位（番号つき）／(まとまり, "e1" か 手続き) → 番号
        self.page_memo = {}                                          # (資料, ページ) → ページの単位（page_lines で読んだもの）
        for s in self.systems:
            mt, k = SYS_ID.match(str(s.get("id") or "")), 0
            for p in self.projects:
                if self.sys_of_proj.get(p["id"]) is not s or p["id"] not in self.structs:
                    continue
                for u in self.structs[p["id"]]["units"]:
                    k += 1
                    aid = f"A-{mt.group(1)}-{k}" if mt else ""
                    self.units.append(dict(u, project=p["id"], system=s, activity=aid))
                    if u["kind"] == "shared" and f"{p['id']}.{u['procedure']}" in self.folded:
                        continue
                    self.activity[(p["id"], u.get("entry") or u["procedure"])] = aid
                    self.activity.setdefault((p["id"], u["procedure"]), aid)   # 起動点の手続きを、ほかの手続きから呼ぶことがある（呼び出しの箱の行き先）
        self.excluded = {str(x.get("entry")): str(x.get("reason") or "") for x in self.raw.get("excluded_entries") or [] if isinstance(x, dict)}
        self.links = [x for x in self.raw.get("links") or [] if isinstance(x, dict)]

    def off(self, u):
        """図にしない単位なら、その理由（図にしない起動点・図にしない手続き）。図にするなら None"""
        return self.excluded.get(f"{u['project']}.{u['entry']}") if u["kind"] == "entry" else self.folded.get(f"{u['project']}.{u['procedure']}")

    def helpers(self):
        """補助の候補（入出力・やり取り・プログラムの終了が無い、共通の処理）と、台帳で図にしないとした手続き:
        [単位に {"key", "reason"（図にしないものだけ。ほかは None）, "candidate", "boxes", "exec", "sites": [(呼ぶ手続き, 行)]} を足したもの]"""
        out, cand = [], {p: set(cs.helper_candidates(st)) for p, st in self.structs.items()}
        for u in self.units:
            key = f"{u['project']}.{u['procedure']}"
            if u["kind"] != "shared" or not (u["procedure"] in cand.get(u["project"], ()) or key in self.folded):
                continue
            procs = {p["id"]: p for p in self.structs[u["project"]]["procedures"]}
            me = procs[u["procedure"]]
            out.append(dict(u, key=key, reason=self.folded.get(key), candidate=u["procedure"] in cand[u["project"]], boxes=me["boxes"], exec=len(me["exec"]),
                            sites=[(p["id"], c["line"]) for p in procs.values() for c in p["calls"] if c.get("to") == u["procedure"]]))
        return out

    def folded_errors(self):
        """folded_procedures の誤り（無ければ空）"""
        out, seen = [], set()
        units = {f"{u['project']}.{u['procedure']}": u for u in self.units if u["kind"] == "shared"}
        for x in self.raw.get("folded_procedures") or []:
            if not isinstance(x, dict):
                continue
            key = str(x.get("procedure"))
            if key in seen:
                out.append(site(f"「{key}」が、2回書かれている"))
            seen.add(key)
            u = units.get(key)
            if u is None:
                proj, _dot, pid = key.partition(".")
                st = self.structs.get(proj) or {}
                why = "という手続きが無い（書き方は まとまり.モジュール.手続き。python3 scripts/systems.py entries で一覧を見る）"
                if any(p["id"] == pid for p in st.get("procedures") or []):
                    why = "は、起動点の手続き（図にしない起動点は、excluded_entries に書く）" if any(e["procedure"] == pid for e in st.get("entries") or []) else \
                        "は、呼ぶ側の図に入っている（図にする単位でない。その図の中で、畳む・まとめる）" if pid in (st.get("inline") or []) else "は、図にする単位でない"
                out.append(site(f"「{key}」{why}"))
                continue
            procs = {p["id"]: p for p in self.structs[u["project"]]["procedures"]}
            hit = cs.touches(procs, u["procedure"])
            if hit:
                out.append(site(f"「{key}」は、補助の候補でない（{'・'.join(hit[:4])} がある）。図のままにする"))
            more = [q for q in cs.reach(procs, u["procedure"])[1:] if f"{u['project']}.{q}" in units and f"{u['project']}.{q}" not in self.folded]
            if more:
                out.append(site(f"「{key}」が呼ぶ {'・'.join(more[:4])} が、図のまま（そちらも図にしないとするか、「{key}」を図に戻す）"))
        return out

    def excluded_errors(self):
        """図にしない起動点の理由が、中身と合っているか（決まり 6.4）: 中身が空 → 実行文が無い／画面の表示の制御だけ → 入出力・やり取り・終了が無い／
        ほかの起動点と同じ処理を呼ぶだけ → 実行文が、図にする起動点からたどれる手続きの呼び出しだけ。「開発・試験用」は、中身からは決まらないので見ない"""
        out, all_e = [], self.entries()
        for e in all_e:
            why = self.excluded.get(e["key"])
            procs = {p["id"]: p for p in (e["struct"] or {}).get("procedures") or []}
            p = procs.get(e["procedure"])
            if why is None or p is None:
                continue
            if why == "中身が空" and p.get("exec"):
                out.append(site(f"「{e['key']}」の理由は「中身が空」なのに、実行文が {len(p['exec'])} 行ある（図にするか、理由を直す）"))
            elif why == "画面の表示の制御だけ" and cs.touches(procs, p["id"]):
                out.append(site(f"「{e['key']}」の理由は「画面の表示の制御だけ」なのに、{'・'.join(cs.touches(procs, p['id'])[:3])} がある（図にする）"))
            elif why == "ほかの起動点と同じ処理を呼ぶだけ":
                drawn = {q for x in all_e if x["project"] == e["project"] and x["key"] != e["key"] and x["key"] not in self.excluded for q in cs.reach(procs, x["procedure"])}
                calls = {c.get("to") for c in p.get("calls") or []}
                if not calls or None in calls or not calls <= drawn or len(p.get("exec") or []) > len(p.get("calls") or []):
                    out.append(site(f"「{e['key']}」の理由は「ほかの起動点と同じ処理を呼ぶだけ」なのに、図にする起動点の処理を呼ぶ文のほかに、実行文がある"
                               "か、呼ぶ先が図にする起動点の処理でない（図にする）"))
        return out

    def entries(self):
        """起動点の一覧: [{"key", "project", "id", "label", "trigger", "procedure", "system", "activity", "excluded", "links"}]"""
        out = []
        for p in self.projects:
            st = self.structs.get(p["id"])
            for e in (st or {}).get("entries") or []:
                key = f"{p['id']}.{e['id']}"
                out.append({"key": key, "project": p["id"], "id": e["id"], "label": e["label"], "trigger": e["trigger"], "procedure": e["procedure"],
                            "system": self.sys_of_proj.get(p["id"]), "activity": self.activity.get((p["id"], e["id"]), ""), "excluded": self.excluded.get(key),
                            "links": [x for x in self.links if x.get("entry") == key], "struct": st})
        return out

    def entry(self, key):
        return next((e for e in self.entries() if e["key"] == key), None)

    def title(self, activity, fallback):
        """アクティビティ図の名前（まだ無ければ fallback）"""
        a = try_load(self.work / "activities" / f"{activity}.activity.json")[0] if activity else None
        return str(a.get("name")) if isinstance(a, dict) and a.get("name") else fallback

    # ---- 根拠の照合 ----

    def code_names(self, e):
        """起動点の、プログラムの側の名前（名指し・入出力の一致の根拠に使える文字）: {"named": [名前], "io": [名前]}"""
        st = e["struct"]
        procs = {p["id"]: p for p in st["procedures"]}
        reach = cs.reach(procs, e["procedure"])
        named = re.findall(r"「([^」]+)」", e["label"]) + [e["procedure"].split(".")[-1]]
        named += [x["text"] for pid in reach for x in procs[pid].get("interactions") or [] if x.get("text")]
        if sum(1 for x in st["entries"] if f"{e['project']}.{x['id']}" not in self.excluded) == 1:   # 起動点が1つだけのプログラムは、ファイルの名前が、その起動点を名指す
            named += [x for p in self.projects if p["id"] == e["project"] for d in p["docs"] for x in (file_stem(d["name"]), d["name"])]
        named = [x for x in named if norm(x).lower() not in GENERIC]
        io = [x["name"] for pid in reach for x in procs[pid].get("io") or [] if not x["name"].startswith("（")]
        io += [file_stem(x) for x in io if file_stem(x) != x]
        return {"named": list(dict.fromkeys(x for x in named if x)), "io": list(dict.fromkeys(io))}

    def step(self, key):
        """手順のキー（M02.3.s4）→ (資料, page, 項目, カードの手順) か None"""
        mt = re.fullmatch(r"([^.]+)\.(\d+)\.(s\d+)", str(key or ""))
        if not mt:
            return None
        doc, page, item = mt.group(1), int(mt.group(2)), mt.group(3)
        it = next((s for s in (load_card(self.work, doc, page) or {}).get("steps") or [] if isinstance(s, dict) and s.get("id") == item), None)
        return (doc, page, item, it) if it else None

    def step_text(self, doc, page, it):
        pj = load_page(page_json(self.work, doc, page))
        line = span_text(page_units(pj), it.get("lines")) or ""
        return "\n".join([str(it.get("action", "")), str(it.get("quote", "")), line, str(it.get("system", ""))] + list(it.get("inputs") or []) + list(it.get("outputs") or []))

    def link_errors(self, x):
        """紐付け1件の誤り（無ければ空）"""
        e, st = self.entry(x.get("entry")), self.step(x.get("step"))
        if e is None:
            return [site(f"起動点「{x.get('entry')}」が無い（python3 scripts/systems.py entries で一覧を見る）")]
        if st is None:
            return [site(f"手順「{x.get('step')}」が無い（書き方は 資料.ページ.項目。例 M02.3.s4。カードの手順だけを指せる）")]
        if e["excluded"] is not None:
            return [site(f"起動点 {e['key']} は、図にしない（excluded_entries）としているのに、紐付けがある")]
        if e["system"] is None:
            return [site(f"起動点 {e['key']} のプログラムの資料が、どのシステムにも入っていない（systems[].docs）")]
        basis, ev = x.get("basis"), norm(x.get("evidence", ""))
        if basis not in BASIS:
            return [site(f"basis「{basis}」は決まった語でない（named 名指し／io 入出力の一致／inferred 推定）")]
        if basis == "inferred":
            return []
        if len(ev) < 2:
            return [site("evidence（両方に出てくる文字）が空か、短すぎる")]
        names = self.code_names(e)[basis]
        if not any(ev in norm(n) or (len(norm(n)) >= 2 and norm(n) in ev) for n in names):
            return [site(f"evidence「{x.get('evidence')}」が、プログラムの側の{'名前（ボタン・手続き・ファイル・メッセージ）' if basis == 'named' else '入出力（シート・ファイル・テーブル）'}に無い"
                    f"（あるのは {'・'.join(names[:8]) or 'なし'}）。根拠が無ければ basis を inferred にする")]
        if ev not in norm(self.step_text(*st[:2], st[3])):
            return [site(f"evidence「{x.get('evidence')}」が、マニュアルの手順 {x.get('step')} の文・帳票に無い。根拠が無ければ basis を inferred にする")]
        return []

    def good_links(self):
        return [x for x in self.links if not self.link_errors(x)]


# ---------- 下書き・一覧・候補 ----------

def draft(work, led):
    """台帳の下書き: いまの台帳に、足りないシステム（用語の正本のシステム・どのシステムにも入っていないプログラムのまとまり）を足す"""
    names, used = system_names(work), set()
    g = load(work / "glossary.json", {}) or {}
    terms = [t["name"] for t in g.get("terms") or [] if isinstance(t, dict) and t.get("kind") == "system" and t.get("name")]
    systems = [dict(s) for s in led.systems]
    for s in systems:
        mt = SYS_ID.match(str(s.get("id") or ""))
        if mt:
            used.add(int(mt.group(1)))

    def new_id():
        k = 1
        while k in used:
            k += 1
        used.add(k)
        return f"SYS-{k:02d}"
    have = {norm(s["name"]) for s in systems}
    notes = []
    for t in terms:
        if norm(t) not in have:
            systems.append({"id": new_id(), "name": t, "kind": "", "docs": []})
            have.add(norm(t))
    taken = {d for s in systems for d in s.get("docs") or []}
    for p in led.projects:
        ids = [d["doc_id"] for d in p["docs"]]
        if set(ids) & taken:
            for s in systems:                                        # まとまりの一部だけが入っているシステムには、残りも入れる
                if set(ids) & set(s.get("docs") or []):
                    s["docs"] += [x for x in ids if x not in s["docs"]]
            continue
        d0 = p["docs"][0]
        acc = (led.structs.get(p["id"]) or {}).get("access")
        cand = [file_stem(d["name"]) for d in p["docs"]] + ([d0["folder"][-1]] if d0.get("folder") and d0["kind"] == "code" else [])
        if acc:                                                      # Access を書き出したフォルダ: 元のファイルの名前（無ければ、フォルダの名前）
            cand = [file_stem(acc["source"])] if acc.get("source") else [re.sub(r"_export$", "", d0["folder"][-1])] if d0.get("folder") else cand
        cand = [c for c in cand if norm(c).lower() not in GENERIC] or [file_stem(d0["name"])]
        if not acc and d0["kind"] == "code" and len(p["docs"]) > 1:   # ファイルが2つ以上のまとまり: 先頭のファイルの名前だけでは、ほかのファイルの処理まで、その名前に見える
            own = [c for c in cand if c not in [file_stem(d["name"]) for d in p["docs"]]]
            cand = (own + cand) if own else [f"{cand[0]} ほか {len(p['docs']) - 1} ファイル"] + cand
        hit = next((s for s in systems for c in cand if len(norm(c)) >= 2 and (norm(c) in norm(s["name"]) or norm(s["name"]) in norm(c)
                                                                               or any(norm(c) == k for k, v in names.items() if v == s["name"]))), None)
        kind = "euc_access" if acc else KIND_OF_DOC.get(d0["kind"]) or KIND_OF_LANG.get(p["language"], "")
        if hit is None:
            name = cand[0]
            hit = {"id": new_id(), "name": name, "kind": kind, "docs": []}
            systems.append(hit)
            if norm(name) not in names:
                notes.append(f"「{name}」（{'・'.join(ids)}）は、用語の正本（glossary.json）にシステムとして無い。マニュアルでの呼び方を確かめ、"
                             f"あればその名前に直して docs を移す。無ければ、この名前を glossary.json に kind: system で足す")
        hit["docs"] = list(hit.get("docs") or []) + ids
        hit["kind"] = hit.get("kind") or kind
    out = dict(led.raw)
    out.update(systems=systems, links=led.raw.get("links") or [], excluded_entries=led.raw.get("excluded_entries") or [], questions=led.raw.get("questions") or [])
    if led.raw.get("folded_procedures"):
        out["folded_procedures"] = led.raw["folded_procedures"]
    return out, notes


def show_entries(led, out=print):
    for s in led.systems:
        out(f"== {s.get('id')} {s['name']}（{SYSTEM_KIND.get(s.get('kind'), s.get('kind') or '種類は未設定')}）  プログラムの資料: {'・'.join(s.get('docs') or []) or 'なし'}")
        for e in led.entries():
            if e["system"] is not s:
                continue
            p = next(x for x in e["struct"]["procedures"] if x["id"] == e["procedure"])
            state = f"図にしない（{e['excluded']}）" if e["excluded"] is not None else \
                "／".join(f"{x.get('step')}（{BASIS.get(x.get('basis'), x.get('basis'))}）" for x in e["links"]) or "紐付けなし（業務フローに出てこない）"
            out(f"  {e['key']}  {e['activity']}  {e['label']}  → {e['procedure']}（箱 {p['boxes']}・実行文 {len(p['exec'])} 行）  {state}")
        helpers = {h["key"]: h for h in led.helpers()}
        for u in led.units:
            if u["system"] is s and u["kind"] == "shared":
                h = helpers.get(f"{u['project']}.{u['procedure']}")
                state = "" if h is None else f"  図にしない（{h['reason']}）" if h["reason"] is not None else \
                    f"  ★補助の候補（入出力・やり取りが無い。箱 {h['boxes']}・呼ぶ所 {len(h['sites'])}）: 図にしないなら folded_procedures に「{h['key']}」と理由を書く"
                out(f"  （共通）  {u['activity']}  {u['procedure']}{state}")
    loose = [e for e in led.entries() if e["system"] is None]
    if loose:
        out("== どのシステムにも入っていないプログラム（systems[].docs に資料を入れる）")
        for e in loose:
            out(f"  {e['key']}  {e['label']}  → {e['procedure']}")


def candidates(led, out=print):
    """紐付けの候補: 起動点ごとに、プログラムの側の名前が出てくるマニュアルの手順"""
    steps = []
    for d in led.m["docs"]:
        if d.get("duplicate_of"):
            continue
        for p in d.get("page_list") or []:
            c = load_card(led.work, d["doc_id"], p["page"]) or {}
            for s in c.get("steps") or []:
                if isinstance(s, dict) and s.get("id"):
                    steps.append((d["doc_id"], p["page"], s, norm(led.step_text(d["doc_id"], p["page"], s))))
    names = system_names(led.work)
    for e in led.entries():
        if e["excluded"] is not None:
            continue
        cn = led.code_names(e)
        sysname = e["system"]["name"] if e["system"] else ""
        out(f"== {e['key']}  {e['activity']}  {sysname}  {e['label']}  → {e['procedure']}")
        out(f"   プログラムの側の名前: {'・'.join(cn['named'][:10]) or 'なし'}／入出力: {'・'.join(cn['io'][:12]) or 'なし'}")
        rows = []
        for doc, page, s, text in steps:
            hit_n = [n for n in cn["named"] if len(norm(n)) >= 2 and norm(n).lower() not in GENERIC and norm(n) in text]
            hit_io = [n for n in cn["io"] if len(norm(n)) >= 2 and norm(n).lower() not in GENERIC and norm(n) in text]
            same_sys = bool(sysname) and (norm(sysname) in text or any(k in text for k, v in names.items() if v == sysname))
            score = 3 * len(hit_n) + 2 * len(hit_io) + (2 if same_sys else 0)
            if score:
                rows.append((-score, doc, page, s, hit_n, hit_io, same_sys))
        for _sc, doc, page, s, hit_n, hit_io, same_sys in sorted(rows, key=lambda r: (r[0], r[1], r[2]))[:6]:
            why = [f"named「{x}」" for x in hit_n[:3]] + [f"io「{x}」" for x in hit_io[:3]] + (["手順がこのシステムを名指ししている（名前だけでは、どの起動点かは決まらない）"] if same_sys else [])
            out(f"   候補 {doc}.{page}.{s['id']}  {str(s.get('actor') or '担当なし')}: {str(s.get('action', ''))[:40]}  ← {'、'.join(why)}")
        if not rows:
            out("   候補なし（マニュアルに、このプログラムの名前・入出力が出てこない。紐付けなければ「業務フローに出てこない起動点」として確認事項になる）")


# ---------- 紐付いた起動点の箱を、フローに置く ----------

def placed(fl):
    return [n for n in fl.get("nodes") or [] if isinstance(n, dict) and n.get("activity")]


def unplace(fl):
    """機械が置いた起動点の箱を外し、線と確認事項を、置く前の形に戻す"""
    for s in placed(fl):
        sid = s["id"]
        ins = [e for e in fl["edges"] if e.get("to") == sid]
        outs = [e for e in fl["edges"] if e.get("from") == sid]
        fl["edges"] = [e for e in fl["edges"] if not any(e is i for i in ins) and (ins or not any(e is o for o in outs))]
        for o in outs if ins else []:                                # 箱 → 行き先 の線を、その場所のまま 作業 → 行き先 に戻す（線の並びを変えない）
            o["from"] = ins[0]["from"]
            o.pop("kind", None)
        for q in fl.get("questions") or []:
            if isinstance(q.get("edge"), list) and len(q["edge"]) == 2 and q["edge"][0] == sid and ins:
                q["edge"] = [ins[0]["from"], q["edge"][1]]
        fl["nodes"] = [n for n in fl["nodes"] if n is not s]
    fl["questions"] = [q for q in fl.get("questions") or [] if not str(q.get("id", "")).startswith("m-sys-")]
    used = {n.get("lane") for n in fl["nodes"]}
    fl["lanes"] = [l for l in fl["lanes"] if l["id"] in used or l.get("kind") != "system"]


def entry_source(led, e):
    """起動点の箱の出典: 手続きの先頭の行（item に起動点の id）"""
    st = e["struct"]
    p = next(x for x in st["procedures"] if x["id"] == e["procedure"])
    src = cs.cite(st, p["module"], p["lines"][0], p["lines"][0])[0]
    text = span_text(page_units(load_page(page_json(led.work, src["doc_id"], src["page"]))), src["lines"]) or ""
    return dict(src, item=e["id"], quote=text.strip(), from_image=False)


def place_entries(fl, led):
    """紐付いた起動点ごとに、システム処理の箱（S-）を、その手順を引く作業のすぐあとに置く。前に置いた箱は外してから置き直すので、何度実行しても同じ。
    線は、作業 → 箱 → もとの行き先（システム連携の線）。推定の紐付けは、作業 → 箱 の線を推定にして、確認事項を付ける。変わったら True"""
    from renumber import renumber
    before = json.dumps(fl, ensure_ascii=False, sort_keys=True)
    unplace(fl)
    lane = {l["id"]: l for l in fl["lanes"]}
    todo = []
    for x in led.good_links():
        e, st = led.entry(x["entry"]), led.step(x["step"])
        host = next((n for n in fl["nodes"] if n.get("type") in ("process", "system") and any(
            isinstance(s, dict) and (s.get("doc_id"), s.get("page"), s.get("item")) == st[:3] for s in n.get("sources") or [])), None)
        if host is not None and e["activity"]:
            todo.append((host, e, x))
    for k, (host, e, x) in enumerate(todo, 1):
        name = e["system"]["name"]
        lid = next((l["id"] for l in fl["lanes"] if norm(l["name"]) == norm(name)), None)
        if lid is None:
            lid = f"L{max([int(l['id'][1:]) for l in fl['lanes'] if re.fullmatch(r'L[0-9]+', l['id'])] or [0]) + 1}"
            fl["lanes"].append({"id": lid, "name": name, "kind": "system"})
        sid = f"S-new{k}"
        box = {"id": sid, "type": "system", "lane": lid, "label": led.title(e["activity"], e["procedure"].split(".")[-1]), "activity": e["activity"],
               "sources": [entry_source(led, e)]}
        tail = host
        while True:                                                  # 同じ作業に2つ以上の起動点が紐付くときは、置いた順につなぐ
            nxt = next((n for n in placed(fl) for ed in fl["edges"] if ed["from"] == tail["id"] and ed["to"] == n["id"] and n.get("activity")), None)
            if nxt is None:
                break
            tail = nxt
        for ed in [ed for ed in fl["edges"] if ed["from"] == tail["id"]]:
            ed["from"] = sid
            ed.setdefault("kind", "system")
        for q in fl.get("questions") or []:
            if isinstance(q.get("edge"), list) and len(q["edge"]) == 2 and q["edge"][0] == tail["id"]:
                q["edge"] = [sid, q["edge"][1]]
        link = {"from": tail["id"], "to": sid, "kind": "system"}
        fl["nodes"].insert(fl["nodes"].index(tail) + 1, box)
        if x.get("basis") == "inferred":
            link["inferred"] = True
            fl.setdefault("questions", []).append({
                "id": f"m-sys-{e['key'].replace('.', '-')}", "severity": "low", "category": "システム", "effect": "システム・帳票", "edge": [tail["id"], sid],   # 札の付け足し（決まり 5.4）
                "question": f"「{host.get('label', '')[:40]}」では、{name}の「{e['label'][:40]}」を使っていますか。原典には、このプログラムを使うとは書かれていません。",
                "assumption": "作業の内容と、プログラムの入出力から、ここで使うと仮に置いた",
                "sources": [{"doc_id": s["doc_id"], "page": s["page"]} for s in host.get("sources") or [] if isinstance(s, dict)][:1],
                "ask": lane.get(host.get("lane"), {}).get("name", "")})
        fl["edges"].append(link)
    if todo:
        if not (led.work / "review" / f"{flow_key(fl)}.review.json").exists():
            renumber(fl)
        else:                                                        # 採否表があるフローは、ほかの番号を変えない。空いている番号を付ける
            free = max([int(n["id"][2:]) for n in fl["nodes"] if re.fullmatch(r"S-[0-9]{3}", str(n.get("id")))] or [0])
            ren = {}
            for n in fl["nodes"]:
                if str(n["id"]).startswith("S-new"):
                    free += 1
                    ren[n["id"]] = f"S-{free:03d}"
                    n["id"] = ren[n["id"]]
            for ed in fl["edges"]:
                ed["from"], ed["to"] = ren.get(ed["from"], ed["from"]), ren.get(ed["to"], ed["to"])
            for q in fl.get("questions") or []:
                if isinstance(q.get("edge"), list):
                    q["edge"] = [ren.get(v, v) for v in q["edge"]]
    return json.dumps(fl, ensure_ascii=False, sort_keys=True) != before


def has_ledger(work):
    return (pathlib.Path(work) / "structure" / "systems.json").exists()


# ---------- 決まった形の確認事項 ----------

# 答えで、図のどこが変わるか（元の id の頭 → effect）。名前・呼び方だけが変わる問い（つづりから読んだ言葉・読めない名前）は、作らない（一覧の「用語の対応」に、色つきで出る）
MQ_EFFECT = {"unlinked": "システム・帳票", "unplaced": "システム・帳票", "dynamic": "システム・帳票", "manual": "フローの範囲", "unreached": "フローの範囲", "vba-unread": "フローの範囲",
             "skipped": "箱", "partial": "箱", "export-error": "箱", "export-failed": "箱", "direct": "箱", "unread": "箱", "locked": "箱"}

def machine_questions(work, m=None):
    """台帳とプログラムの構造から、機械が作る確認事項: [(元の id, 確認事項)]。フローに置くもの（推定の紐付け）は place_entries が作る"""
    work = pathlib.Path(work)
    if not has_ledger(work):
        return []
    led, out = Ledger(work, m), []
    title = labels(led.m, work).title                                 # 資料は題名で書く
    flows = load_flows(work, [])
    boxed = {n.get("activity") for fl in flows for n in placed(fl)}

    def q(local, sev, text, assume, docs, ask=""):
        out.append((local, {"severity": sev, "category": "システム", "effect": MQ_EFFECT[local.split(":")[0]], "question": text, "assumption": assume,
                            "sources": [{"doc_id": d} for d in docs], "ask": ask}))
    no_manual = not any(True for _x in all_pages(led.m))             # マニュアルが入力に無い（プログラムのソースだけ）: 紐付けの問いは、システムごとに1つにまとめる
    if no_manual:
        for s in led.systems:
            mine = [e for e in led.entries() if e["system"] is s and e["excluded"] is None]
            if mine:
                q(f"unlinked:{s['id']}", "mid", f"{s['name']}の起動点 {len(mine)} 件（{mine[0]['activity']}〜{mine[-1]['activity']}）は、どの業務の、どの作業で使っていますか。"
                  "マニュアルが入力に無いので、業務フローとの紐付けはしていません（この確認事項の回答の欄に、アクティビティ図の番号ごとに、使う業務と作業を書いてください）。",
                  "業務フローは作っていない（アクティビティ図だけを作った）", sorted({e["project"] for e in mine}))
    for e in led.entries():
        if e["system"] is None or e["excluded"] is not None:
            continue
        name, docs = e["system"]["name"], [e["project"]]
        how = (f"どの操作で実行していますか（ボタン・ショートカットの割り当てが、ファイルの中に"
               f"{'見つかりません' if '見つからない' in e['label'] or '（マクロ）' in e['procedure'] else '読めていません'}）") if e["trigger"] == "manual" else ""
        if no_manual:
            pass
        elif not led.good_links() or not any(x.get("entry") == e["key"] for x in led.good_links()):
            # 同じ起動点に、紐付けと実行のしかたの 2 つを聞くときは、1 つの問いにまとめる（決まり 5.6）
            q(f"unlinked:{e['key']}", "mid", f"{name}の「{e['label']}」（{e['activity']}）は、どの業務の、どの作業で使っていますか。マニュアルには、このプログラムを使う作業が見つかりません。"
              + (f"あわせて、{how}。" if how else ""), "業務フローには置いていない（アクティビティ図だけを作った）" + ("。マクロの一覧から実行すると仮に置いた" if how else ""), docs)
            how = ""
        elif flows and e["activity"] not in boxed:
            q(f"unplaced:{e['key']}", "mid", f"{name}の「{e['label']}」（{e['activity']}）を使う作業が、どのフローにも描かれていません（紐付けた手順を、フローに入れていない）。"
              + (f"あわせて、{how}。" if how else ""), "業務フローには置いていない" + ("。マクロの一覧から実行すると仮に置いた" if how else ""), docs)
            how = ""
        if how:
            q(f"manual:{e['key']}", "low", f"{name}のマクロ「{e['procedure'].split('.')[0 if e['procedure'].endswith('（マクロ）') else -1]}」は、{how}。",
              "マクロの一覧から実行すると仮に置いた", docs)
    for p in led.projects:
        st, s = led.structs.get(p["id"]), led.sys_of_proj.get(p["id"])
        if not st or s is None:
            continue
        if st["unreached"]:
            kind_of = {m["name"]: m["kind"] for m in st["modules"]}
            refd = {r["procedure"]: r["by"] for r in st.get("referenced") or []}
            klass = [x for x in st["unreached"] if kind_of.get(x.rsplit(".", 1)[0]) == "class"]     # オブジェクト越しに呼ばれる（機械では、呼び出しを追えない）
            plain = [x for x in st["unreached"] if x not in klass and x not in refd]
            text = ""
            if plain:
                text += (f"{s['name']}の次の手続きは、どの起動点からも呼ばれていません。図にする処理（どこかから動かしているもの）は、ありますか: {'、'.join(x.split('.')[-1] for x in plain[:12])}"
                         + (f" ほか {len(plain) - 12} 件" if len(plain) > 12 else "") + "。")
            if refd:
                text += (f"次の関数は、クエリや式の中で使われています（起動点からは、たどれません）: "
                         + "、".join(f"{k.split('.')[-1]}（{v}）" for k, v in list(refd.items())[:8]) + (f" ほか {len(refd) - 8} 件" if len(refd) > 8 else "") + "。")
            if klass:
                mods = list(dict.fromkeys(x.rsplit(".", 1)[0] for x in klass))
                text += (f"クラス モジュールの手続き {len(klass)} 件（{'、'.join(mods[:6])}{' ほか' if len(mods) > 6 else ''}）は、オブジェクト越しに呼ばれるので、"
                         "呼び出しを機械で追えていません。図にする処理があれば、教えてください。")
            q(f"unreached:{p['id']}", "low", text.strip(), "図にしていない", [p["id"]])
        dyn = [(x["procedure"], y) for x in ({"procedure": pr["id"], "dynamic": pr.get("dynamic") or []} for pr in st["procedures"]) for y in x["dynamic"]]
        if dyn:
            q(f"dynamic:{p['id']}", "mid", f"{s['name']}で、実行するときに決まる名前（ファイル・シート・テーブル）が {len(dyn)} か所あります。実際には何を指していますか: "
              + "、".join(f"{a.split('.')[-1]} {b['line']} 行（{b['what']}）" for a, b in dyn[:8]), "図では「（実行時に決まる）」と書いた", [p["id"]])
        rd = st.get("read") or {}
        if rd.get("skip"):
            q(f"skipped:{p['id']}", "mid", f"{s['name']}のプログラム（{'・'.join(title(x) for x in st.get('docs') or []) or s['name']}）は、読んでいません（{rd['skip']}）。図にする処理があれば、範囲を教えてください。",
              "図にしていない", [p["id"]])
        elif rd.get("lines") and rd["lines"] < rd["total_lines"] * 0.8:
            q(f"partial:{p['id']}", "low", f"{s['name']}のプログラムは、全 {rd['total_lines']} 行のうち {rd['lines']} 行（手続き {len(st['procedures'])}）を読みました"
              "（起動点からたどれる範囲）。ほかに、図にする処理はありますか。", "読んだ範囲だけを図にした", [p["id"]])
        acc = st.get("access") or {}
        if acc.get("error"):
            q(f"export-error:{p['id']}", "high", f"{s['name']}の書き出しが、途中で止まっています（{acc['error']}）。書き出しをやり直せますか。", "書き出せた範囲だけを読んだ", [p["id"]])
        direct = acc.get("via") == "direct"                          # Access のファイルを直接読んだ（書き出したテキストが無い）
        if acc.get("failed"):
            q(f"export-failed:{p['id']}", "mid", f"{s['name']}の次のものは、" + ("ファイルを直接読むやり方では、読めませんでした。Access のあるパソコンで、テキストに書き出してもらえますか: "
                                                                               if direct else "Access から書き出せませんでした。中身を確かめる方法はありますか: ")
              + "、".join(f"{ACCESS_OBJECT.get(x['kind'], x['kind'])}「{x['name']}」" for x in acc["failed"][:12]), "読んでいない（図にしていない）", [p["id"]])
        if direct:
            q(f"direct:{p['id']}", "mid", f"{s['name']}は、Access のファイルを直接読みました（書き出したテキストが無いため）。次の点は、読めていないか、確かめられていません。"
              "Access のあるパソコンで、テキストに書き出してもらえますか: " + "／".join((acc.get("notes") or [])[:10]), "読めた範囲だけを図にした", [p["id"]])
        if acc.get("unread"):
            q(f"unread:{p['id']}", "mid", f"{s['name']}の次のものは、{'直接読んだ' if direct else '書き出した'}テキストから読み取れませんでした: "
              + "、".join(f"{x['what']}（{x['where']}。{x['why']}）" for x in acc["unread"][:8]), "図にしていない", [p["id"]])
    for d in led.m["docs"]:
        if d.get("vba_locked") and not d.get("code"):
            q(f"locked:{d['doc_id']}", "high", f"資料「{title(d['doc_id'])}」のマクロには、閲覧のロックが掛かっています。中身を読んでよいですか。", "マクロは読んでいない", [d["doc_id"]])
        if d.get("vba_unread") and d.get("code"):
            q(f"vba-unread:{d['doc_id']}", "mid", f"資料「{title(d['doc_id'])}」は、次のものを読めていません: {'／'.join(d['vba_unread'])}。"
              "シートのボタン・図形に割り当ててあるマクロがあれば、教えてください（新形式 xlsm で保存し直したものをもらえば、読み取れます）。",
              "ボタンの割り当ては無いものとして、起動点を並べた", [d["doc_id"]])
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft", "entries", "links"])
    ap.add_argument("-w", "--work", required=True)
    a = ap.parse_args()
    work = work_dir(a.work)
    led = Ledger(work)
    if a.cmd == "draft":
        if not led.structs and led.projects:
            raise SystemExit("プログラムの構造がまだ無い。先に python3 scripts/code_struct.py -w <作業フォルダ> を実行する")
        data, notes = draft(work, led)
        save(work / "structure" / "systems.json", data)
        print(f"structure/systems.json を書いた: システム {len(data['systems'])}（プログラムのあるもの {sum(bool(s.get('docs')) for s in data['systems'])}）")
        for n in notes:
            print("  ※ " + n)
        print("  種類（kind）の空いているシステムを埋め、python3 scripts/systems.py links で候補を見て、links と excluded_entries を書く（references/code-guide.md）")
        n = sum(1 for h in Ledger(work).helpers() if h["reason"] is None)
        if n:
            print(f"  補助の候補が {n} 個ある（python3 scripts/systems.py entries の ★）。図にしないものを、理由つきで folded_procedures に書く（{'／'.join(FOLD_REASONS)}）")
        return
    if not has_ledger(work):
        raise SystemExit("structure/systems.json が無い。先に python3 scripts/systems.py draft -w <作業フォルダ> を実行する")
    if a.cmd == "entries":
        show_entries(led)
    else:
        candidates(led)


if __name__ == "__main__":
    main()
