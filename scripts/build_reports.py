#!/usr/bin/env python3
"""build_reports.py — 確認事項一覧と出典対応表を書き出す。

  python3 scripts/build_reports.py -w <作業フォルダ>
  python3 scripts/build_reports.py -w <作業フォルダ> --keep <回答の入ったブック.xlsx>   # お客様から戻ってきた確認事項一覧の回答を取り込んでから書き出す

  out/確認事項一覧.md / .xlsx   全工程の確認事項（番号・重要度・種別・答えで変わる所・対象のフローとノード・質問・仮置きした内容・出典・聞く相手）。
                              xlsx には回答を書き込む列がある。番号は図の中のバッジ（Q-012）と同じ。
                              回答・回答者・回答日（様式で足した、人が埋める欄も）は、書き出す前に、いまのブックから Q 番号で answers.json に取り込み、書き出しで戻す
                              xlsx の列の順・見出し・体裁・人が埋める欄は、テンプレート（assets/確認事項一覧_テンプレート.xlsx）が決める
  out/出典対応表.md / .csv     ノード → 出典ページ（フローごと）と、ページ → ノード（資料ごと。使っていないページは理由つき）、
                              フローに入れなかった手順と理由
  out/階層と業務分類.md        業務の木・結び付く資料と役割・資料どうしの関係・業務どうしの「続く」（根拠つき）
  out/原典照合レビュー_採否表.md  review/*.review.json（指摘と採否）をまとめたもの。支持率つき（レビューをしたときだけ）
ページは、資料に刷られたページ番号で書く（Excel はシート名とセル）。ファイルの何枚目かは出さない。
確認事項の番号は questions.json で引き継ぐので、作り直しても変わらない。
"""
import argparse
import copy
import csv
import io
import pathlib

from common import (NODE_LABEL, HIERARCHY_LABEL, NON_CONTENT_TYPES, Answers, BadJson, CardItems, all_pages, answer_gaps, collect_questions, item_reasons, labels,
                    load, load_card, load_flows, manifest, node_details, q_targets, src_label, tree_nodes, try_load, unused_reasons,
                    why_unused, work_dir)
from common import ASSETS, OUTPUT_FONT, REL_LABEL as REL, ROLE_LABEL as ROLE, SEV_LABEL as SEV
from validate import blocking, validate

SEV_ORDER = {"high": 0, "mid": 1, "low": 2}
SEV_NOTE = {"high": "フローの範囲や、複数の箱・線が変わる", "mid": "箱 1 つ・線 1 本・担当 1 つが変わる（仮置きで描けている）", "low": "図への付け足し（帳票・システムの札・開始条件・終了条件）が変わる"}
KIND = NODE_LABEL
NUMBERING = {"printed": "資料に刷られた番号", "count": "番号の印字なし（先頭から数えた番号）", "sheet": "シート名"}
PTYPE = {"cover": "表紙", "toc": "目次", "overview": "概要", "flow_diagram": "フロー図", "procedure": "手順", "screen": "画面",
         "form": "帳票", "table": "表", "reference": "参考", "code": "マクロ", "blank": "白紙", "other": "その他"}
REF_REL = {"describes": "書いている", "details": "詳しくしている"}
TAX_BASIS = {"toc": "目次の業務フロー一覧に準拠", "llm": "LLM が分類", "mixed": "目次に準拠し、無い部分は LLM が分類"}
ORIGIN = {"inventory": "入力資料", "hierarchy": "階層", "glossary": "用語", "taxonomy": "業務分類", "systems": "システム台帳"}   # フロー・アクティビティ図の問いは、その番号と名前
ITEM_KIND = {"s": "手順", "d": "分岐", "g": "図の箱", "n": "補足"}             # カードの項目の種類（項目の番号は、作業の記録にだけ書く）
Q_HEAD = ["番号", "重要度", "種別", "答えで変わる所", "対象", "ノード・線", "質問", "仮置きした内容", "出典", "聞く相手"]
Q_BOOK, Q_SHEET, Q_TEMPLATE = "確認事項一覧.xlsx", "確認事項", "確認事項一覧_テンプレート.xlsx"
# 欄の決まり（テンプレートの印）と、既定のテンプレートの体裁（回答・回答者・回答日は、人が埋める欄）
Q_SPEC = {"sheets": {Q_SHEET: Q_HEAD}, "required": {Q_SHEET: ["番号", "質問"]}, "keys": {Q_SHEET: ["番号"]}}   # 人が埋める欄は、番号の行に戻す
Q_LOOK = {"font": OUTPUT_FONT, "size": 11, "head_font": dict(b=True), "border": False, "head_align": False, "head_fill": "DAE8FC", "order": [Q_SHEET], "manual": {Q_SHEET: ["回答", "回答者", "回答日"]},
          "widths": {Q_SHEET: [9, 8, 14, 16, 20, 16, 60, 40, 22, 14, 50, 12, 12]}}


def cell(t):
    return str(t if t is not None else "").replace("|", "｜").replace("\n", " ")


def row(values):
    """Markdown の表の1行"""
    return "| " + " | ".join(cell(v) for v in values) + " |"


def table(head):
    """Markdown の表の見出しと区切りの2行"""
    return ["| " + " | ".join(head) + " |", "| " + " | ".join("---" for _ in head) + " |"]


def write_md(path, lines):
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def target(q):
    if q_targets(q):
        return "・".join(q_targets(q))
    if q.get("edge"):
        return "→".join(q["edge"])
    return ""


class Reports:
    """報告書の材料（作業フォルダから1回だけ読む）と、報告書ごとの書き出し"""

    def __init__(self, work, write=True, template=None):
        self.work, self.out, self.template, self.answers = work, work / "out", template, None
        self.made = None if write else {}                            # 点検用（write=False）: 書かずに、ファイル名 → 中身の文字を貯める
        if write:
            self.out.mkdir(parents=True, exist_ok=True)
        self.m = manifest(work)
        self.lab = labels(self.m)
        self.flows = load_flows(work)
        self.flow_names = {f["flow_id"]: f.get("name", "") for f in self.flows}
        for f in sorted((self.work / "activities").glob("*.activity.json")) if (self.work / "activities").exists() else []:   # アクティビティ図の確認事項の「対象」
            act = load(f, {}) or {}
            if act.get("flow_id"):
                self.flow_names[act["flow_id"]] = f"{act.get('name', '')}（アクティビティ図）"
        self.qs = collect_questions(work, write)                     # write=False は点検用（番号を引くだけで、questions.json を書かない）
        self.hier = load(work / "structure" / "hierarchy.json", {}) or {}
        self.tax = load(work / "taxonomy.json", {}) or {}
        self.roles = {d.get("doc_id"): d.get("role") for d in self.hier.get("docs") or []}
        self.items_of = CardItems(work)                              # そのページのカードの項目（手順・分岐・図の箱）

    def src(self, s):
        return src_label(self.m, s)

    def emit(self, name, lines):
        if self.made is not None:
            self.made[name] = "\n".join(lines) + "\n"
        else:
            write_md(self.out / name, lines)

    def all(self):
        """報告書をすべて作る（点検用なら、書かずに中身を返す）"""
        self.questions()
        counts = self.sources()
        self.hierarchy()
        reviewed = self.reviews()
        return counts, reviewed

    def evidence(self, r):
        """根拠の文（原文）と、そのページ"""
        text = cell(r.get("evidence", ""))
        return (text and f"「{text}」") + ("（" + "、".join(self.src(x) for x in r.get("sources") or []) + "）" if r.get("sources") else "")

    # ---- 確認事項一覧 ----

    def question_row(self, q):
        origin = f"{q['flow']} {self.flow_names.get(q['flow'], '')}".strip() if q.get("flow") else ORIGIN.get(q["origin"], q["origin"])
        srcs = "、".join(self.src(s) for s in q.get("sources") or [])
        return [q["id"], SEV.get(q["severity"], q["severity"]), q["category"], q.get("effect", ""), origin, target(q), q["question"], q["assumption"], srcs, q["ask"]]

    def questions(self):
        qs = self.qs
        rows = [self.question_row(q) for q in sorted(qs, key=lambda q: (SEV_ORDER.get(q["severity"], 1), q["id"]))]
        md = ["# 確認事項一覧", "",
              f"全 {len(qs)} 件（高 {sum(q['severity'] == 'high' for q in qs)}・中 {sum(q['severity'] == 'mid' for q in qs)}・"
              f"低 {sum(q['severity'] == 'low' for q in qs)}）。番号はフロー図のバッジと同じ。", "",
              "載せているのは、答えによって図が変わる問いだけ（変わる所は「答えで変わる所」の列）。値だけの確かめ・名前や呼び方だけの確かめは、載せていない。", "",
              "重要度: " + "／".join(f"{SEV[k]}＝{v}" for k, v in SEV_NOTE.items()), ""] + table(Q_HEAD)
        self.emit("確認事項一覧.md", md + [row(r) for r in rows])
        if self.made is not None:
            return
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            print("openpyxl が無いので 確認事項一覧.xlsx は作らない（md だけ）")
            return
        from table_template import Cell, TemplateError, fill
        kept = (self.answers or Answers(self.work)).kept(Q_BOOK)      # 人が埋める欄の値（取り込んだもの）を、同じ番号の行に戻す
        try:
            tpl = fill(self.template or ASSETS / Q_TEMPLATE, Q_SPEC, {Q_SHEET: [[Cell(h) for h in Q_HEAD]] + [[Cell(str(v)) for v in r] for r in rows]}, {}, self.out / Q_BOOK,
                       kept=kept)
        except TemplateError as e:
            raise SystemExit(f"確認事項一覧のテンプレートが、決まりどおりでない（python3 scripts/table_template.py check で確かめる）:\n{e}")
        for n in tpl.notes:
            if "人が埋める欄" not in n or self.template:                # 既定のテンプレートの回答欄は、毎回は知らせない
                print(f"テンプレート: {n}")

    def take_answers(self, books):
        """書き出す前に、ブックの人が埋める欄（回答・回答者・回答日など）を、Q 番号で answers.json に取り込む。返すのは知らせ。
        読めないブック・列や行を動かしたブックは、上書きせずに止める（人が入れた回答を落とさない）"""
        from table_template import TemplateError, manual_back
        self.answers = Answers(self.work)
        now = {q["id"]: q["question"] for q in self.qs}
        said = []
        for p in books:
            try:
                got, notes = manual_back(p, Q_SPEC, self.template or ASSETS / Q_TEMPLATE)
            except TemplateError as e:
                raise SystemExit(f"{p.name}: {e}。人が入れた回答を取り込めないので、上書きしない（取り込まずに作り直すなら --fresh）")
            except Exception as e:                                    # 壊れた xlsx は、どの例外で落ちるか決まっていない
                raise SystemExit(f"{p.name} を読めない（{type(e).__name__}: {str(e)[:60]}）。人が入れた回答を取り込めないので、上書きしない（取り込まずに作り直すなら --fresh）")
            n = self.answers.take(Q_BOOK, got, {Q_SHEET: now})
            said += notes + ([f"人が埋める欄の値を取り込んだ: {p.name} から {n} 行（answers.json）"] if n else [])
        gone, changed = answer_gaps(self.answers, Q_BOOK, Q_SHEET, now)
        said += ([f"WARN  確認事項に無くなった番号の回答 {len(gone)} 件（{'・'.join(gone[:8])}）。answers.json に残してある（図に反映して確認事項が消えたのなら、そのままでよい）"] if gone else []) \
            + ([f"WARN  回答を取り込んだあとで、質問の文が変わった確認事項 {len(changed)} 件（{'・'.join(changed[:8])}）。回答が、いまの問いに合っているか見直す"] if changed else [])
        self.answers.save()                                          # ブックより先に書く（ブックのほうが新しくなる。進み具合が「作り直し」と見ない）
        return said

    # ---- 出典対応表 ----

    def doc_rows(self):
        """資料の表（題名・ファイル・役割・ページ数・ページ番号の数え方・備考）。資料は題名で書く（資料ID は出さない）"""
        out = []
        for d in self.m["docs"]:
            note = f"「{self.lab.title(d['duplicate_of'])}」と同一" if d.get("duplicate_of") else (d.get("render_note") or "")
            info = self.lab.docs.get(d["doc_id"], {})
            how = NUMBERING.get(info.get("basis"), "")
            if info.get("basis") == "printed" and info.get("numbered", 0) < info.get("paged", 0):
                how += f"（番号の無いページ {info['paged'] - info['numbered']}。表紙・目次などは種別で書く）"
            if info.get("basis") == "count" and any(p.get("kind") == "slide" for p in d.get("page_list") or []):
                how = "スライド番号"
            out.append([self.lab.title(d["doc_id"]), d["rel_path"], ROLE.get(self.roles.get(d["doc_id"]), ""), d.get("pages") or "", how, note])
        return out

    def node_rows(self, fl, cited):
        """フロー1本の、ノード → 出典の行。cited（ページ → 使っているノード）に足していく"""
        F = fl["flow_id"]
        for n in fl["nodes"]:
            for s in (n.get("sources") or []) if n["type"] not in ("start", "end") else []:
                users = cited.setdefault((s.get("doc_id"), s.get("page")), [])
                if f"{F} {n['id']}" not in users:
                    users.append(f"{F} {n['id']}")
        return [[d["id"], KIND.get(d["type"], d["type"]), d["lane"],
                 d["label"] + (f"（{d['detail']}）" if d["detail"] else "") + (f"〔参考: {d['reference']}〕" if d["reference"] else ""),
                 "、".join(d["sources"]), d["quote"], "推定" if d["inferred"] else "", "補足" if d["note"] else "", ", ".join(d["questions"])]
                for d in node_details(fl, self.m, self.qs, self.items_of)]

    def page_rows(self, cited):
        """ページ → ノードの行と、(内容ページの数, 出典になっているページの数)"""
        reasons, items = unused_reasons(self.work, self.flows), item_reasons(self.flows)     # 点検（check_flow.py）と同じ見方
        out, n_content, n_used = [], 0, 0
        for d, p in all_pages(self.m):
            c = load_card(self.work, d["doc_id"], p["page"])
            k = (d["doc_id"], p["page"])
            content = c.get("page_type") not in NON_CONTENT_TYPES or bool(self.items_of(*k))
            n_content += content
            n_used += content and k in cited
            why = "" if k in cited else (why_unused(k, reasons, items) or ("（対象外の種別）" if not content else "★理由なし"))
            out.append(row((self.lab.title(d["doc_id"]), self.lab.short(*k), PTYPE.get(c.get("page_type"), c.get("page_type", "")), c.get("title", ""),
                            "、".join(cited.get(k, [])), why)))
        return out, n_content, n_used

    def excluded_rows(self):
        """フローに入れなかった手順・分岐と理由"""
        out = []
        for fl in self.flows:
            for x in fl.get("excluded_items") or []:
                it = self.items_of(x.get("doc_id"), x.get("page")).get(x.get("item"), {})
                out.append(row((fl["flow_id"], self.src(dict(x, quote=it.get("quote", ""))), ITEM_KIND.get(str(x.get("item") or "")[:1], ""), it.get("text", ""),
                                "補足" if it.get("origin") == "note" else "", x.get("reason", ""))))
        return out or ["| （なし） | | | | | |"]

    def sources(self):
        md = ["# 出典対応表", "", "フローのノードが、どの資料の何ページを根拠にしているかの対応。", "",
              "資料は題名で、ページは、資料に刷られたページ番号で書いている（PowerPoint はスライド番号、Excel はシート名とセル）。", "", "## 資料", ""]
        md += table(["資料の題名", "ファイル", "役割", "ページ数", "ページ番号の数え方", "備考"]) + [row(r) for r in self.doc_rows()]
        md += ["", "## ノード → 出典（フローごと）"]
        node_head = ["ノード", "種別", "担当", "内容", "出典", "引用", "推定", "補足由来", "確認事項"]
        cited, csv_rows = {}, []
        for fl in self.flows:
            rows = self.node_rows(fl, cited)
            md += ["", f"### {fl['flow_id']} {fl.get('number', '')} {fl['name']}".rstrip(), ""] + table(node_head) + [row(r) for r in rows]
            csv_rows += [[fl["flow_id"], fl["name"]] + r for r in rows]
        page_rows, n_content, n_used = self.page_rows(cited)
        md += ["", "## ページ → ノード（資料ごと）", "", "「使っていない」ページには理由を書く。「★理由なし」の内容ページは check_flow.py が FAIL にする。", ""]
        md += table(["資料", "ページ", "種別", "見出し", "使っているノード", "使っていない理由"]) + page_rows
        md.insert(3, f"内容ページ {n_content} のうち、ノードの出典になっているのは {n_used}（{(n_used / n_content if n_content else 1):.0%}）。")
        md += ["", "## フローに入れなかった手順・分岐", "", "カードに書かれた手順・分岐のうち、どのノードにもしなかったものと理由。", ""]
        md += table(["フロー", "出典", "項目", "内容", "補足由来", "理由"]) + self.excluded_rows()
        self.emit("出典対応表.md", md)
        buf = io.StringIO(newline="")
        wr = csv.writer(buf)
        wr.writerow(["フローID", "フロー名"] + node_head)
        wr.writerows(csv_rows)
        if self.made is not None:
            self.made["出典対応表.csv"] = buf.getvalue()
        else:
            with open(self.out / "出典対応表.csv", "w", newline="", encoding="utf-8-sig") as f:
                f.write(buf.getvalue())
        return n_content, n_used

    # ---- 階層と業務分類 ----

    def tree_rows(self, of_h):
        """業務の木の行（結び付く資料ごとに1行。業務の名前は最初の行だけ）"""
        out = []

        def walk(nodes, depth):
            for n in nodes or []:
                name = "　" * depth + str(n.get("title") or "")
                for i, r in enumerate(n.get("refs") or [{}]):
                    pages = "、".join(self.lab.short(r["doc_id"], p) for p in r.get("pages") or []) if r.get("doc_id") else ""
                    out.append(row((name if i == 0 else "", "・".join(of_h.get(n.get("id"), [])) if i == 0 else "", self.lab.title(r["doc_id"]) if r.get("doc_id") else "",
                                    ROLE.get(self.roles.get(r.get("doc_id")), ""), REF_REL.get(r.get("rel"), ""), pages,
                                    self.evidence(r) if r else "")))
                walk(n.get("children"), depth + 1)
        walk(self.hier.get("tree"), 0)
        return out

    def hierarchy(self):
        hier, lab, cats = self.hier, self.lab, self.tax.get("categories") or []
        of_h = {}                                                    # 業務の木の id → 業務分類
        for c in cats:
            if c.get("hierarchy_id"):
                of_h.setdefault(c["hierarchy_id"], []).append(c["id"])
        md = ["# 階層と業務分類", "", "入力資料の階層の判定と、業務の木・資料の結び付き・資料どうしの関係。", "",
              "- 階層構造: " + HIERARCHY_LABEL.get(hier.get("is_hierarchical"), "未判定"),
              f"- 判定の理由: {cell(hier.get('rationale', ''))}",
              "- 業務分類の決め方: " + TAX_BASIS.get(self.tax.get("basis"), "未設定"),
              "", "## 業務の木と、結び付く資料", ""]
        md += table(["業務", "業務分類（フロー）", "資料", "役割", "結び付き", "ページ", "根拠"]) + self.tree_rows(of_h)
        out_scope = [d for d in hier.get("docs") or [] if d.get("out_of_scope")]
        if out_scope:
            md += ["", "どの業務にも結び付けていない資料: " + "、".join(f"{self.lab.title(d.get('doc_id'))}（{cell(d['out_of_scope'])}）" for d in out_scope)]
        md += ["", "## 資料どうしの関係", ""] + table(["資料", "関係", "資料", "根拠"])
        md += [f"| {cell(lab.title(r.get('from')))} | {REL.get(r.get('type'), r.get('type'))} | {cell(lab.title(r.get('to')))} | {self.evidence(r)} |"
               for r in hier.get("doc_relations") or []] or ["| （なし） | | | |"]
        titles = {n.get("id"): n.get("title") for n, _parent in tree_nodes(hier.get("tree"))}
        md += ["", "## 業務どうしの「続く」", ""] + table(["前の業務", "続く業務", "根拠"])
        md += [f"| {cell(titles.get(q.get('from'), ''))} | {cell(titles.get(q.get('to'), ''))} | {self.evidence(q)} |"
               for q in hier.get("sequence") or []] or ["| （なし） | | |"]
        md += ["", "## 業務分類とページの割当", ""] + table(["分類", "番号", "名前", "根拠", "フロー", "割り当てたページ"])
        for c in cats:
            pages = "／".join(lab.join(ps["doc_id"], "、".join(lab.short(ps["doc_id"], p) for p in ps.get("pages") or [])) for ps in c.get("pages") or [])
            md.append(row((c.get("id"), c.get("number", ""), c.get("name"), "目次" if c.get("basis") == "toc" else "LLM",
                           "あり" if c.get("make_flow") else "", pages)))
        self.emit("階層と業務分類.md", md)

    # ---- 原典照合レビューの採否表 ----

    def reviews(self):
        """review/*.review.json をまとめる。レビューをしていなければ何も書かない。(レビューの本数, 指摘の件数) を返す"""
        files = sorted((self.work / "review").glob("*.review.json")) if (self.work / "review").exists() else []
        if not files:
            return 0, 0
        import review as RV
        reviews = {}
        body, k = table(["#", "フロー", "対象", "指摘（原文のまま）", "種別", "採否", "対応／理由"]), 0
        for f in files:
            rv, err = try_load(f)
            if err or not isinstance(rv, dict):
                raise SystemExit(f"review/{f.name} を読めない（{err or 'オブジェクトでない'}）。採否表を直してから書き出し直す")
            filled = copy.deepcopy(rv)                                   # 採否の未記入は、点検（--stage review）が FAIL にする。ここでは、ほかの崩れだけで止める
            for x in filled.get("findings") if isinstance(filled.get("findings"), list) else []:
                if isinstance(x, dict):
                    x["decision"] = x.get("decision") or "保留"
                    x["action"] = x.get("action") or "未記入"
            errs = validate(filled, "review")
            if errs:
                raise SystemExit(f"review/{f.name} の形が違う — " + "／".join(errs[:5]) + "。採否表を直してから書き出し直す")
            F = rv.get("flow_id") or f.name.split(".")[0]
            reviews[F] = rv
            for x in rv.get("findings") if isinstance(rv.get("findings"), list) else []:
                if isinstance(x, dict):
                    k += 1
                    body.append(row((k, F, x.get("target"), x.get("text"), x.get("kind"), x.get("decision"), x.get("action"))))
        heads, gaps = [], []
        for kind, lead in RV.tally(RV.figures(self.work), reviews).items():
            if not lead["figs"]:
                continue
            nodes, judged, tot, name = lead["nodes"], lead["judged"], lead["tot"], {"flow": "", "activity": "アクティビティ図: "}[kind]
            heads.append(f"{name}開始・終了を除くノード {nodes} 個のうち、判定したのは {judged} 個: 支持 {tot['支持']}・一部支持 {tot['一部支持']}・不支持 {tot['不支持']}"
                         + (f"・その他 {tot['その他']}" if tot["その他"] else "") + f"。支持率 {(tot['支持'] / nodes if nodes else 0):.0%}（支持 ÷ 開始・終了を除く全ノード）。")
            gaps += ([f"{name}判定の無いノード {nodes - judged} 個（判定のあとに足したノードを含む）。"] if nodes > judged else []) \
                + ([f"{name}判定のあとに無くしたノード {lead['dropped']} 個の判定は、数えていない。"] if lead["dropped"] else []) \
                + ([f"{'レビューをしていないフロー' if kind == 'flow' else '照合していないアクティビティ図'}: {'・'.join(lead['rest'])}。"] if lead["rest"] else [])
        heads = heads or [""]
        heads[0] += f"指摘 {k} 件。"
        self.emit("原典照合レビュー_採否表.md", [
            "# 原典照合レビューの採否表", "", "作り方を知らない別のエージェントが、原典とフローを照らして出した判定と指摘。判定は、指摘を反映する前のフローに対するもの（採用した指摘は「対応／理由」のとおりに直してある）。", ""]
            + heads + (["", "**" + " ".join(gaps) + "**"] if gaps else []) + [""] + body)
        return len(files), k


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--template", help=f"確認事項一覧のテンプレートの xlsx（既定: 前の書き出しで使ったもの。はじめは assets/{Q_TEMPLATE}。決まりは references/template-spec.md）")
    ap.add_argument("--keep", help="回答を入れたブック（お客様から戻ってきた確認事項一覧の xlsx）。回答・回答者・回答日と、様式で足した人が埋める欄を取り込んでから、書き出す")
    ap.add_argument("--fresh", action="store_true", help="いま out にあるブックから、人が入れた値を取り込まない（読めないブックを作り直すとき。answers.json にある値は戻す）")
    a = ap.parse_args()
    work = work_dir(a.work)
    blocking(work)
    from table_template import lock_of, remembered
    book = work / "out" / Q_BOOK
    chosen = a.template or remembered(book, ASSETS / Q_TEMPLATE)
    if chosen and not a.template:
        print(f"テンプレート: 前の書き出しと同じ {pathlib.Path(chosen).name} を使う（既定の様式に戻すなら --template {ASSETS / Q_TEMPLATE}）")
    for p in [book] + ([pathlib.Path(a.keep)] if a.keep else []):
        if lock_of(p):                                                # 開いたまま書き換えると、人が入れた回答を取り込めず、上書きもできない
            raise SystemExit(f"{p.name} を表計算ソフトで開いている（{lock_of(p).name} がある）。閉じてから、もう一度")
    if a.keep and not pathlib.Path(a.keep).exists():
        raise SystemExit(f"--keep のブックが無い: {a.keep}")
    try:
        r = Reports(work, template=chosen)
        said = r.take_answers(([book] if book.exists() and not a.fresh else []) + ([pathlib.Path(a.keep)] if a.keep else []))
    except BadJson as e:
        raise SystemExit(str(e))
    (n_content, n_used), (n_reviews, n_findings) = r.all()
    if n_reviews:
        print(f"採否表 → {r.out / '原典照合レビュー_採否表.md'}（レビュー {n_reviews} 本・指摘 {n_findings} 件）")
    print(f"確認事項 {len(r.qs)} 件 → {r.out / '確認事項一覧.md'} / .xlsx")
    print(f"出典対応表 → {r.out / '出典対応表.md'} / .csv（内容ページ {n_content}・使用 {n_used}）")
    print(f"階層と業務分類 → {r.out / '階層と業務分類.md'}")
    for x in said:
        print(x)


if __name__ == "__main__":
    main()
