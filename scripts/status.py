#!/usr/bin/env python3
"""status.py — いまどこまで進んでいるかを、作業フォルダだけを見て1画面で出す。次に打つコマンドも出す。

  python3 scripts/status.py -w <作業フォルダ>
  python3 scripts/status.py -w <作業フォルダ> --fast     # 点検（FAIL の数）を走らせず、ファイルの有無だけで見る

状態はすべて作業フォルダにある。会話の記憶（文脈）には頼らない。次のときは、必ずこれから始める:
  ・親（進行役）の文脈が圧縮された・あふれそうになったあと
  ・別のセッションで続きをするとき
  ・サブエージェントが途中で止まったあと（読み残したページ・書けていないフローが「残り」に出る）
どの工程も、済んだ資料・ページ・フローは飛ばして、残りだけをやり直せる。
"""
import argparse
import pathlib
import re
import subprocess
import sys

from common import ACT_FIG, DELIVERABLES, MATRIX_BOOK, NUMBERS_FILE, REPORTS, all_pages, card_json, is_code_page, load, split_figs, try_load
from validate import files_of, validate_file

S = pathlib.Path(__file__).resolve().parent


def fails(script, work, *args):
    """点検を走らせて (FAIL の数, WARN の数, 動かなかった理由) を返す。
    集計の行が出なかった（点検が途中で止まった）ときは、FAIL 1 件として扱う。「点検が動かなかった」を「FAIL 0」と取り違えない"""
    r = subprocess.run([sys.executable, str(S / script), "-w", str(work), *args], capture_output=True, text=True)
    mt = re.search(r"(\d+) FAIL / (\d+) WARN", r.stdout)
    if mt:
        return int(mt.group(1)), int(mt.group(2)), None
    return 1, 0, f"{script} が動かない（{(r.stderr.strip().splitlines() or ['出力なし'])[-1][:80]}）"


def read(path):
    """(中身, 読めない理由)。無ければ (None, None)。辞書でなければ読めない扱い"""
    d, err = try_load(path)
    if err is None and d is not None and not isinstance(d, dict):
        d, err = None, "全体がオブジェクトになっていない"
    return d, err


def tally(c):
    return "" if not c else f"・{c[2]}" if c[2] else f"・点検 {c[0]} FAIL / {c[1]} WARN"


def newest(paths):
    ts = [p.stat().st_mtime for p in paths if p.exists()]
    return max(ts) if ts else 0


W = "<作業フォルダ>"                                                  # 次に打つコマンドの中で、作業フォルダを書く所


def source_files(work, acts):
    """図・表・報告書のもとになるデータ（これより古い書き出しは、作り直す）。カード（引用・出典の書き方）・用語・棚卸し（資料の題名）も、図と報告書に出る"""
    return (files_of(work, "flow") + [work / "taxonomy.json", work / "structure" / "hierarchy.json", work / "glossary.json", work / "manifest.json"]
            + list(acts) + ([work / "structure" / "systems.json"] if acts else []) + list((work / "cards").glob("*/*.json")))


class Status:
    """作業フォルダの中身だけから、工程ごとの進み具合（rows）と、次に打つコマンド（todo）を組み立てる"""

    def __init__(self, work, m, fast):
        self.work, self.m, self.fast = work, m, fast
        self.out = work / "out"
        self.rows, self.todo = [], []
        self.blocked = []                                            # 前の工程に直す所が残っている（書き出し・納品に進まない）

    def row(self, no, name, state, text):
        if getattr(self, "code_only", False) and no in (3, 4, 5, 6, 8) and state == "未":
            state, text = "―", ("アクティビティ図の原典照合は、references/code-review-prompt.md で行う（業務フローの照合シートは無い）" if no == 8
                                else "マニュアルが無い（プログラムのソースだけ）ので、しない")
        self.rows.append(f"{no} {name:<8} {state:<3} {text}")

    def bad_shape(self, kind):
        return sum(bool(validate_file(f, kind)) for f in files_of(self.work, kind))

    def check(self, script, *args):
        """点検を実際に回して (FAIL の数, WARN の数, 動かなかった理由)。--fast のときは回さない"""
        return None if self.fast else fails(script, self.work, *args)

    def run(self):
        for step in (self.inventory, self.rendering, self.hierarchy, self.reading, self.structure, self.programs, self.flows, self.export, self.review,
                     self.delivery, self.gate):
            step()

    def inventory(self):
        m = self.m
        self.docs = [d for d in m["docs"] if not d.get("duplicate_of") and d.get("kind") != "unsupported"]
        self.row(1, "棚卸し", "済み", f"資料 {len(m['docs'])} 件（読む対象 {len(self.docs)}・重複 {sum(bool(d.get('duplicate_of')) for d in m['docs'])}）")

    def rendering(self):
        # inventory.py は rendered: false を書く。描画を試して失敗した資料には、理由（render_note）が残る
        self.wait = wait = [d["doc_id"] for d in self.docs if not d.get("rendered") and not d.get("render_note")]   # まだページ画像にしていない資料
        failed = [d["doc_id"] for d in self.docs if not d.get("rendered") and d.get("render_note")]
        self.pages = list(all_pages(self.m))
        extra = [(d, p) for d, p in all_pages(self.m, code=True) if (d["doc_id"], p["page"]) not in {(x["doc_id"], y["page"]) for x, y in self.pages}]
        code = sum(1 for _d, p in extra if is_code_page(p))
        ref = len(extra) - code                                          # 道具のブックのシート（参考にするだけ。カードにしない）
        self.code_only = bool(code) and not self.pages and not wait      # マニュアルが無い（階層・読取・分類・フロー・原典照合の工程は、しない）
        self.tool_hint = [d["doc_id"] for d in self.docs if d.get("code") and d.get("kind") != "code" and not d.get("tool")
                          and any(not is_code_page(p) for p in d.get("page_list") or [])]
        self.row(2, "ページ化", "途中" if wait else "済み", f"{len(self.pages)} ページ" + (f"（ほかに、プログラムのソース {code} ページ）" if code else "")
                 + (f"（道具のブックのシート {ref} ページは、参考）" if ref else "")
                 + (f"・まだの資料 {len(wait)} 件（{', '.join(wait[:8])}）" if wait else "")
                 + (f"・描画できなかった資料 {len(failed)} 件（{', '.join(failed[:8])}。確認事項に載る）" if failed else ""))
        if wait:
            self.todo.append(f"python3 scripts/render_pages.py -w {W}    # まだの資料をページ画像にする")

    def hierarchy(self):
        hier, err = read(self.work / "structure" / "hierarchy.json")
        hb = bool(err) or (self.bad_shape("hierarchy") if hier is not None else 0)
        toc = hier.get("toc_flow_list") if hier and isinstance(hier.get("toc_flow_list"), dict) else {}
        self.row(3, "階層", "要直し" if hb else "未" if hier is None else "済み",
                 f"hierarchy.json が JSON として読めない（{err}）" if err else "hierarchy.json が無い" if hier is None
                 else f"階層 {hier.get('is_hierarchical')}・目次の業務一覧 {'あり' if toc.get('found') else 'なし'}" + ("・形の違反あり" if hb else ""))
        if hb:
            self.blocked.append("階層")
            self.todo.append(f"python3 scripts/validate.py -w {W} --kind hierarchy    # hierarchy.json の形を直す")
        elif hier is None and not self.wait and self.pages:
            self.todo.append(f"python3 scripts/detect_structure.py -w {W}    # 手がかりを見て structure/hierarchy.json を書く（references/structure-guide.md）")

    def reading(self):
        work, pages = self.work, self.pages
        self.lack = lack = [(d["doc_id"], p["page"]) for d, p in pages if not card_json(work, d["doc_id"], p["page"]).exists()]   # カードの無いページ
        left, rb = set(), self.bad_shape("report_read")              # そのうち、読み残しの報告があるページ／形の崩れた報告
        for f in files_of(work, "report_read"):
            r = read(f)[0] or {}                                     # 壊れた報告は、報告なしとして扱う
            left |= {(r.get("doc_id"), x) for x in (r.get("pages_left") if isinstance(r.get("pages_left"), list) else [])} & set(lack)
        self.ce = ce = None if len(lack) == len(pages) else self.check("check_extract.py")
        state = "未" if len(lack) == len(pages) else "途中" if lack else "要直し" if ce and ce[0] else "済み"
        self.row(4, "読取", state, f"カード {len(pages) - len(lack)}/{len(pages)} ページ"
                 + (f"・残り {len(lack)}（{', '.join(f'{d}/p{p:04d}' for d, p in lack[:5])}{' …' if len(lack) > 5 else ''}）" if lack else "")
                 + (f"・うち読み残しの報告があるページ {len(left)}" if left else "") + (f"・形の崩れた報告 {rb}" if rb else "") + tally(ce))
        if rb:
            self.blocked.append("読取")
            self.todo.append(f"python3 scripts/validate.py -w {W} --kind report_read    # 形の崩れた読取の報告 {rb} 件を直す（読み残したページを束に戻すため）")
        if lack:
            self.blocked.append("読取")
        if lack and not self.wait:
            self.todo.append(f"python3 scripts/pages.py batches -w {W}    # 残り {len(lack)} ページを束にして読ませる（references/extraction-guide.md）")
        elif ce and ce[0]:
            self.blocked.append("読取")
            self.todo.append(f"python3 scripts/check_extract.py -w {W}    # " + (ce[2] or f"FAIL {ce[0]} 件を、そのページの担当に直させる"))

    def structure(self):
        work, ce = self.work, self.ce
        (gl, gerr), (tax, terr) = read(work / "glossary.json"), read(work / "taxonomy.json")
        self.tax = tax
        rows = lambda d, key: [x for x in (d.get(key) if d and isinstance(d.get(key), list) else []) if isinstance(x, dict)]
        self.cats = cats = [c for c in rows(tax, "categories") if c.get("make_flow") and isinstance(c.get("id"), str)]
        st = (1, 0, f"JSON として読めない（{gerr or terr}）") if gerr or terr else None if tax is None else self.check("check_flow.py", "--stage", "structure")
        self.flow_ids = {f.name.split(".")[0] for f in files_of(work, "flow")}
        self.row(5, "用語・分類", "要直し" if st and st[0] else "未" if tax is None else "済み",
                 ("glossary.json なし・" if gl is None else f"用語 {len(rows(gl, 'terms'))}・")
                 + ("taxonomy.json なし" if tax is None else f"分類 {len(rows(tax, 'categories'))}（フローを作る {len(cats)}）") + tally(st))
        if st and st[0]:
            self.blocked.append("用語・分類")
            self.todo.append(f"python3 scripts/check_flow.py -w {W} --stage structure    # " + (st[2] or f"FAIL {st[0]} 件（階層・用語・業務分類）を直す"))
        elif not self.lack and not (ce and ce[0]) and self.pages:
            if gl is None:
                self.todo.append(f"python3 scripts/pages.py terms -w {W} --draft    # 表記の一覧を見て glossary.json（用語の正本）を書く")
            if tax is None:
                self.todo.append(f"python3 scripts/pages.py taxonomy -w {W} --draft    # 目次に業務の一覧があれば、taxonomy.json の下書きを機械で書く。"
                                 f"無ければ pages.py outline で全体を見て書く（references/structure-guide.md）")

    def programs(self):
        """プログラム（マクロ・スクリプト・ソース）があるときだけ出る工程: 構造 → システム台帳 → アクティビティ図（references/code-guide.md）"""
        import code_struct as cs
        from systems import Ledger, has_ledger
        work = self.work
        projs = cs.projects(self.m)
        self.acts = files_of(work, "activity")
        if not projs:
            return
        structs = cs.load_structs(work)
        lost = [p["id"] for p in projs if p["id"] not in structs]
        notree = [p["id"] for p in projs if (structs.get(p["id"]) or {}).get("read") and (not structs[p["id"]]["read"]["tree"] or structs[p["id"]]["read"].get("errors"))]
        ledger = has_ledger(work)
        sc = self.check("check_flow.py", "--stage", "systems") if ledger and not lost else None
        units, have, ready = [], [], []
        if ledger and not lost and not (sc and sc[0]):
            led = Ledger(work, self.m)
            units = [u["activity"] for u in led.units if led.off(u) is None]      # 図にしない起動点・図にしない手続き（補助）は、数えない
            have = [a for a in units if (work / "activities" / f"{a}.activity.json").exists()]
            ready = [a for a in units if a not in have and (work / "activities" / f"{a}.edits.json").exists()]
        ac = self.check("check_flow.py", "--stage", "activity") if units and len(have) == len(units) else None
        done = ledger and not lost and not notree and not (sc and sc[0]) and len(have) == len(units) and not (ac and ac[0])
        self.row("5b", "プログラム", "済み" if done else "未" if lost or not ledger else "要直し" if (sc and sc[0]) or (ac and ac[0]) else "途中",
                 f"まとまり {len(projs)}・構造 {'まだ ' + '・'.join(lost[:5]) if lost else 'あり'}" + (f"・LLM が読む木がまだ・誤りあり {'・'.join(notree[:5])}" if notree else "")
                 + f"・システム台帳 {'あり' if ledger else 'なし'}"
                 + (f"・アクティビティ図 {len(have)}/{len(units)} 枚" if units else "") + tally(sc) + tally(ac))
        if done:
            return
        self.blocked.append("プログラム")                            # フローの起動点の箱は、台帳とアクティビティ図の名前から置く
        if self.wait:
            return
        if lost:
            self.todo.append(f"python3 scripts/code_struct.py -w {W}    # プログラムの構造（手続き・分岐・呼び出し・入出力・起動点）を機械で取り出す")
        elif notree:
            self.todo.append(f"python3 scripts/code_struct.py -w {W} --hints    # 機械で構文を読めない言語（{'・'.join(notree[:5])}）: ソースを読んで、ブロックの木 code/<ID>.tree.json を書く"
                             f"（references/code-guide.md）。書いたら code_struct.py をもう一度")
        elif not ledger:
            if ((work / "glossary.json").exists() and not self.lack) or self.code_only:
                self.todo.append(f"python3 scripts/systems.py draft -w {W}    # システム台帳の下書きを書く。見て、structure/systems.json を直す（references/code-guide.md）")
        elif sc and sc[0]:
            self.todo.append(f"python3 scripts/check_flow.py -w {W} --stage systems    # " + (sc[2] or f"FAIL {sc[0]} 件（システム台帳）を直す"))
        elif ready:
            self.todo.append(f"python3 scripts/activity.py apply -w {W} --all    # 直しの一覧がある {len(ready)} 枚を作る")
        elif len(have) < len(units):
            self.todo.append(f"python3 scripts/activity.py list -w {W}    # 残り {len(units) - len(have)} 枚: 下書き（activity.py draft --id …）を見て、"
                             f"直しの一覧 activities/<番号>.edits.json を書く（references/code-guide.md）")
        else:
            self.todo.append(f"python3 scripts/activity.py apply -w {W} --all && python3 scripts/check_flow.py -w {W} --stage activity    # "
                             + (ac[2] or f"作り直してから点検する。残った FAIL {ac[0]} 件を、その図の担当に直させる"))

    def flows(self):
        cats = self.cats
        self.miss = miss = [c["id"] for c in cats if c["id"] not in self.flow_ids]      # まだフローの無い分類
        fb, rfb = self.bad_shape("flow"), self.bad_shape("report_flow")
        self.fc = fc = None if not self.flow_ids else self.check("check_flow.py", "--stage", "flow")
        self.row(6, "フロー化", "未" if cats and len(miss) == len(cats) else "途中" if miss else "要直し" if fc and fc[0] else "済み" if cats else "未",
                 f"フロー {len(cats) - len(miss)}/{len(cats)} 本" + (f"・残り {', '.join(miss[:10])}{' …' if len(miss) > 10 else ''}" if miss else "")
                 + (f"・形の違反 {fb} 本" if fb else "") + (f"・形の崩れた報告 {rfb}" if rfb else "") + tally(fc))
        if rfb:
            self.todo.append(f"python3 scripts/validate.py -w {W} --kind report_flow    # 形の崩れたフロー担当の報告 {rfb} 件を直す")
        if self.tax is not None and miss and not self.blocked:
            self.todo.append(f"python3 scripts/pages.py plan-flows -w {W}    # 材料の大きさを見てから、残り {len(miss)} 本を分類ごとに作らせる（references/flow-edits-guide.md。機械の下書きに、直しの一覧を書く）")
        elif fc and fc[0] and not miss:
            self.todo.append(f"python3 scripts/quotes.py fill -w {W} && python3 scripts/renumber.py -w {W} && python3 scripts/check_flow.py -w {W} --stage flow    # "
                             + (fc[2] or f"引用・番号・推定の印を機械で入れ直してから点検する。残った FAIL {fc[0]} 件を、そのフローの担当に直させる"))
        self.ready = bool(cats) and not miss and not (fc and fc[0]) and not self.blocked    # フローがそろい、ここまでの点検も通っている

    def export(self):
        work, out = self.work, self.out
        self.drawio = drawio = sorted(out.glob("*.drawio")) if out.exists() else []
        self.src_new = newest(source_files(work, self.acts))
        self.stale = stale = bool(drawio) and min(p.stat().st_mtime for p in drawio) < self.src_new   # 図より、もとのデータのほうが新しい
        need_act = bool(self.acts) and not (out / ACT_FIG).exists()
        book = out / MATRIX_BOOK
        old_book = book.exists() and newest([book]) < max(self.src_new, newest([work / "times.json", work / "answers.json"]))   # 業務マトリクス一覧より、もとのデータ・人が入れた値のほうが新しい
        self.row(7, "書き出し", "未" if not drawio else "作り直し" if stale or old_book or not book.exists() else "済み",
                 "drawio なし" if not drawio else f"drawio {len(drawio)} ファイル" + ("（flow.json のほうが新しい）" if stale else "")
                 + f"・PNG {len(list((out / 'png').glob('*.png'))) if (out / 'png').exists() else 0} 枚"
                 + ("・業務マトリクス一覧なし" if not book.exists() else "・業務マトリクス一覧" + ("（flow.json のほうが新しい）" if old_book else "")))
        if self.ready and (not drawio or stale or need_act):
            self.todo.append(f"python3 scripts/quotes.py fill -w {W} && python3 scripts/renumber.py -w {W} && python3 scripts/build_drawio.py -w {W}{' --split' if split_figs(out) else ''} "
                             + (f"&& python3 scripts/build_activity.py -w {W} " if self.acts else "")
                             + f"&& python3 scripts/check_layout.py -w {W} && python3 scripts/build_matrix.py -w {W} && python3 scripts/check_flow.py -w {W}"
                             + "    # 連鎖の順は納品の門と同じ（図 → PNG → 業務マトリクス一覧。PNG へのリンクが初回から付く）")
        elif self.ready and (old_book or not book.exists()):
            self.todo.append(f"python3 scripts/build_matrix.py -w {W} && python3 scripts/check_flow.py -w {W} --stage matrix    # 業務マトリクス一覧を書き出す")

    def review(self):
        out, cats = self.out, self.cats
        flows = {c["id"] for c in cats}
        ids = flows | {f.name.split(".")[0] for f in self.acts}      # 照合する図（業務フロー・アクティビティ図）
        sheets = {re.sub(r"_\d+$", "", f.stem.replace("照合シート_", "")) for f in out.glob("照合シート_*.pdf")} if out.exists() else set()
        reviews = {f.name.split(".")[0] for f in files_of(self.work, "review")}
        rc = self.check("check_flow.py", "--stage", "review") if reviews else None
        old = sorted(F for F in flows & sheets if newest(list(out.glob(f"照合シート_{F}.pdf")) + list(out.glob(f"照合シート_{F}_*.pdf")))
                     < newest([self.work / "flows" / f"{F}.flow.json"] + list((out / "png").glob(f"*_flow-{F}.png"))))   # シートより、フロー・図の PNG のほうが新しい
        self.row(8, "原典照合", "未" if not reviews else "途中" if ids - reviews else "要直し" if rc and rc[0] else "済み",
                 f"照合シート {len(sheets & flows)}/{len(flows)}・採否表 {len(reviews & ids)}/{len(ids)}"
                 + (f"・残り {', '.join(sorted(ids - reviews)[:10])}" if reviews and ids - reviews else "")
                 + (f"・古いシート {', '.join(old[:10])}" if old else "") + tally(rc))
        if self.ready and self.drawio and not self.stale and ids - reviews:
            todo = sorted(ids - reviews)
            if flows - sheets or old:
                self.todo.append(f"python3 scripts/make_compare.py -w {W}" + (f" --only {','.join(sorted((flows - sheets) | set(old)))}" if sheets else "")
                                 + "    # 照合シートを作る（フロー・図より古いものは作り直す）")
            else:
                self.todo.append(f"# 残り {', '.join(todo[:10])} を、図ごとに別のエージェントに照合させ（報告は reports/）、python3 scripts/review.py draft -w {W} で"
                                 f"採否表の下書きを作って、採否と対応を書く（references/source-review-prompt.md・code-review-prompt.md）")
        elif old and self.ready and self.drawio and not self.stale:
            self.todo.append(f"python3 scripts/make_compare.py -w {W} --only {','.join(old)}    # 古い照合シートを作り直す（作業の記録として渡す）")
        if rc and rc[0]:
            self.blocked.append("原典照合")
            self.todo.append(f"python3 scripts/check_flow.py -w {W} --stage review    # " + (rc[2] or f"FAIL {rc[0]} 件（採否の未記入・報告との違い・反映・支持率）を直す"))

    def delivery(self):
        out = self.out
        rep = [out / name for name in REPORTS]
        have = all(p.exists() for p in rep)
        old = have and newest(rep) < max(self.src_new, newest(files_of(self.work, "review") + [self.work / "answers.json"]))
        self.row(9, "納品", "未" if not have else "作り直し" if old else "済み",
                 "報告書なし" if not have else "確認事項一覧・出典対応表・階層と業務分類" + ("（もとのデータのほうが新しい）" if old else "")
                 + ("・業務マトリクス一覧" if (out / MATRIX_BOOK).exists() else ""))
        if self.ready and self.drawio and (not have or old):
            self.todo.append(f"python3 scripts/build_reports.py -w {W}")


    def gate(self):
        """納品の門（deliver.py）: 門を通って写したあと、出力が変わっていないか"""
        out = self.out
        nums = read(out / f"{NUMBERS_FILE}.json")[0] or {}
        made = [p for pat in DELIVERABLES for p in out.glob(pat)] if out.exists() else []
        done = bool(nums.get("to")) and bool(made) and (out / f"{NUMBERS_FILE}.json").stat().st_mtime >= newest(made)
        self.row(10, "納品の門", "済み" if done else "未", f"通って写した（{nums['to']}）" if done else
                 "deliver.py が書き出しと点検を回し直し、通ったら出力先の 納品物/・作業の記録/ に写す")
        if not done and not self.todo and made:
            self.todo.append(f"# 納品の門: python3 scripts/deliver.py -w {W} --to <出力先>    # 出力先（納品物/・作業の記録/ を作るフォルダ）は利用者に聞く")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--fast", action="store_true")
    a = ap.parse_args()
    work = pathlib.Path(a.work).expanduser().resolve()                # 見るだけの道具なので、無いフォルダは作らない
    m = load(work / "manifest.json") if work.is_dir() else None
    if m is None:
        print(f"== 進み具合 ==\n1 棚卸し    未   {'manifest.json が無い' if work.is_dir() else f'作業フォルダが無い: {work}'}\n\n"
              f"次にやること:\n  python3 scripts/inventory.py <入力フォルダ> -w {W}")
        return
    st = Status(work, m, a.fast)
    st.run()
    print(f"== 進み具合（{work.name}）==")
    print("\n".join(st.rows))
    print("\n次にやること:" if st.todo else "\nすべて済み。納品の門が写した出力先の一式を渡す")
    for t in st.todo[:3]:
        print("  " + t)
    if st.todo and getattr(st, "tool_hint", None):
        print(f"\nマクロ付きのブック（{', '.join(st.tool_hint[:8])}）のシートは、手順書として読む対象に入っている。シートに手順が書かれていない道具のブックなら、"
              f"利用者に確かめてから python3 scripts/inventory.py <入力> -w {W} --tool <資料ID,…>（シートは参考にするだけになる）")


if __name__ == "__main__":
    main()
