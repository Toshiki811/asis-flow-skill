#!/usr/bin/env python3
"""sc_status.py — 拠点間比較が、いまどこまで進んでいるかを、比較の作業フォルダだけを見て 1 画面で出す。次に打つコマンドも出す。

  python3 scripts/sc_status.py -w <比較の作業フォルダ>
  python3 scripts/sc_status.py -w <比較の作業フォルダ> --fast     # 拠点の門（asis の点検）を回さない

状態はすべて作業フォルダにある（会話の記憶に頼らない）。文脈が圧縮されたあと・別のセッションで続けるときは、必ずこれから始める。
拠点ごとの as-is は、asis-flow-skill の作業フォルダの進み具合（asis の status.py）も 1 行で出す。拠点の as-is は、拠点ごとに別のセッションで進めてよい。
次に打つコマンドのうち「python3 scripts/sc_…」で始まるものは機械の工程（sc_run.py が順に回す）。「#」で始まるものは LLM・人の仕事。
"""
import argparse
import re
import subprocess
import sys

from sc_common import ASIS_SCRIPTS, A, Report, load, match_path, shared_work, site_names, site_work, sites_of, snap_path, try_load, work_dir
from sc_validate import validate

W = "<比較の作業フォルダ>"
ASIS_CMD = f"python3 {ASIS_SCRIPTS}"


def stage_fails(cx, stage):
    """sc_check の段を 1 つ回して (FAIL の数, 最初の FAIL の文)"""
    import sc_check
    rp = Report(stage, fold=None)
    with A.guarded(rp, stage):
        sc_check.CHECKS[stage](cx, rp)
    return len(rp.fails), (rp.fails[0] if rp.fails else "")


def asis_status(w):
    """asis の status.py --fast の、工程の行（いちばん先の済んでいない工程）"""
    r = subprocess.run([sys.executable, str(ASIS_SCRIPTS / "status.py"), "-w", str(w), "--fast"], capture_output=True, text=True)
    rows = [ln for ln in r.stdout.splitlines() if re.match(r"^\d+ ", ln)]
    left = [ln for ln in rows if not re.search(r"\s(済み|―)\s", ln)]
    return (left[0] if left else (rows[-1] if rows else "（読めない）")).strip()


class Status:
    def __init__(self, cmp, fast):
        import sc_check
        self.cmp, self.fast = cmp, fast
        self.cx = sc_check.Ctx(cmp, fast)
        self.sites, self.names = self.cx.sites, site_names(self.cx.sites)
        self.rows, self.todo, self.blocked = [], [], False

    def row(self, no, name, state, text):
        self.rows.append(f"{no:<3}{name:<10}{state:<5}{text}")

    def add(self, *cmds):
        if not self.blocked:
            self.todo += list(cmds)

    STEPS = [("0", "始める", "start"), ("1a", "共通の資料", "shared"), ("1b", "書き分け", "applies"), ("1c", "拠点の as-is", "sites_asis"),
             ("2", "写し", "snapshot"), ("3", "用語の対応", "terms"), ("4", "業務種別", "axis"), ("5", "タスク", "match"), ("6", "計算", "compute"),
             ("7", "照合", "verify"), ("8", "書き出し", "outputs"), ("9", "納品の門", "gate")]

    def run(self):
        for no, name, fn in self.STEPS:
            if self.blocked and no not in ("1a", "1b", "1c"):
                self.row(no, name, "未", "前の工程が済んでから")
                continue
            getattr(self, fn)()

    def start(self):
        errs = validate(self.sites, "sites")
        self.row("0", "始める", "要直し" if errs else "済み", f"拠点 {len(self.cx.ids)}・基準 {self.names.get(self.sites.get('baseline'), 'なし')}"
                 + (f"・形の違反 {len(errs)}（{errs[0][:40]}）" if errs else ""))
        if errs:
            self.add("# sites.json を直す（references/site-guide.md の 0）")
            self.blocked = True

    def shared(self):
        sh = shared_work(self.cmp, self.sites)
        if sh is None:
            self.row("1a", "共通の資料", "―", "無い")
            return
        m = load(sh / "manifest.json", None)
        if not m:
            self.row("1a", "共通の資料", "未", "棚卸しをしていない")
            self.add(f"{ASIS_CMD}/inventory.py {' '.join(self.sites['shared']['inputs'])} -w {sh} && {ASIS_CMD}/render_pages.py -w {sh}    # 共通の資料の棚卸しとページ化")
            return
        pages = [(d, p) for d in m["docs"] if not d.get("duplicate_of") for p in d.get("page_list") or []]
        lack = [1 for d, p in pages if not A.card_json(sh, d["doc_id"], p["page"]).exists()]
        if not pages or lack:
            self.row("1a", "共通の資料", "途中", f"カード {len(pages) - len(lack)}/{len(pages)} ページ")
            self.add(f"# 共通の資料の読取: {ASIS_CMD}/pages.py batches -w {sh} の束ごとに、asis の references/extraction-guide.md の指示文でサブエージェントに読ませ、"
                     f"{ASIS_CMD}/check_extract.py -w {sh} を FAIL 0 に（references/site-guide.md の 1a）")
            return
        self.row("1a", "共通の資料", "済み", f"資料 {len(m['docs'])}・カード {len(pages)} ページ")

    def applies(self):
        if not self.sites.get("shared"):
            self.row("1b", "書き分け", "―", "共通の資料が無い")
            return
        ap, err = try_load(self.cmp / "applies.json")
        n, first = stage_fails(self.cx, "applies") if ap is not None else (0, "")
        if ap is None:
            self.row("1b", "書き分け", "未", "applies.json が無い")
            self.add(f"python3 scripts/sc_applies.py draft -w {W}    # 共通の資料の書き分けの候補",
                     "# applies.json の未定の項目に、当てはまる拠点を書く（references/site-guide.md の 1b）")
            return
        left = sum(1 for it in ap.get("items") or [] if it.get("scope") == "未定")
        self.row("1b", "書き分け", "要直し" if n else "済み", f"項目 {len(ap.get('items') or [])}・未定 {left}" + (f"・点検 {n} FAIL（{first[:50]}）" if n else ""))
        if n:
            self.add("# applies.json を直す・拠点のフローの外す直しを直す（python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage applies の FAIL を見る）")

    def sites_asis(self):
        bad = 0
        for sid in self.cx.ids:
            w = site_work(self.cmp, self.sites, sid)
            if not (w / "manifest.json").exists():
                self.row("1c", self.names[sid][:8], "未", f"asis の作業フォルダが無い（{w}）")
                ins = ([*self.sites["shared"]["inputs"]] if self.cx.uses_shared(sid) else []) + self.cx.site(sid)["inputs"]
                self.add(f"{ASIS_CMD}/inventory.py {' '.join(ins)} -w {w}    # {self.names[sid]}の棚卸し（共通の資料も入れる）")
                bad += 1
                continue
            m = load(w / "manifest.json")
            if self.cx.uses_shared(sid) and any(not d.get("rendered") for d in m["docs"] if not d.get("duplicate_of")):
                sh = shared_work(self.cmp, self.sites)
                shm = load(sh / "manifest.json", {"docs": []}) if sh else {"docs": []}
                shas = {d["sha1"] for d in shm["docs"]}
                if any(d["sha1"] in shas and not d.get("rendered") for d in m["docs"]):
                    self.row("1c", self.names[sid][:8], "途中", "共通の資料をまだ写していない")
                    self.add(f"python3 scripts/sc_share.py -w {W} --site {sid}    # 共通の資料のページ化とカードを写す（拠点の render_pages.py の前）")
                    bad += 1
                    continue
            gate = None if self.fast else [x for x in __import__("sc_collect").gate(w) if x[1]]
            if gate == []:
                self.row("1c", self.names[sid][:8], "済み", "拠点の門を通る（asis のフローと原典照合が FAIL 0）")
                continue
            line = asis_status(w)
            self.row("1c", self.names[sid][:8], "途中" if gate else "―", f"asis: {line[:70]}" + (f"・門 {sum(x[1] for x in gate)} FAIL" if gate else "（--fast: 門は見ていない）"))
            if gate:
                bad += 1
                self.add(f"# {self.names[sid]}の as-is を進める: {ASIS_CMD}/status.py -w {w} の次の仕事（references/site-guide.md の 1c。共通の資料の書き分けは "
                         f"python3 scripts/sc_site_ops.py -w {W} --site {sid}）")
        if bad:
            self.blocked = True

    def snapshot(self):
        n, first = stage_fails(self.cx, "snapshot")
        have = sum(1 for sid in self.cx.ids if snap_path(self.cmp, sid).exists())
        self.row("2", "写し", "要直し" if n else "済み", f"写し {have}/{len(self.cx.ids)}" + (f"・{first[:60]}" if n else ""))
        if n:
            self.add(f"python3 scripts/sc_collect.py -w {W}    # 拠点の写しを取る（取り直す）")
            self.blocked = True

    def terms(self):
        n, first = stage_fails(self.cx, "terms")
        if not (self.cmp / "terms.json").exists():
            self.row("3", "用語の対応", "未", "terms.json が無い")
            self.add(f"python3 scripts/sc_terms.py draft -w {W}    # 用語の対応の下書き",
                     "# terms.json: 訳・言い換えで同じ用語を、理由つきでまとめ、担当・帳票の組の名前を日本語に（references/terms-guide.md）")
            self.blocked = True
            return
        self.row("3", "用語の対応", "要直し" if n else "済み", f"組 {len((try_load(self.cmp / 'terms.json')[0] or {}).get('groups') or [])}" + (f"・点検 {n} FAIL（{first[:50]}）" if n else ""))
        if n:
            self.add("# terms.json を直す（python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage terms）")
            self.blocked = True

    def axis(self):
        n, first = stage_fails(self.cx, "axis")
        if not (self.cmp / "axis.json").exists():
            self.row("4", "業務種別", "未", "axis.json が無い")
            self.add(f"python3 scripts/sc_axis.py draft -w {W}    # 比べる行の下書き",
                     "# axis.json: ほかの拠点のフローを行に当てる（references/axis-guide.md。材料は sc_axis.py material）")
            self.blocked = True
            return
        ax = try_load(self.cmp / "axis.json")[0] or {}
        self.row("4", "業務種別", "要直し" if n else "済み", f"行 {len(ax.get('rows') or [])}" + (f"・点検 {n} FAIL（{first[:50]}）" if n else ""))
        if n:
            self.add("# axis.json を直す（python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage axis。まだ当てていないフローは sc_axis.py show）")
            self.blocked = True

    def match(self):
        from sc_rows import Model
        self.mdl = mdl = Model(self.cmp, self.sites)
        rows = [r["id"] for r in mdl.rows()]
        miss = [r for r in rows if not match_path(self.cmp, r).exists()]
        n, first = stage_fails(self.cx, "rows")
        self.row("5", "タスク", "未" if len(miss) == len(rows) else "途中" if miss else "要直し" if n else "済み",
                 f"行 {len(rows) - len(miss)}/{len(rows)}" + (f"・点検 {n} FAIL（{first[:50]}）" if n and not miss else ""))
        if miss:
            self.add(*[f"python3 scripts/sc_match.py draft -w {W} --row {r}" for r in miss[:3]],
                     f"# 行ごとにサブエージェントでタスクの対応を書く（references/match-guide.md）: {', '.join(miss[:8])}{' ほか' if len(miss) > 8 else ''}")
            self.blocked = True
        elif n:
            self.add("# タスクの対応を直す（python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage rows。記載なしの候補は sc_match.py candidates）")
            self.blocked = True

    def compute(self):
        n, first = stage_fails(self.cx, "compute")
        self.row("6", "計算", "作り直し" if n else "済み", first[:70] if n else f"行 {len(self.mdl.rows())}")
        if n:
            self.add(f"python3 scripts/sc_compute.py -w {W}    # 並び・区分・差・複雑性・依頼を計算する")
            self.blocked = True

    def verify(self):
        n, first = stage_fails(self.cx, "verify")
        has = (self.cmp / "review" / "versions.json").exists()
        self.row("7", "照合", "未" if not has else "要直し" if n else "済み", first[:80] if n else "支持率 95% 以上")
        if n:
            if "対応の照合をしていない" in first:
                self.add(f"python3 scripts/sc_verify.py sheet -w {W}    # 対応の照合シート（行ごと）",
                         "# 照合シートを、作り方を知らない別のエージェントに渡して照合させる（references/verify-prompt.md）→ python3 scripts/sc_verify.py draft → 採否と対応を書く")
            else:
                self.add(f"# 照合の点検の FAIL を直す（python3 scripts/sc_check.py -w <比較の作業フォルダ> --stage verify。変わった主張は、出てくるコマンドで差分照合）")

    def outputs(self):
        n, first = stage_fails(self.cx, "outputs")
        self.row("8", "書き出し", "作り直し" if n else "済み", first[:80] if n else "拠点比較.xlsx・拠点比較フロー.drawio")
        if n:
            self.add(f"python3 scripts/sc_book.py -w {W} && python3 scripts/sc_drawio.py -w {W} && python3 scripts/sc_layout.py -w {W}    # 書き出しと配置の点検・PNG")

    def gate(self):
        done = (self.cmp / "out" / "比較の数.md").exists()
        self.row("9", "納品の門", "済み" if done else "未", "出力先に写した" if done else "")
        if not done:
            self.add(f"python3 scripts/sc_deliver.py -w {W} --to <出力先>    # 納品の門（出力先を利用者に聞く）")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-w", "--work", required=True)
    ap.add_argument("--fast", action="store_true")
    a = ap.parse_args()
    cmp = work_dir(a.work)
    st = Status(cmp, a.fast)
    st.run()
    print(f"== 進み具合（{cmp.name}）==")
    print("\n".join(st.rows))
    print("\n次にすること:")
    for t in st.todo[:5]:
        print(f"  {t}")


if __name__ == "__main__":
    main()
