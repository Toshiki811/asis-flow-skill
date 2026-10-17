#!/usr/bin/env python3
"""sc_common.py — 拠点間比較スキルの共通の部品。asis-flow-skill の部品を読み込み、比較の作業フォルダの道の名前・読み書きをまとめる。

比較の作業フォルダ（<CMP>。どのスクリプトにも -w で渡す）:
  sites.json                      拠点の一覧・基準・共通の資料（工程 0。利用者と決める）
  asis/shared/                    共通の資料の asis の作業フォルダ（棚卸し・ページ化・読取だけ）
  asis/<拠点>/                    拠点ごとの asis の作業フォルダ（sites.json の work で別の場所も指せる）
  applies.json                    共通の資料の、拠点ごとの書き分け（工程 1b。機械の候補に LLM が当てはまる拠点を書く）
  snap/<拠点>.json                拠点の写し（工程 2。機械だけが書く）
  terms.json                      拠点間の用語の対応（工程 3）
  axis.json                       比べる行（業務分類 › 業務種別）と、各拠点のフローの当て方（工程 4）
  rows/<行>.match.json            行の中のタスクと、各拠点の箱（工程 5）
  rows/<行>.result.json           計算の結果（工程 6。機械だけ）
  numbers.json                    タスクの通し番号（T-001。機械だけ。番号を引き継ぐ）
  reports/                        サブエージェント・照合の担当の報告
  review/<行>.review.json         採否表（親は採否と対応だけを書く）・review/versions.json（照合に出した版。機械だけ）
  out/                            納品物（拠点比較.xlsx・拠点比較フロー.drawio）と点検の記録

asis の場所: 環境変数 ASIS_FLOW_SKILL → このスキルの 1 つ上のフォルダ → となりの asis-flow-skill（シンボリックリンクを解いて探す）。
モジュール名は sc_ で始める（asis の common・validate などと重ねない）。
"""
import hashlib
import json
import os
import pathlib
import re
import sys

SKILL_DIR = pathlib.Path(__file__).resolve().parent.parent
ASSETS = SKILL_DIR / "assets"


def find_asis():
    """asis-flow-skill のフォルダ（scripts/common.py と SKILL.md があるもの）"""
    cands = []
    if os.environ.get("ASIS_FLOW_SKILL"):
        cands.append(pathlib.Path(os.environ["ASIS_FLOW_SKILL"]).expanduser())
    cands += [SKILL_DIR.parent, SKILL_DIR.parent / "asis-flow-skill"]
    for c in cands:
        c = c.resolve()
        if (c / "scripts" / "common.py").exists() and (c / "SKILL.md").exists() and "name: asis-flow-skill" in (c / "SKILL.md").read_text(encoding="utf-8")[:300]:
            return c
    sys.exit("asis-flow-skill が見つからない（環境変数 ASIS_FLOW_SKILL に、asis-flow-skill のフォルダを書く）")


ASIS = find_asis()
ASIS_SCRIPTS = ASIS / "scripts"
if str(ASIS_SCRIPTS) not in sys.path:
    sys.path.insert(1, str(ASIS_SCRIPTS))

import common as A                                                   # noqa: E402  asis の共通の部品

Report, load, save, try_load, norm, same_text = A.Report, A.load, A.save, A.try_load, A.norm, A.same_text
ends_with_verb, ends_as_question, CJK, site = A.ends_with_verb, A.ends_as_question, A.CJK, A.site

SITE_ID = re.compile(r"^[A-Z][A-Z0-9]{1,5}$")
STEP_TYPES = ("process", "decision", "system")                       # タスクに入れる箱（別のフローへ進む箱・開始・終了・帳票は入れない）
OTHER_SITE = "ほかの拠点だけの手順"                                     # 共通の資料の、ほかの拠点だけの手順を外す理由の頭（決まった形）
BOOK = "拠点比較.xlsx"
FIG = "拠点比較フロー.drawio"
MATERIAL_CAP = A.MATERIAL_CAP                                         # 1 体に渡す材料の上限（字）。asis と同じ
REVIEW_MIN = A.REVIEW_MIN                                             # 照合の支持率の下限


def work_dir(arg):
    w = pathlib.Path(arg).expanduser().resolve()
    if not w.exists():
        sys.exit(f"比較の作業フォルダが無い: {w}（先に python3 scripts/sc_init.py を実行する）")
    return w


def sites_of(cmp):
    """sites.json（無ければ止まる）"""
    s, err = try_load(cmp / "sites.json")
    if err or not isinstance(s, dict):
        sys.exit(f"sites.json が読めない（{err or '無い'}）。python3 scripts/sc_init.py で下書きを作る")
    return s


def site_ids(sites):
    return [x["id"] for x in sites.get("sites") or [] if isinstance(x, dict) and "id" in x]


def site_names(sites):
    return {x["id"]: x.get("name", x["id"]) for x in sites.get("sites") or [] if isinstance(x, dict) and "id" in x}


def ordered_sites(sites):
    """表・図に並べる順: 基準の拠点を先頭に、あとは sites.json の順"""
    ids = site_ids(sites)
    b = sites.get("baseline")
    return ([b] if b in ids else []) + [i for i in ids if i != b]


def site_work(cmp, sites, sid):
    """拠点の asis の作業フォルダ"""
    for x in sites.get("sites") or []:
        if x.get("id") == sid:
            return pathlib.Path(x.get("work") or cmp / "asis" / sid).expanduser().resolve()
    raise KeyError(sid)


def shared_work(cmp, sites):
    sh = sites.get("shared")
    if not sh:
        return None
    return pathlib.Path(sh.get("work") or cmp / "asis" / "shared").expanduser().resolve()


def snap_path(cmp, sid):
    return cmp / "snap" / f"{sid}.json"


def load_snap(cmp, sid):
    s, err = try_load(snap_path(cmp, sid))
    return None if err or not isinstance(s, dict) else s


def row_ids(axis):
    return [r["id"] for r in (axis or {}).get("rows") or [] if isinstance(r, dict) and "id" in r]


def match_path(cmp, rid):
    return cmp / "rows" / f"{rid}.match.json"


def result_path(cmp, rid):
    return cmp / "rows" / f"{rid}.result.json"


def digest(obj, n=12):
    """中身の印（16 進）。並びが変わらなければ同じ"""
    return hashlib.sha1(json.dumps(obj, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:n]


def node_key(flow, nid):
    """拠点の中の箱の鍵「F03-1:P-004」（フロー ID と、そのフローの中の番号）"""
    return f"{flow}:{nid}"


def split_key(key):
    f, _, n = str(key).partition(":")
    return f, n


def mentions(text, sites):
    """文に出てくる拠点（名前・別名）: [拠点 ID]"""
    t = norm(str(text or ""))
    out = []
    for x in sites.get("sites") or []:
        names = [x.get("name", "")] + list(x.get("aliases") or [])
        if any(n and norm(n) in t for n in names):
            out.append(x["id"])
    return out


SWITCH_WORDS = ("拠点ごと", "拠点によって", "支店ごと", "支店によって", "各拠点", "各支店", "拠点により", "支店により", "by site", "by branch", "each branch", "each site")


def is_jp(text):
    """日本語の字がある（納品物に出す文は、日本語だけ）"""
    return bool(CJK.search(str(text or "")))


def bigrams(text):
    """似ている度合いを見る字の組: 日本語は、かな（助詞・送りがな）と記号を除いた字の 2 字ずつ。日本語の字の無い文は、語（小文字）"""
    t = norm(str(text or ""))
    if not is_jp(t):
        return {w for w in re.findall(r"[a-z0-9]+", t.lower()) if len(w) > 2}
    t = re.sub(r"[\sぁ-ゖ、。，．・（）()「」『』:：/／\-ー]+", "", t)
    return {t[i:i + 2] for i in range(len(t) - 1)} or ({t} if t else set())


def likeness(a, b):
    """2 つの文の似ている度合い（2 文字の組の重なり。0〜1）。記載なしの候補を拾うのに使う"""
    x, y = bigrams(a), bigrams(b)
    if not x or not y:
        return 0.0
    return len(x & y) / min(len(x), len(y))
