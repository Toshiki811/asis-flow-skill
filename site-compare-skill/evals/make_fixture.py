#!/usr/bin/env python3
"""make_fixture.py — 拠点間比較の動作確認に使う、架空の「経費精算」の資料一式（3 拠点＋全社共通）を作る。

  python3 evals/make_fixture.py <出力フォルダ>

できる形（sc_init.py がそのまま読める）:
  共通/経費精算マニュアル（全社共通）.pdf    2 ページ。申請と承認。承認の項に「※ 大阪支店では、課長に代えて支店長が承認する。」（拠点ごとの書き分け）
  東京支店/東京支店_支払処理手順.pdf          3 ページ。支払を詳しく。振込は下位の手順（別の節）。基準の拠点
  大阪支店/大阪支店_支払手順.pdf              1 ページ。支払を 1 文で（粗い）。支払対象の一覧の出力・記帳は書いていない
  シンガポール拠点/SG_Expense_Procedure.pdf   2 ページ。英語。申請・承認を粗く 2 文で。支払に GST の請求書の確認（この拠点だけ）。共通の資料は使わない

仕込んだもの（試験が見る）:
  - 共通の資料の書き分け（大阪支店だけ承認者が違う）。東京支店の図に大阪支店の手順が入る・大阪支店が自分の手順を外す、は点検が FAIL
  - 粒度の差: 承認はシンガポールが 1 文、支払は大阪支店・シンガポールが 1 文 → タスクは粗い拠点に合わせ、東京支店の箱は内訳
  - 親子: 東京支店の支払処理（親）の中の振込（子）。開いて 1 本で比べる
  - 記載なし: 大阪支店に、支払対象の一覧の出力と記帳が無い。拠点固有: シンガポールの GST の請求書の確認
  - 英語の資料（シンガポール）: カードとフローの文は日本語に訳し、引用は原文のまま
"""
import pathlib
import sys

import pymupdf

JP_FONTS = ["/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", "/Library/Fonts/Arial Unicode.ttf", "C:/Windows/Fonts/meiryo.ttc"]

SHARED = [
    ("経費精算マニュアル（全社共通）", [
        ("h", "1. 経費申請"),
        ("t", "経費を立て替えた社員（以下、申請者）は、次の手順で申請する。"),
        ("t", "(1) 申請者は、経費精算システムで経費申請書を作成する。"),
        ("t", "(2) 申請者は、領収書の画像を添付して、経費申請書を提出する。"),
    ]),
    ("", [
        ("h", "2. 承認"),
        ("t", "(1) 課長は、申請の内容と領収書を照合する。"),
        ("t", "(2) 不備がある場合は、課長は申請者に差し戻す。"),
        ("t", "(3) 不備がない場合は、課長が申請を承認する。"),
        ("t", "※ 大阪支店では、課長に代えて支店長が申請を承認する。"),
        ("t", "承認された申請は、各拠点の経理の支払処理に回る。"),
    ]),
]

TOKYO = [
    ("東京支店 支払処理手順", [
        ("h", "1. 支払処理"),
        ("t", "(1) 経理部は、経費精算システムから承認済みの申請を一覧で出力する。"),
        ("t", "(2) 経理部は、会計システムで支払データを作成する。"),
        ("t", "(3) 経理課長は、支払データを確認する。"),
        ("t", "(4) 経理部は、振込を行う。振込の手順は「2. 振込」による。"),
    ]),
    ("", [
        ("h", "2. 振込"),
        ("t", "(1) 経理部は、ファームバンキングで振込を実行する。"),
        ("t", "(2) 振込エラーがある場合は、経理部は口座情報を申請者に確認し、再振込する。"),
        ("t", "(3) 振込エラーがない場合は、「3. 記帳」に進む。"),
    ]),
    ("", [
        ("h", "3. 記帳"),
        ("t", "振込が終わると、会計システムが振込結果を自動で記帳する。"),
    ]),
]

OSAKA = [
    ("大阪支店 支払手順", [
        ("h", "支払"),
        ("t", "経理担当は、承認済みの申請について、支払データを作成して振込を行う。"),
    ]),
]

SINGAPORE = [
    ("Expense Claim Procedure (Singapore)", [
        ("h", "1. Claim and approval"),
        ("t", "The employee submits the expense claim with the receipts."),
        ("t", "The manager reviews the claim and approves it."),
    ]),
    ("", [
        ("h", "2. Payment"),
        ("t", "Finance checks the GST tax invoice for each claim."),
        ("t", "Finance pays the approved claims by bank transfer."),
        ("t", "Finance records the payment in Oracle."),
    ]),
]


def font_file():
    for f in JP_FONTS:
        if pathlib.Path(f).exists():
            return f
    sys.exit("日本語のフォントが見つからない（IPAゴシックなどを入れる）")


def make_pdf(path, pages, jp=True):
    """1 ページ = (題名, [(種類 h|t, 文)])。文は左上から 1 行ずつ。英語は Helvetica、日本語は IPAゴシック"""
    doc = pymupdf.open()
    ff = font_file() if jp else None
    for title, lines in pages:
        pg = doc.new_page(width=595, height=842)
        y = 70
        if title:
            pg.insert_text((60, y), title, fontsize=18, fontname="jp" if jp else "helv", fontfile=ff)
            y += 44
        for kind, text in lines:
            size = 13 if kind == "h" else 11
            pg.insert_text((60 if kind == "h" else 72, y), text, fontsize=size, fontname="jp" if jp else "helv", fontfile=ff)
            y += 30 if kind == "h" else 24
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.set_metadata({"title": path.stem, "creationDate": "D:20260401000000", "modDate": "D:20260401000000", "producer": "make_fixture"})
    doc.save(path, no_new_id=True)                                   # 何度作っても同じ中身（sha1 が変わらない。共通の資料を拠点の作業フォルダで見つけるのに使う）


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    out = pathlib.Path(sys.argv[1]).expanduser().resolve()
    make_pdf(out / "共通" / "経費精算マニュアル（全社共通）.pdf", SHARED)
    make_pdf(out / "東京支店" / "東京支店_支払処理手順.pdf", TOKYO)
    make_pdf(out / "大阪支店" / "大阪支店_支払手順.pdf", OSAKA)
    make_pdf(out / "シンガポール拠点" / "SG_Expense_Procedure.pdf", SINGAPORE, jp=False)
    print(f"入力一式 → {out}")


if __name__ == "__main__":
    main()
