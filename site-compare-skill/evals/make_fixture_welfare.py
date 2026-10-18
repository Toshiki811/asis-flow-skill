#!/usr/bin/env python3
"""make_fixture_welfare.py — 架空の社会福祉法人（3 拠点）の経理事務の資料を作る（通しの試験の入力。LLM は使わない）。

  python3 evals/make_fixture_welfare.py <出力フォルダ>            # <出力>/共通/・本部拠点/・特別養護老人ホーム拠点/・保育所拠点/ に資料を書く
  python3 evals/make_fixture_welfare.py <出力フォルダ> --check    # もう一度作って、全ファイルの sha1 が同じことを確かめる（決定性）

資料（全部日本語。もとは「社会福祉法人 経理事務マニュアル」の 8 業務分類・19 業務）:
  共通/法人経理規程・経理事務マニュアル.docx     法人全体の決まり。19 業務の手順を規程の粒度で。拠点ごとの書き分け「※ 保育所拠点では…」が 8 か所
  本部拠点/本部_経理事務手順書.pdf               詳しい手順（担当・判断・システム・帳票つき）。表紙と目次。刷られたページ番号は表紙の分 1 ずれる。
                                                  支払処理に下位の手順「振込実行の手順」（100 万円以上は理事長の承認）
  特別養護老人ホーム拠点/特養_経理事務手順表.xlsx  業務分類ごとに 1 シートの手順表（手順・担当・システム・帳票・備考）。粗い。介護報酬の国保連への伝送が固有
  保育所拠点/保育所_経理の流れ.pptx              業務ごとに 1 枚のフロー図（図形と線）。いちばん粗い。小口現金・契約・給与計算・決算は本部

仕込んだ差（期待される答え）は evals/welfare/expected.md。
決定性: docx・xlsx・pptx は、中の日時を固定し、zip の項目の日時もそろえる（pin_zip）。PDF は set_metadata と no_new_id。
同じ入力を作り直しても sha1 が同じ（共通の資料は sha1 で突き合わせる）。
字体: LibreOffice の日本語は、字体の名前を指さないと中国語の字形に落ちるので、IPAPGothic を指す（無い環境では、その環境の既定の字体に落ちる）。
"""
import argparse
import datetime
import hashlib
import io
import pathlib
import sys
import tempfile
import zipfile

JP_FONT = "IPAPGothic"
STAMP = datetime.datetime(2026, 4, 1, 0, 0, 0)
SITES = {"shared": "共通", "HQ": "本部拠点", "TY": "特別養護老人ホーム拠点", "HO": "保育所拠点"}

# ---------- 共通の資料（規程の粒度。拠点ごとの書き分けは「※」） ----------

SHARED = [
    ("経理体制・帳簿管理", [
        ("経理体制整備", ["理事長は、会計責任者、出納職員、予算管理責任者および固定資産管理責任者を任命する。",
                   "会計責任者は、出納職員を監督し、拠点区分ごとの担当範囲を確認する。"], []),
        ("帳簿・帳票管理", ["会計責任者は、仕訳日記帳、総勘定元帳、小口現金出納帳、固定資産管理台帳その他の帳簿を整備する。",
                   "出納職員は、会計システムの勘定科目、サービス区分および帳票の出力設定を確認する。"], []),
    ]),
    ("予算管理", [
        ("当初予算編成", ["予算管理責任者は、次年度の事業計画、現場からの聞き取りおよび承認済みの社会福祉充実計画を収集する。",
                  "予算管理責任者は、勘定科目ごとに収入予算と支出予算を積算し、法人全体および拠点区分ごとの資金収支予算案を作成する。",
                  "理事長は、予算案について理事会の承認を得て、予算を確定する。"],
         ["※ 特別養護老人ホーム拠点および保育所拠点は、(1) の収集のあと、予算要求書を本部拠点へ提出する。(2) の予算案の作成および (3) の理事会承認は、本部拠点が行う。"]),
        ("補正予算管理", ["予算管理責任者は、予算の執行中に、予算額との乖離や変更事由がないかを確認する。",
                  "変更事由がある場合、予算管理責任者は補正予算案を作成し、理事長が理事会へ提出する。"],
         ["※ 特別養護老人ホーム拠点および保育所拠点は、(1) の確認で乖離があれば本部拠点へ連絡する。(2) の補正予算案の作成は、本部拠点が行う。"]),
    ]),
    ("収納事務", [
        ("請求・収納", ["出納職員は、補助金、利用料、サービス報酬などの請求データを作成する。",
                 "会計責任者は、請求額、利用者負担および補助金の内容を確認する。",
                 "出納職員は、請求書または請求書兼領収書を発行し、利用者および関係機関へ送付する。"], []),
        ("現金収納", ["出納職員は、現金を直接受け取った場合、収納の内容を確認する。",
                "出納職員は、福祉サービス利用者負担金の収納時に、領収書を発行する。",
                "出納職員は、収納した現金を、経理規程に定める期日までに金融機関へ預け入れる。"],
         ["※ 特別養護老人ホーム拠点では、利用者負担金の現金収納時の領収書は、出納職員に代えて生活相談員が発行する。"]),
        ("会計伝票作成", ["出納職員は、入金を確認したのち、会計システムで収納に係る会計伝票を作成する。",
                  "出納職員は、領収書の控え、入金明細および会計伝票を整理して保管する。"], []),
    ]),
    ("支払・契約事務", [
        ("支払処理", ["担当者は、物品の購入や修繕工事などの発注前に、稟議書または購入依頼書で起案する。",
                "出納職員は、請求書を、契約書、稟議書、購入依頼書および納品書と照合し、必要事項を確認する。",
                "会計責任者は、出納職員が作成した支払情報を請求書と照合して承認する。",
                "出納職員は、承認済みの支払情報をもとに、総合振込またはインターネットバンキングで支払を実行する。",
                "出納職員は、会計システムで支払に係る会計伝票を作成する。"],
         ["※ 特別養護老人ホーム拠点および保育所拠点では、(1) の起案および (2) の請求書の照合を施設で行い、確認済みの請求書を本部拠点へ送付する。(3) の承認から (5) の伝票の作成までは、本部拠点が行う。"]),
        ("契約事務", ["契約担当者は、入札契約か随意契約か、契約金額と内容に応じた手続きを確認する。",
                "契約担当者は、随意契約では原則として 3 社以上の見積を徴取し、比較する。",
                "契約担当者は、契約先の決定後、標準契約書または業者の指定様式で契約書を作成する。",
                "契約担当者は、納品書、納品物および業務報告書で契約の履行を確認する。"],
         ["※ 契約事務は、本部拠点の契約担当者が法人全体をまとめて行う。特別養護老人ホーム拠点および保育所拠点では行わない。"]),
    ]),
    ("現金管理", [
        ("小口現金支払", ["出納職員は、小口現金から支払う申請および領収書を受け付ける。",
                  "出納職員は、領収書、支払内容および承認の有無を確認する。",
                  "出納職員は、小口現金出納帳に支払内容を記録し、証憑を保管する。"],
         ["※ 保育所拠点は、小口現金を置かない。保育所拠点の小口の支払は、本部拠点の小口現金から行う。"]),
        ("現金残高確認", ["出納職員は、金庫内の現金を数え、小口現金出納帳および会計帳簿と照合する。",
                  "残高に差異がある場合、出納職員は原因を調査し、会計責任者へ報告する。"], []),
    ]),
    ("人件費支払", [
        ("給与支払", ["出納職員は、日報および勤怠情報を収集し、給与計算の元データを整える。",
                "出納職員は、給与計算システムに勤怠、手当および控除の情報を入力する。",
                "会計責任者は、給与計算データを確認し、承認する。",
                "出納職員は、承認済みの給与データをもとに給与を支払う。",
                "出納職員は、給与支払時の会計仕訳を会計システムに入力する。"],
         ["※ (2) の給与計算の入力から (5) の仕訳の入力まで、ならびに賞与・年末調整および社会保険・源泉の事務は、本部拠点が法人全体をまとめて行う。"
          "特別養護老人ホーム拠点および保育所拠点は、(1) の勤怠情報の収集を行い、本部拠点へ送付する。"]),
        ("賞与・年末調整", ["出納職員は、賞与の支給額を計算し、給与計算システムへ入力する。",
                   "出納職員は、年末調整に必要な申告書および控除の資料を確認し、計算する。"], []),
        ("社会保険・源泉", ["出納職員は、源泉徴収の事務、算定基礎届および賞与支払届を作成し、提出する。"], []),
    ]),
    ("資産・負債管理", [
        ("債権債務管理", ["出納職員は、毎月末日の未収金および未払金など、債権債務の残高内訳を調査する。",
                  "出納職員は、必要に応じて取引先と残高を確認し、差異を調査する。",
                  "会計責任者は、残高の差異や不明な残高について報告し、必要な措置を行う。"], []),
        ("実地棚卸", ["会計責任者は、棚卸の方法を周知し、棚卸表を担当職員に回付する。",
                "担当職員は、品名、品番および数量を棚卸表に記入し、帳簿にない棚卸資産も確認する。",
                "会計責任者は、棚卸表を回収し集計のうえ、承認する。"], []),
        ("固定資産管理", ["固定資産管理責任者は、取得した固定資産を固定資産管理台帳へ登録する。",
                  "固定資産管理責任者は、年度末に固定資産管理台帳と現物の存在および使用状況を確認する。",
                  "固定資産管理責任者は、固定資産の廃棄または売却時に所定の承認を得て、台帳へ反映する。"],
         ["※ 保育所拠点の固定資産の除却は、本部拠点の固定資産管理責任者が承認する。"]),
    ]),
    ("決算・報告", [
        ("月次報告", ["出納職員は、現金、預金、仮払金および仮受金など、月次処理に必要な残高を確認する。",
                "出納職員は、会計システムで月次試算表を作成し、内容を確認する。",
                "会計責任者は、経理規程で定める期日までに、理事長へ月次報告を行う。"],
         ["※ 保育所拠点の月次報告は、施設長を経由して会計責任者へ提出する。", "※ 本部拠点では、理事長に代えて常務理事が月次報告を受ける。"]),
        ("年次決算", ["出納職員は、未収・未払、棚卸、減価償却、引当金および見越繰延など、決算整理仕訳を作成する。",
                "出納職員は、作成した決算整理仕訳を会計システムへ入力する。",
                "会計責任者は、貸借対照表、資金収支計算書、事業活動計算書、注記および附属明細書を作成する。",
                "収益事業または消費税の課税がある場合、会計責任者は申告書を作成し、申告および納税を行う。"],
         ["※ 年次決算は、本部拠点が法人全体をまとめて行う。特別養護老人ホーム拠点および保育所拠点では行わない。"]),
    ]),
]

# ---------- 本部拠点の手順書（詳しい。1 業務 1 ページ） ----------
# 行: ("step", 担当, 文, システム, 帳票) / ("decision", 問い, はいの行き先, いいえの行き先) / ("note", 文)

HQ = [
    ("当初予算編成", "次年度の予算を編成し、理事会の承認を得るまでの手順。", [
        ("step", "予算管理責任者", "各拠点から提出された予算要求書と、次年度の事業計画および承認済みの社会福祉充実計画を収集する。", "", "予算要求書"),
        ("step", "予算管理責任者", "予算要求書の内容を、各拠点の施設長に聞き取って確認する。", "", ""),
        ("step", "予算管理責任者", "会計システムの予算管理表に、勘定科目ごとの収入予算と支出予算を入力し、法人全体と拠点区分ごとの資金収支予算案を作成する。", "会計システム", "予算管理表"),
        ("step", "会計責任者", "予算案の内容を確認し、理事長に提出する。", "", ""),
        ("step", "理事長", "理事会に予算案を提出し、承認を得て予算を確定する。", "", ""),
    ]),
    ("補正予算管理", "予算の執行中に乖離を確かめ、必要なら補正予算を作る手順。", [
        ("step", "予算管理責任者", "毎月、会計システムの予算管理表で、予算額と執行額を比較する。", "会計システム", "予算管理表"),
        ("decision", "乖離が予算額の 10% 以上か", "(2) へ", "確認を終える"),
        ("step", "予算管理責任者", "乖離の原因を所管の施設長に確認し、変更事由を整理する。", "", ""),
        ("step", "予算管理責任者", "補正予算案を作成し、会計責任者の確認を受ける。", "", "補正予算案"),
        ("step", "理事長", "補正予算案を理事会へ提出し、承認を得る。", "", ""),
    ]),
    ("請求・収納", "補助金・利用料・サービス報酬を請求する手順。", [
        ("step", "出納職員", "毎月 5 日までに、請求システムで補助金、利用料およびサービス報酬の請求データを作成する。", "請求システム", "請求データ"),
        ("step", "会計責任者", "請求データの請求額、利用者負担額および補助金の内訳を確認する。", "", ""),
        ("decision", "請求データに不備があるか", "出納職員へ差し戻し、(1) をやり直す", "(3) へ"),
        ("step", "出納職員", "請求システムで請求書を発行し、利用者および関係機関へ送付する。", "請求システム", "請求書"),
    ]),
    ("現金収納", "窓口で現金を受け取ったときの手順。", [
        ("step", "出納職員", "窓口で現金を受け取ったとき、金額と収納の内容を確認する。", "", ""),
        ("step", "出納職員", "利用者負担金の収納時に、領収書を発行して渡す。", "", "領収書"),
        ("step", "出納職員", "収納した現金を、当日または翌営業日に金融機関へ預け入れる。", "", "入金明細"),
    ]),
    ("会計伝票作成", "入金を確かめて収納伝票を作り、証憑を保管する手順。", [
        ("step", "出納職員", "預金通帳の記帳で入金を確認したのち、会計システムで収納伝票を作成する。", "会計システム", "収納伝票"),
        ("step", "会計責任者", "収納伝票を承認する。", "会計システム", ""),
        ("step", "出納職員", "領収書の控え、入金明細および収納伝票を、月ごとに整理して保管する。", "", ""),
    ]),
    ("支払処理", "起案から支払伝票の作成までの手順。振込の実行は「振込実行の手順」（次項）による。", [
        ("step", "担当者", "物品の購入や修繕工事の発注前に、稟議書または購入依頼書で起案する。", "", "稟議書・購入依頼書"),
        ("step", "会計責任者", "稟議書または購入依頼書を承認する。", "", ""),
        ("step", "出納職員", "届いた請求書を、契約書、稟議書、購入依頼書および納品書と照合し、支払期日と振込先を確認する。", "", "請求書・納品書"),
        ("decision", "照合に不一致があるか", "担当者に確認を依頼し、(3) をやり直す", "(4) へ"),
        ("step", "出納職員", "会計システムで支払情報を作成する。", "会計システム", "支払情報"),
        ("step", "会計責任者", "支払情報を請求書と照合して承認する。", "会計システム", ""),
        ("step", "出納職員", "「振込実行の手順」に従って振込を実行する。", "", ""),
        ("step", "出納職員", "会計システムで支払伝票を作成する。", "会計システム", "支払伝票"),
    ]),
    ("振込実行の手順", "支払処理の (7) の振込を実行する手順（支払処理の下位の手順）。", [
        ("step", "出納職員", "承認済みの支払情報から、インターネットバンキングで総合振込のデータを作成する。", "インターネットバンキング", "総合振込データ"),
        ("decision", "1 件の振込額が 100 万円以上か", "(2) へ", "(3) へ"),
        ("step", "理事長", "振込の内容を確認し、承認する。", "", ""),
        ("step", "会計責任者", "インターネットバンキングで振込を承認し、送信する。", "インターネットバンキング", ""),
        ("step", "出納職員", "振込の結果を確認し、エラーがあれば振込先に口座情報を確認して、再度振込を行う。", "インターネットバンキング", ""),
    ]),
    ("契約事務", "契約の方式を決め、契約書を作り、履行を確かめる手順。", [
        ("step", "契約担当者", "稟議書の内容から、契約の方式と予定価格を確認する。", "", "稟議書"),
        ("decision", "予定価格が 500 万円以上か", "入札とし、(2) を飛ばして (3) へ", "(2) へ"),
        ("step", "契約担当者", "原則として 3 社以上から見積書を徴取し、見積比較表を作成する。", "", "見積書・見積比較表"),
        ("step", "契約担当者", "契約先の決定後、標準契約書または業者の指定様式で契約書を作成し、理事長の決裁を受ける。", "", "契約書"),
        ("step", "契約担当者", "納品書、納品物および業務報告書で、契約の履行を確認する。", "", "納品書・業務報告書"),
    ]),
    ("小口現金支払", "小口現金から支払うときの手順。", [
        ("step", "出納職員", "小口現金から支払う申請書と領収書を受け付ける。", "", "小口現金支払申請書・領収書"),
        ("step", "出納職員", "領収書の宛名と金額、支払の内容および所属長の承認の有無を確認する。", "", ""),
        ("step", "出納職員", "小口現金出納帳に支払の内容を記録し、領収書を貼付して保管する。", "", "小口現金出納帳"),
        ("step", "出納職員", "小口現金の残高が 5 万円を下回ったとき、会計責任者の承認を得て補充する。", "", ""),
    ]),
    ("現金残高確認", "金庫の現金を実査する手順。", [
        ("step", "出納職員", "月 1 回、金庫内の現金を数え、小口現金出納帳および会計システムの帳簿残高と照合する。", "会計システム", "現金実査表"),
        ("decision", "残高に差異があるか", "(2) へ", "現金実査表に署名して終える"),
        ("step", "出納職員", "差異の原因を調査し、会計責任者へ報告する。", "", ""),
        ("step", "会計責任者", "差異の処理を決め、理事長へ報告する。", "", ""),
    ]),
    ("給与支払", "勤怠の収集から給与の仕訳までの手順。", [
        ("step", "出納職員", "各拠点から送付された勤怠表と本部の日報を収集し、給与計算の元データを整える。", "", "勤怠表"),
        ("step", "出納職員", "給与計算システムに、勤怠、手当および控除の情報を入力する。", "給与計算システム", ""),
        ("step", "会計責任者", "給与計算システムの計算結果を確認し、承認する。", "給与計算システム", ""),
        ("step", "出納職員", "承認済みの給与データを、インターネットバンキングで振り込む。", "インターネットバンキング", ""),
        ("step", "出納職員", "給与支払の仕訳を会計システムに入力する。", "会計システム", "給与仕訳伝票"),
    ]),
    ("賞与・年末調整", "賞与の計算と年末調整の手順。", [
        ("step", "出納職員", "理事長が決めた支給率をもとに賞与の支給額を計算し、給与計算システムへ入力する。", "給与計算システム", ""),
        ("step", "会計責任者", "賞与の計算結果を確認し、承認する。", "給与計算システム", ""),
        ("step", "出納職員", "年末調整の申告書および控除証明書を職員から回収し、給与計算システムで年末調整を計算する。", "給与計算システム", "年末調整申告書"),
    ]),
    ("社会保険・源泉", "源泉所得税と社会保険の届出の手順。", [
        ("step", "出納職員", "毎月の源泉所得税を計算し、翌月 10 日までに納付する。", "給与計算システム", ""),
        ("step", "出納職員", "算定基礎届および賞与支払届を作成し、年金事務所へ提出する。", "給与計算システム", "算定基礎届・賞与支払届"),
    ]),
    ("債権債務管理", "未収金・未払金の残高を調べる手順。", [
        ("step", "出納職員", "毎月末日の未収金および未払金の残高内訳表を、会計システムから出力する。", "会計システム", "残高内訳表"),
        ("step", "出納職員", "3 か月以上残っている未収金について、取引先に残高を確認する。", "", ""),
        ("step", "会計責任者", "差異や不明な残高について理事長へ報告し、回収または処理の方針を決める。", "", ""),
    ]),
    ("実地棚卸", "年度末の棚卸の手順。", [
        ("step", "会計責任者", "年度末に棚卸の方法を周知し、棚卸表を各拠点の担当職員に回付する。", "", "棚卸表"),
        ("step", "担当職員", "品名、品番および数量を棚卸表に記入し、帳簿にない棚卸資産も確認する。", "", "棚卸表"),
        ("step", "会計責任者", "棚卸表を回収して集計し、承認する。", "", "棚卸表"),
    ]),
    ("固定資産管理", "固定資産の登録・現物確認・除却の手順。", [
        ("step", "固定資産管理責任者", "取得価額 10 万円以上の資産を、固定資産管理台帳へ登録する。", "", "固定資産管理台帳"),
        ("step", "固定資産管理責任者", "年度末に、固定資産管理台帳と現物を照合する。", "", "固定資産管理台帳"),
        ("step", "固定資産管理責任者", "除却の申請を受けて理事長の承認を得て、台帳へ反映する。", "", "除却申請書"),
    ]),
    ("月次報告", "月次試算表を作り、常務理事へ報告する手順。", [
        ("step", "出納職員", "月初に、現金、預金、仮払金および仮受金の残高を確認する。", "", ""),
        ("step", "出納職員", "会計システムで月次試算表を作成し、内容を確認する。", "会計システム", "月次試算表"),
        ("step", "会計責任者", "月次試算表を確認し、翌月 15 日までに常務理事へ月次報告を行う。", "", "月次試算表"),
    ]),
    ("年次決算", "決算整理から申告・納税までの手順。", [
        ("step", "出納職員", "未収・未払、棚卸、減価償却、引当金および見越繰延の決算整理仕訳を作成する。", "", "決算整理仕訳"),
        ("step", "出納職員", "決算整理仕訳を会計システムへ入力する。", "会計システム", ""),
        ("step", "会計責任者", "計算書類（貸借対照表、資金収支計算書、事業活動計算書、注記および附属明細書）を作成する。", "会計システム", "計算書類"),
        ("decision", "収益事業または消費税の課税があるか", "(4) へ", "終える"),
        ("step", "会計責任者", "申告書を作成し、申告および納税を行う。", "", "申告書"),
    ]),
]

# ---------- 特養拠点の手順表（粗い。分類ごとに 1 シート） ----------

TY = [
    ("経理体制・帳簿管理", [
        ("—", "法人の経理規程・経理事務マニュアルのとおり。特養拠点での独自の手順はない。", "", "", "", ""),
    ]),
    ("予算管理", [
        ("1", "事業計画の情報を集め、次年度の予算要求をまとめる", "施設長・事務員", "", "", "現場の職員から聞き取る"),
        ("2", "予算要求書を本部へ提出する", "事務員", "", "予算要求書", "11 月末まで"),
        ("3", "予算の執行状況を予算管理表で確認し、乖離があれば本部へ連絡する", "施設長", "会計システム", "予算管理表", "補正予算案は本部が作る"),
    ]),
    ("収納事務", [
        ("1", "介護記録システムから当月のサービス実績を出力する", "事務員", "介護記録システム", "実績一覧", "介護報酬の請求"),
        ("2", "介護請求ソフトで請求データを作成し、内容を確認する", "事務員", "介護請求ソフト", "請求データ", ""),
        ("3", "国保連へ請求データを伝送する", "事務員", "介護請求ソフト", "", "毎月 10 日まで"),
        ("4", "利用者負担金の請求書を発行し、家族へ送付する", "事務員", "介護請求ソフト", "請求書", ""),
        ("5", "窓口で現金を受け取り、内容を確認する", "事務員", "", "", "現金収納"),
        ("6", "領収書を発行して渡す", "生活相談員", "", "領収書", ""),
        ("7", "現金を週 1 回まとめて金融機関に預け入れる", "事務員", "", "", ""),
        ("8", "領収書の控えと入金明細を綴る", "事務員", "", "", "会計伝票"),
        ("9", "月末に、会計システムで収納伝票をまとめて入力する", "事務員", "会計システム", "収納伝票", ""),
    ]),
    ("支払・契約事務", [
        ("1", "購入依頼書で起案し、施設長の承認を受ける", "事務員・施設長", "", "購入依頼書", ""),
        ("2", "届いた請求書を納品書と照合し、施設長が確認して本部へ送付する", "事務員・施設長", "", "請求書・納品書", "支払情報の作成と振込は本部"),
        ("—", "契約は本部が行う（施設では行わない）", "", "", "", "契約事務"),
    ]),
    ("現金管理", [
        ("1", "小口現金の支払申請を受け付け、領収書と承認を確認する", "事務員", "", "領収書", "小口現金"),
        ("2", "小口現金出納帳に記入し、領収書を保管する", "事務員", "", "小口現金出納帳", ""),
        ("3", "毎日、終業時に金庫の現金を数え、小口現金出納帳と照合する", "施設長", "", "小口現金出納帳", "現金残高確認"),
        ("4", "差異があれば本部の会計責任者へ報告する", "施設長", "", "", ""),
    ]),
    ("人件費支払", [
        ("1", "介護記録システムから勤怠データを出力し、勤怠表として本部へ送付する", "事務員", "介護記録システム", "勤怠表", "給与計算以降は本部"),
    ]),
    ("資産・負債管理", [
        ("1", "月末に、未収の利用者負担金の内訳を本部へ報告する", "事務員", "介護請求ソフト", "未収一覧", "債権管理"),
        ("2", "棚卸表を受け取り、介護用品と食材の棚卸をする", "事務員・介護職員", "", "棚卸表", "実地棚卸"),
        ("3", "棚卸表を施設長が確認して本部へ送る", "施設長", "", "棚卸表", ""),
        ("4", "年度末に固定資産の現物を確認し、台帳との差を本部へ報告する", "事務員", "", "固定資産管理台帳", "固定資産"),
        ("5", "除却する資産は、施設長の承認を得て本部へ申請する", "事務員・施設長", "", "除却申請書", "台帳の登録・反映は本部"),
    ]),
    ("決算・報告", [
        ("1", "月初に現金と預金の残高を確認する", "事務員", "", "", "月次報告"),
        ("2", "会計システムで月次試算表を作成し、施設長が確認する", "事務員・施設長", "会計システム", "月次試算表", ""),
        ("3", "本部の会計責任者へ月次試算表を送付する", "事務員", "", "月次試算表", "決算は本部"),
    ]),
]

# ---------- 保育所拠点のフロー図（いちばん粗い。1 業務 1 枚） ----------
# 箱: ("start"|"end"|"proc"|"dec", 文, 担当)。線: (from, to, ラベル)

HO_TEXT = [
    ("はじめに", ["保育所拠点の経理は、法人の経理規程・経理事務マニュアルに従う。この資料は、保育所で行う部分の流れをまとめたもの。",
              "小口現金は置かない（小口の支払は本部に申請する）。契約、給与計算、年次決算は本部が行う。"]),
]
HO_FLOWS = [
    ("予算", [("start", "開始", ""), ("proc", "事業計画の情報を集める", "事務担当"), ("proc", "予算要求書を作成する", "園長"), ("proc", "予算要求書を本部へ提出する", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, "")]),
    ("保育料の収納", [("start", "開始", ""), ("proc", "保育料が市町村から口座に入金される", "市町村"), ("proc", "入金明細で入金を確認する", "事務担当"),
                ("proc", "会計システムで収納伝票を入力する", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, "")]),
    ("延長保育料の収納", [("start", "開始", ""), ("proc", "延長保育料の請求書を保護者へ渡す", "事務担当"), ("proc", "現金を受け取る", "事務担当"),
                  ("proc", "領収書を渡す", "事務担当"), ("proc", "翌日までに金融機関へ預け入れる", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, ""), (4, 5, "")]),
    ("物品の購入と支払", [("start", "開始", ""), ("proc", "購入依頼書を作る", "事務担当"), ("proc", "購入依頼書を承認する", "園長"),
                  ("proc", "届いた請求書を確認して本部へ回付する", "事務担当"), ("end", "終了（支払は本部）", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, "")]),
    ("現金の管理", [("start", "開始", ""), ("proc", "週 1 回、金庫の現金を数える", "事務担当"), ("dec", "帳簿と合っているか", ""),
              ("proc", "本部へ報告する", ""), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 4, "はい"), (2, 3, "いいえ"), (3, 4, "")]),
    ("給与", [("start", "開始", ""), ("proc", "勤怠表を作って本部へ送る", "事務担当"), ("end", "終了（給与計算と支払は本部）", "")],
     [(0, 1, ""), (1, 2, "")]),
    ("棚卸", [("start", "開始", ""), ("proc", "棚卸表に記入する", "保育士"), ("proc", "棚卸表を確認する", "園長"), ("proc", "棚卸表を本部へ送る", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, "")]),
    ("固定資産", [("start", "開始", ""), ("proc", "年度末に固定資産の現物を確認する", "事務担当"), ("dec", "使わなくなった資産があるか", ""),
             ("proc", "本部へ除却を申請する（本部の固定資産管理責任者が承認）", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, "はい"), (2, 4, "いいえ"), (3, 4, "")]),
    ("月次報告", [("start", "開始", ""), ("proc", "現金と預金の残高を確認する", "事務担当"), ("proc", "会計システムで月次試算表を作る", "事務担当"),
             ("proc", "月次試算表を確認する", "園長"), ("proc", "本部の会計責任者へ月次試算表を提出する", "事務担当"), ("end", "終了", "")],
     [(0, 1, ""), (1, 2, ""), (2, 3, ""), (3, 4, ""), (4, 5, "")]),
]


# ---------- 決定性 ----------

def pin_zip(path):
    """zip（docx・xlsx・pptx）の項目の日時をそろえ、同じ順に書き直す（python-docx・openpyxl・python-pptx は「いま」の日時を書く）。
    openpyxl は保存のたびに docProps/core.xml の modified を「いま」に書き換えるので、そこも固定する"""
    import re
    path = pathlib.Path(path)
    with zipfile.ZipFile(path) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items:
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(?:modified|created)[^>]*>)[^<]*<", rb"\g<1>2026-04-01T00:00:00Z<", data)
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, data)
    path.write_bytes(buf.getvalue())


# ---------- docx ----------

def make_docx(path):
    from docx import Document
    from docx.enum.text import WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    doc = Document()
    for name in ("Normal", "Title", "Heading 1", "Heading 2", "Heading 3"):
        st = doc.styles[name]
        st.font.name = JP_FONT
        rpr = st.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.append(fonts)
        for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if fonts.get(qn(k)) is not None:
                del fonts.attrib[qn(k)]
        for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(k), JP_FONT)
    doc.styles["Normal"].font.size = Pt(10.5)
    sec = doc.sections[0]
    foot = sec.footer.paragraphs[0]
    foot.alignment = 1
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    foot._p.append(fld)

    doc.add_heading("社会福祉法人 ○○会　経理規程・経理事務マニュアル", 0)
    doc.add_paragraph("この規程およびマニュアルは、社会福祉法人 ○○会（以下「法人」）の経理事務について、法人全体に共通する決まりと手順を定める。")
    doc.add_paragraph("法人の拠点区分は、本部拠点、特別養護老人ホーム拠点および保育所拠点の 3 つである。拠点によって手順が異なる所は、各項の「※」に書く。"
                      "各拠点の詳しい手順は、拠点ごとの手順書による。")
    doc.add_paragraph("主なシステム・帳票: 会計システム、給与計算システム、請求システム、仕訳日記帳、総勘定元帳、小口現金出納帳、固定資産管理台帳、月次試算表、予算管理表。")
    for ci, (cat, bizs) in enumerate(SHARED, 1):
        h = doc.add_heading(f"第{ci}章　{cat}", 1)
        h.paragraph_format.page_break_before = True
        for bi, (biz, steps, notes) in enumerate(bizs, 1):
            doc.add_heading(f"{ci}.{bi}　{biz}", 2)
            for si, s in enumerate(steps, 1):
                doc.add_paragraph(f"({si}) {s}")
            for n in notes:
                doc.add_paragraph(n)
    cp = doc.core_properties
    cp.created = cp.modified = STAMP
    cp.last_modified_by = "make_fixture"
    cp.author = "社会福祉法人 ○○会"
    cp.revision = 1
    doc.save(path)
    pin_zip(path)


# ---------- pdf ----------

def make_pdf(path):
    import fitz
    W, H = 595, 842
    doc = fitz.open()
    FS = 10.5

    def body(page, lines, y0=90):
        y = y0
        for kind, text in lines:
            size = {"h": 15, "lead": FS, "step": FS, "dec": FS, "meta": 9}[kind]
            x0 = 50 if kind in ("h", "lead") else 60
            rect = fitz.Rect(x0, y, W - 50, H - 60)
            used = page.insert_textbox(rect, text, fontname="japan", fontsize=size, lineheight=1.45)
            if used < 0:
                raise RuntimeError(f"ページに収まらない: {text[:20]}")
            y = H - 60 - used + (6 if kind != "h" else 10)
        return y

    cover = doc.new_page(width=W, height=H)
    cover.insert_textbox(fitz.Rect(50, 300, W - 50, 400), "本部拠点　経理事務手順書", fontname="japan", fontsize=24, align=1)
    cover.insert_textbox(fitz.Rect(50, 420, W - 50, 500), "社会福祉法人 ○○会　本部拠点\n2026 年 4 月改訂", fontname="japan", fontsize=12, align=1)
    toc = doc.new_page(width=W, height=H)
    lines = [("h", "目次")]
    for i, (name, _lead, _rows) in enumerate(HQ, 1):
        pad = "…" * max(1, 24 - len(name))
        lines.append(("lead", f"{i}. {name} {pad} {i + 1}"))
    lines.append(("meta", "ページ番号は、紙面の下に刷られた番号。法人の経理規程・経理事務マニュアルの決まりに従い、本部拠点で行う手順を書いたもの。"))
    body(toc, lines)
    toc.insert_textbox(fitz.Rect(0, H - 40, W, H - 20), "- 1 -", fontname="helv", fontsize=9, align=1)
    for i, (name, lead, rows) in enumerate(HQ, 1):
        page = doc.new_page(width=W, height=H)
        lines = [("h", f"{i}. {name}"), ("lead", lead)]
        n = 0
        for row in rows:
            if row[0] == "step":
                n += 1
                _k, actor, text, system, files = row
                extra = []                                           # 内蔵の日本語の字体は半角の空白が広く出るので、空白を入れない
                if system:
                    extra.append(f"システム:{system}")
                if files:
                    extra.append(f"帳票:{files}")
                lines.append(("step", f"({n}){actor}は、{text}" + (f"［{'／'.join(extra)}］" if extra else "")))
            elif row[0] == "decision":
                _k, q, yes, no = row
                lines.append(("dec", f"【判断】{q}→はい:{yes}／いいえ:{no}"))
            else:
                lines.append(("meta", row[1]))
        body(page, lines)
        page.insert_textbox(fitz.Rect(0, H - 40, W, H - 20), f"- {i + 1} -", fontname="helv", fontsize=9, align=1)
    doc.set_metadata({"title": "本部拠点 経理事務手順書", "author": "社会福祉法人 ○○会", "producer": "make_fixture", "creator": "make_fixture",
                      "creationDate": "D:20260401000000", "modDate": "D:20260401000000"})
    doc.save(path, no_new_id=True)
    doc.close()


# ---------- xlsx ----------

def make_xlsx(path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = ["No", "手順", "担当", "システム", "帳票", "備考"]
    widths = [5, 46, 14, 16, 18, 26]
    for cat, rows in TY:
        ws = wb.create_sheet(cat)
        ws["A1"] = f"特別養護老人ホーム拠点　経理事務手順表　{cat}"
        ws["A1"].font = Font(name=JP_FONT, size=12, bold=True)
        ws["A2"] = "法人の経理規程・経理事務マニュアルに従う。この表は、特養拠点で行う手順をまとめたもの（詳しい決まりはマニュアルを見る）。"
        ws["A2"].font = Font(name=JP_FONT, size=9)
        for c, (h, w) in enumerate(zip(head, widths), 1):
            cell = ws.cell(4, c, h)
            cell.font = Font(name=JP_FONT, size=10, bold=True)
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.border = border
            cell.alignment = Alignment(vertical="center", horizontal="center")
            ws.column_dimensions[chr(64 + c)].width = w
        for r, row in enumerate(rows, 5):
            for c, v in enumerate(row, 1):
                cell = ws.cell(r, c, v)
                cell.font = Font(name=JP_FONT, size=10)
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 30
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
    wb.properties.created = wb.properties.modified = STAMP
    wb.properties.creator = "make_fixture"
    wb.properties.lastModifiedBy = "make_fixture"
    wb.save(path)
    pin_zip(path)


# ---------- pptx ----------

def make_pptx(path):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    blank = prs.slide_layouts[6]

    def jp(run, size=14, bold=False, color=None):
        run.font.name = JP_FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = RGBColor(*color)
        rpr = run._r.get_or_add_rPr()
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            ea = rpr.makeelement(qn("a:ea"), {})
            rpr.append(ea)
        ea.set("typeface", JP_FONT)
        return run

    def textbox(slide, text, x, y, w, h, size=14, bold=False, color=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, ln in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            jp(p.add_run(), size, bold, color).text = ln
        return tb

    def title(slide, text):
        textbox(slide, text, 0.5, 0.3, 12, 0.8, 24, True)

    s = prs.slides.add_slide(blank)
    textbox(s, "保育所拠点　経理の流れ", 1, 2.6, 11, 1.2, 36, True)
    textbox(s, "社会福祉法人 ○○会　保育所拠点\n2026 年 4 月", 1, 4.2, 11, 1.2, 18)
    s = prs.slides.add_slide(blank)
    title(s, "目次")
    textbox(s, "\n".join(f"{i}. {name}" for i, name in enumerate(["はじめに"] + [f[0] for f in HO_FLOWS], 1)), 1, 1.4, 11, 5.5, 16)
    for name, paras in HO_TEXT:
        s = prs.slides.add_slide(blank)
        title(s, name)
        textbox(s, "\n\n".join(paras), 1, 1.5, 11, 5, 18)

    def box(slide, kind, text, actor, x, y, w=2.6, h=1.0):
        shape = {"start": MSO_SHAPE.FLOWCHART_TERMINATOR, "end": MSO_SHAPE.FLOWCHART_TERMINATOR, "proc": MSO_SHAPE.ROUNDED_RECTANGLE,
                 "dec": MSO_SHAPE.DIAMOND}[kind]
        if kind == "dec":
            w, h = 3.0, h * 1.4
        sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor(0xDD, 0xEB, 0xF7) if kind == "proc" else RGBColor(0xFF, 0xF2, 0xCC) if kind == "dec" else RGBColor(0xE2, 0xF0, 0xD9)
        sh.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
        tf = sh.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        jp(p.add_run(), 12, False, (0, 0, 0)).text = text
        if actor:
            p2 = tf.add_paragraph()
            jp(p2.add_run(), 10, False, (0x1F, 0x4E, 0x79)).text = f"［{actor}］"
        return sh

    def link(slide, p, q, label, i=2, j=0):
        c = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, 0, 0, 0, 0)
        c.begin_connect(p, i)
        c.end_connect(q, j)
        c.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
        c.line.width = Pt(1.5)
        ln = c.line._get_or_add_ln()
        tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle"})
        ln.append(tail)
        if label:
            x = (p.left + q.left) / 2 / 914400 + 0.2
            y = (p.top + p.height + q.top) / 2 / 914400 - 0.2
            textbox(slide, label, x, y, 1.0, 0.4, 11, False, (0xC0, 0x00, 0x00))

    for name, nodes, edges in HO_FLOWS:
        s = prs.slides.add_slide(blank)
        title(s, name)
        # 本流を縦に、判断の「いいえ」側を右に置く
        main_ids = [i for i, n in enumerate(nodes) if not any(e[0] == i and e[2] == "いいえ" for e in edges) or True]
        side = {}
        for a, b, lab in edges:
            if nodes[a][0] == "dec" and nodes[b][0] == "proc" and lab:
                side[b] = True
        shapes = {}
        # 本流が長いと終端がスライドの下にはみ出すので、縦の間隔と箱の高さを縮めて収める
        mains = [kind for i, (kind, _, _) in enumerate(nodes) if not side.get(i)]
        span = sum(1.25 if k != "dec" else 1.65 for k in mains[:-1]) + (1.4 if mains[-1] == "dec" else 1.0)
        f = min(1.0, (7.5 - 0.3 - 1.3) / span)
        y = 1.3
        col_main, col_side = 1.2, 5.4
        for i, (kind, text, actor) in enumerate(nodes):
            if side.get(i):
                shapes[i] = box(s, kind, text, actor, col_side, y - 1.25 * f, h=f)
                continue
            shapes[i] = box(s, kind, text, actor, col_main, y, h=f)
            y += (1.25 if kind != "dec" else 1.65) * f
        for a, b, lab in edges:
            pa, pb = shapes[a], shapes[b]
            if side.get(b):
                link(s, pa, pb, lab, 3, 1)
            elif side.get(a):
                link(s, pa, pb, lab, 2, 3)
            else:
                link(s, pa, pb, lab)
        textbox(s, "［ ］は担当。線は作業の順。", 8.8, 6.7, 4, 0.5, 10, False, (0x59, 0x59, 0x59))
    cp = prs.core_properties
    cp.created = cp.modified = STAMP
    cp.last_modified_by = "make_fixture"
    cp.author = "社会福祉法人 ○○会"
    cp.revision = 1
    prs.save(path)
    pin_zip(path)


# ---------- 入口 ----------

def build(out):
    out = pathlib.Path(out)
    paths = {}
    for key, folder in SITES.items():
        (out / folder).mkdir(parents=True, exist_ok=True)
    paths["shared"] = out / SITES["shared"] / "法人経理規程・経理事務マニュアル.docx"
    paths["HQ"] = out / SITES["HQ"] / "本部_経理事務手順書.pdf"
    paths["TY"] = out / SITES["TY"] / "特養_経理事務手順表.xlsx"
    paths["HO"] = out / SITES["HO"] / "保育所_経理の流れ.pptx"
    make_docx(paths["shared"])
    make_pdf(paths["HQ"])
    make_xlsx(paths["TY"])
    make_pptx(paths["HO"])
    return paths


def sha1(p):
    return hashlib.sha1(pathlib.Path(p).read_bytes()).hexdigest()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    paths = build(a.out)
    for k, p in paths.items():
        print(f"{k}: {p}（{p.stat().st_size:,} バイト・sha1 {sha1(p)[:10]}）")
    if a.check:
        with tempfile.TemporaryDirectory() as td:
            again = build(td)
            bad = [k for k in paths if sha1(paths[k]) != sha1(again[k])]
        print("決定性: " + ("全ファイルの sha1 が同じ" if not bad else f"違う: {', '.join(bad)}"))
        sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
