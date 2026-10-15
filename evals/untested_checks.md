# 発火を試験しない点検

点検（`check_flow.py`・`check_layout.py`・`check_extract.py`）の FAIL のうち、自己点検で、誤りを仕込んで出すことをしないもの。

- LLM・人が書くファイル（カード・業務分類・フロー・台帳・直しの一覧・木・用語の対応・採否表・照合の報告）を見る点検は、ここに載せない。どれも、自己点検で発火を確かめる。
- ここに載るのは、次の 3 種類だけ。
  - 機械が書くもの（図・報告書・業務マトリクス一覧）を、作り直したものと比べる点検の、細かい内訳。手で直した・書き出し直していない図や表は、作り直しとの一致の点検が先に止める（その点検は、試験している）。
  - 守り。スキルの部品が無い・外の道具が失敗した、など、試験の環境では作れないとき。
  - 同じ種類の FAIL が多いときの、残りの数の行（「ほかに N 件」）。1 つ目の行は、試験している。
- `evals/coverage.py` が、自己点検で一度も出ない FAIL のうち、この表に無いものを数える。終了コード 1 になるのは、次のどれか（試験を足すか、理由を書いてこの表に足す・表を直す）:
  - 一覧に無いのに、一度も出ない点検がある
  - 表の行が、いまの点検のどれにも当たらない（点検の文を変えた・消した）
  - 表に載っているのに、自己点検で出た（試験できたので、表から外す）
  - 自己点検に NG がある
- 数える所: `Report.fail`・`fail_if`・配置の点検の `yield "fail"`、納品の門が FAIL に上げる WARN（`deliver.PROMOTE` の語を含む `Report.warn`）。ほかの WARN は数えない。
  点検の共通の部品（`common.FAIL_HELPERS`）は、部品の中でなく、呼ぶ所ごとに数える（どの点検から出たかを、呼んだ所の行で控える）。
- 表の行は、ファイル・関数・文の一部で点検を指す（行の番号は書かない。点検の文を変えたら、ここも直す）。

| ファイル | 関数 | 文の一部 | 試験しない理由 |
| --- | --- | --- | --- |
| check_flow.py | `read_cells` | 中身の無い object | 図（drawio）は機械が書く。手で直した図は、作り直しとの一致（`check_regenerated`）が先に止める |
| check_flow.py | `check_cells` | 根のセル（id 0 と 1）が無い | 同上 |
| check_flow.py | `check_cells` | に mxGeometry が無い | 同上（線・図形の 2 か所） |
| check_flow.py | `check_page_matches_flow` | flow.json のノード | 同上 |
| check_flow.py | `check_page_matches_flow` | が flow.json に無い（図を手で直した疑い | 同上 |
| check_flow.py | `check_page_matches_flow` | に出典の属性（src）が無い | 同上 |
| check_flow.py | `check_page_matches_flow` | 線が flow.json と違う | 同上 |
| check_flow.py | `check_page_matches_flow` | レーンが flow.json と違う | 同上 |
| check_flow.py | `check_drawio` | システムと起動点の一覧のページが無い | 同上 |
| check_flow.py | `check_drawio` | mxGraphModel/root が無い | 同上 |
| check_flow.py | `check_regenerated` | が無い（書き出していない） | 守り。書き出していない図は、図のページの点検（`check_drawio`。試験している）が先に止める |
| check_flow.py | `check_drawio` | assets/activity_styles.json が無い・古い | 守り。スキルの部品（assets）を消さない限り起きない |
| check_flow.py | `check_code` | 取り出した構造の誤り | 守り。機械の取り出し（`code_struct.py`）の誤り。LLM の木の誤りは、木を構造にする前に止める（「木の点検」の試験） |
| check_flow.py | `matrix_results` | 保存した計算の結果が違う式のセルが、ほかに | 同じ種類の FAIL が多いときの、残りの数の行。1 つ目の行は試験している |
| check_flow.py | `check_matrix` | のシートが無い（テンプレートに要る） | 守り。シートを消したブックは、読み戻し（`read_back`）が先に止める |
| check_flow.py | `check_matrix` | の数が、いまのデータと違う: ブック | 業務マトリクス一覧は機械が書く。全セルの比べ（試験している）の内訳 |
| check_flow.py | `check_matrix` | いまのデータと違うセルが、ほかに | 同じ種類の FAIL が多いときの、残りの数の行。1 つ目の行は試験している |
| check_flow.py | `check_matrix` | のフローの並びが、業務分類の順でない | 業務マトリクス一覧は機械が書く。行を並べ替えたブックは、全セルの比べが先に止める |
| check_layout.py | `main` | PNG に書き出せない（ | 守り。draw.io の書き出しが失敗したとき（アプリが無い・壊れた）。試験の環境では作れない |
