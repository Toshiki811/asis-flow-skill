# 比べる行（業務種別）の進め方（工程 4）

比べる行（`axis.json`）は、表の縦軸（業務分類 › 業務種別）であり、タスクの対応を 1 体のサブエージェントに渡す単位でもある。
決まりは `references/compare-rules.md` の 3 章。

## 親（進行役）のやること

1. `python3 scripts/sc_axis.py draft -w <CMP>` — 基準の拠点（無ければ 1 つ目の拠点）の分類から、行の下書き。ほかの拠点は「none・理由なし」で置かれる。
2. `python3 scripts/sc_axis.py material -w <CMP>` の大きさを見る。4 万字までなら、下の指示文で 1 体のサブエージェント（または親）が全部を決める。
   超えるときは、業務分類ごとに 1 体（指示文の `{SCOPE}` に、その業務分類の行と、当てる候補の拠点の分類を書く）。
3. そろったら `python3 scripts/sc_axis.py show -w <CMP>`（まだどの行にも当てていないフローが出る）→ `python3 scripts/sc_check.py -w <CMP> --stage axis` を FAIL 0 に。
4. `python3 scripts/sc_axis.py plan -w <CMP>` — 行ごとの材料の大きさ。4 万字を超える行は、子の行に分ける（3.6）。

## サブエージェントへの指示文（ここからコピー。`{…}` を置き換える）

あなたは、拠点ごとの業務の一覧を突き合わせて、「どの拠点のどの業務フローが、同じ業務か」を決める担当です。
**名前が似ているだけで同じにしないでください。** 開始のきっかけ・終わりの成果・担当・作業の並びが同じ業務かで決めます。違う業務を同じ行にすると、比較の表で誰も気づけない誤りになります。

### 担当する範囲
- 比較の作業フォルダ: `{CMP}`
- 範囲: {SCOPE}（全部なら「全部」）
- 書く先: `{CMP}/axis.json`（下書きの rows を直し、足す）

### 最初に読むもの
- `{SKILL}/references/compare-rules.md` の §0・§3
- `{SKILL}/references/schemas.md` の「比べる行」（書く形。表に無い項目は使えない）
- 材料: `python3 {SKILL}/scripts/sc_axis.py material -w {CMP}`（拠点ごとの分類とフローの要約）。中身を詳しく見るときは `python3 {SKILL}/scripts/sc_match.py material -w {CMP} --row <行>`

### 決め方
1. 下書きの行ごとに、ほかの拠点のフローを当てる。当て方（relation）: `same`／`part`（この行の業務の一部）／`split`（1 本のフローが 2 つ以上の行にまたがる。`parts[].nodes` に、この行に入れる箱の鍵を書く）／`none`（記載なし）／`out_of_scope`（その拠点で対象外）。same 以外は `reason` を書く（none は、何を探して無かったか）。
2. 下位フローのある分類は、ふつう `expand: true`（開いて 1 本で比べる）。拠点どうしで下位フローの単位がそろっていて、子ごとに比べるほうが分かりやすいときだけ、子の行を足す（親は `expand: false`、子の行の `parent` に親の行の ID、ID は `B03-1`）。
3. どの行にも当たらない業務は、行を足す: 基準ありは `origin: not_in_baseline`（基準の拠点は none）、基準なしは `origin: merged`。行の名前は、その拠点の分類の名前（日本語でなければ訳し、`name_from` にもとを書く）。足す行は、その拠点で前にある業務の行のすぐあとに置く（rows の並びが、表と図の並び）。
4. 日本語でない拠点のフローには、日本語の名前を `parts[].title` に書く。
5. 業務でないフロー（総則・用語の説明だけ）・比べる意味の無い箱は `not_compared` に理由つきで入れる。**どの拠点の作業の箱も、ちょうど 1 つの行か not_compared に入るように**する。
6. 決められない当て方は、いちばん確からしいほうに仮置きし、`questions` に種別「対応」・`effect`「対応」で書く。

### 終わったら
1. `python3 {SKILL}/scripts/sc_validate.py -w {CMP} --kind axis` と `python3 {SKILL}/scripts/sc_check.py -w {CMP} --stage axis --fast` を FAIL 0 に。
2. 返事は 1 行だけ: 行の数・足した行の数・確認事項の数。
