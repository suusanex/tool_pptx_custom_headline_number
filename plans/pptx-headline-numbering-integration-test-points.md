# テスト観点: PowerPoint カスタム見出し連番ツール

> 元Plan: `plans/pptx-headline-numbering.md`
> ランタイムエビデンス: `plans/pptx-headline-numbering-runtime-evidence.md`

---

## 観点一覧

| ID | 分類 | 観点名 | Plan トレーサビリティ |
|---|---|---|---|
| TP-010 | 機能（正常系） | inspect モード — TSV 出力の正確性 | S-010, FR: SlideWalker, InspectCommand |
| TP-020 | 機能（正常系） | apply モード — 番号付与の正確性 | S-020, FR: ApplyCommand, NumberingRule |
| TP-030 | 機能（正常系） | apply モード — カウンタリセットの正確性 | S-020, FR: HeadingCounter |
| TP-040 | 機能（正常系） | apply モード — 先頭トークン差し替え（Run分割対応） | S-020, FR: PrefixReplacer |
| TP-050 | 機能（正常系） | apply モード — 冪等性 | S-030, FR: PrefixReplacer + HeadingCounter |
| TP-060 | 異常 | ファイル I/O エラー | S-040, FR: エラーハンドリング方針 |
| TP-070 | 異常 | CLI 引数・パス検証エラー | S-050, FR: ApplyCommand パス検証 |
| TP-080 | 異常 | JSON ルール不正 | S-060, FR: NumberingRule デシリアライズ |
| TP-090 | 負荷 | 大量データの処理 | FR: SlideWalker, ApplyCommand |
| TP-100 | 連続 | 連続実行・再実行 | S-030, FR: HeadingCounter, PrefixReplacer |

---

## 観点詳細

### TP-010: inspect モード — TSV 出力の正確性

**分類:** 機能（正常系）

**条件:**

- 入力パラメータパターン
  - 1枚のスライドに title プレースホルダーのみ
  - 1枚のスライドに title + body プレースホルダー（段落レベル 0, 1 混在）
  - ctrTitle プレースホルダーを持つスライド（表紙レイアウト等）
  - obj プレースホルダーを持つスライド
  - プレースホルダー無しの自由配置図形を含むスライド
  - 複数スライド（3枚以上）
  - 空テキストの段落を含むスライド
  - 段落レベルが 2 以上のケース（定義レベル超過）
- 外部要素の応答パターン
  - 正常な .pptx ファイル（Open XML SDK でオープン成功）

**期待:**

- TSV ヘッダー行（`SlideIndex\tShapeName\tPlaceholderType\tLevel\tText`）が先頭に出力される
- 各段落が1行1レコードで出力される
- SlideIndex は 0-based でスライド順に出力される
- PlaceholderType が `title` / `ctrTitle` / `body` / `obj` / 空（自由配置図形）で正しく表示される
- Level が `pPr/@lvl` の値（未指定時は 0）で正しく表示される
- Text に段落全体のプレーンテキストが出力される
- 終了コード 0 で終了する
- 自由配置図形もinspectでは出力される（applyではスキップ対象だが、inspectは全図形対象）

---

### TP-020: apply モード — 番号付与の正確性

**分類:** 機能（正常系）

**条件:**

- 入力パラメータパターン
  - H1 対象: title プレースホルダーの段落
  - H1 対象: ctrTitle プレースホルダーの段落
  - H2 対象: body プレースホルダー、paragraphLevel=0
  - H2 対象: obj プレースホルダー、paragraphLevel=0
  - H2 対象: PlaceholderType が空でも `shapeNames` で指定した ShapeName の段落
  - H3 対象: body プレースホルダー、paragraphLevel=1
  - H3 対象: obj プレースホルダー、paragraphLevel=1
  - マッチしない段落: プレースホルダー無し自由配置図形、paragraphLevel=2以上
  - 複数スライドにまたがる連番（例: スライド1=H1:1, スライド2=H1:2, スライド3=H1:3）
- ルールJSON パターン
  - 標準3レベル（H1=`{H1}.`, H2=`{H1}.{H2}`, H3=`{H3})`）
  - `match.shapeNames` を使うケース
  - `matches` 配列で同一 level の OR 条件を持つケース
  - `excludedSlideRanges` で複数のページ範囲を除外するケース
  - separator が半角スペースのケース
  - separator が全角スペースのケース
- 外部要素の応答パターン
  - 正常な .pptx + 正常な rule.json

**期待:**

- マッチした段落の先頭トークンが format テンプレートに基づく番号文字列に差し替えられる
- separator が番号と本文テキストの間に挿入される
- PlaceholderType が空でも ShapeName 条件に一致すれば番号付与される
- `matches` に複数条件を定義した場合、それらが OR として評価され同一 level のカウンタを共有する
- `excludedSlideRanges` に含まれるスライドでは番号付与されず、除外区間ではカウンタも進まない
- マッチしない段落（自由配置図形、対象外レベル）は一切変更されない
- levels 配列を順にチェックし、最初にマッチしたレベルが採用される
- output.pptx が正常に保存される
- 終了コード 0 で終了する

---

### TP-030: apply モード — カウンタリセットの正確性

**分類:** 機能（正常系）

**条件:**

- 入力パラメータパターン
  - H1 → H2 → H2 → H1 → H2 の順に段落が出現（H1増加でH2リセット確認）
  - H1 → H2 → H3 → H3 → H2 → H3 の順に段落が出現（H2増加でH3リセット確認）
  - H1 → H2 → H3 → H1 → H2 → H3 の順（H1増加でH2・H3両方リセット確認）
  - H1 のみ連続（リセット対象なし、単調増加）
  - H3 のみ連続（リセットなし、単調増加）
- ルールJSON パターン
  - `resetsOnNewLevel: []` (リセットなし — H1)
  - `resetsOnNewLevel: ["H1"]` (H1変化でリセット — H2)
  - `resetsOnNewLevel: ["H1", "H2"]` (H1またはH2変化でリセット — H3)

**期待:**

- H1 がインクリメントされたとき、H2 / H3 カウンタが 0 にリセットされ次回インクリメントで 1 から始まる
- H2 がインクリメントされたとき、H3 カウンタが 0 にリセットされる
- H3 のリセット対象に H1 が含まれる場合、H1 変化時にも H3 がリセットされる
- リセット後の番号文字列が正しくフォーマットされる（例: H1=2 → H2=1.1 = `2.1`）

---

### TP-040: apply モード — 先頭トークン差し替え（Run分割対応）

**分類:** 機能（正常系）

**条件:**

- 入力パラメータパターン（Run構造）
  - 先頭トークンが単一 Run に収まっているケース（例: `"1. はじめに"` が1Run）
  - 先頭トークンが複数 Run にまたがるケース（例: `"1"` と `". "` と `"はじめに"` が別Run）
  - 先頭 Run が空テキストのケース
  - 先頭トークンに全角スペース（`\u3000`）がセパレータとして含まれるケース
- prefixRegex パターン
  - デフォルト `^[^\s\u3000]+[\s\u3000]+` で半角/全角スペース前までマッチ
- insertWhenPrefixMissing パターン
  - `true`: プレフィックスが見つからない段落にも番号 + separator を先頭挿入
  - `false`: プレフィックスが見つからない段落はスキップ（変更なし）
- 書式継承
  - 最初の Run の RunProperties（フォント、サイズ、色等）が挿入テキストに継承されるか

**期待:**

- 単一 Run の場合: prefixRegex マッチ部分が新番号に置換される
- 複数 Run の場合: Run チェーンを連結してマッチし、マッチ文字数分の Run テキストを削除、最初の Run 先頭に新番号を挿入
- 全角スペースがセパレータの場合も正しくマッチ・差し替えされる
- `insertWhenPrefixMissing=true` で、既存プレフィックスなしの段落に番号が挿入される
- `insertWhenPrefixMissing=false` で、既存プレフィックスなしの段落が変更されない
- 挿入されたテキストは最初の Run の RunProperties を継承する

---

### TP-050: apply モード — 冪等性

**分類:** 機能（正常系）

**条件:**

- 手順
  1. 初回: 元 .pptx に apply → output1.pptx を生成
  2. 2回目: output1.pptx を入力として同一ルールで apply → output2.pptx を生成
- 入力パラメータパターン
  - 全レベル（H1, H2, H3）を含む複数スライドの .pptx
  - 同一の rule.json を使用
- prefixRegex の適合性
  - 1回目 apply で付与された番号文字列が、prefixRegex で正しくマッチすること

**期待:**

- output1.pptx と output2.pptx の全段落テキストが完全一致する（冪等）
- 番号が二重付与されない
- Run 構造が不必要に変化しない（番号テキスト以外の Run は保持）
- 3回以上繰り返しても結果が同一

---

### TP-060: ファイル I/O エラー

**分類:** 異常

**条件:**

- 入力パラメータパターン
  - inspect: 存在しないファイルパスを指定
  - apply: 存在しない入力 .pptx パスを指定
  - apply: 存在しないルール JSON パスを指定
- 外部要素の応答パターン
  - ファイルシステムが FileNotFoundException を返す
  - 入力 .pptx が壊れた ZIP（Open XML SDK が InvalidDataException / OpenXmlPackageException 等を返す）

**期待:**

- FileNotFoundException（またはそれに相当する例外）が発生する
- 例外の `Exception.ToString()` がトレースログ（`System.Diagnostics.Trace`）に出力される
- stderr にエラーメッセージが出力される
- 終了コード 1 で終了する
- フォールバック処理は行われない（処理失敗として終了）
- 壊れた .pptx の場合も例外が伝播しエラー終了する

---

### TP-070: CLI 引数・パス検証エラー

**分類:** 異常

**条件:**

- 入力パラメータパターン
  - apply: 入力パスと出力パスが同一（`apply same.pptx same.pptx --rule r.json`）
  - apply: 入力パスと出力パスが正規化後に同一（相対パス/絶対パス混在、`./input.pptx` vs `input.pptx` 等）
  - サブコマンド未指定（引数なし実行）
  - 不明なサブコマンド指定
  - apply で `--rule` オプション未指定

**期待:**

- 入出力パス同一: ArgumentException が発生し、ファイル操作前にエラー終了する（データ破壊防止）
- パス正規化後の同一検出: 正規化パスで比較し同一と判定されること
- 例外の `Exception.ToString()` がトレースログに出力される
- stderr にエラーメッセージが出力される
- 終了コード 1 で終了する（ExitCode=0 にならない）
- `System.CommandLine` が引数不足/不明コマンドを検出した場合はヘルプまたはエラーメッセージが表示される

---

### TP-080: JSON ルール不正

**分類:** 異常

**条件:**

- 入力パラメータパターン
  - JSON 構文エラー（波括弧不整合、カンマ欠落等）
  - JSON は正しいがスキーマ不適合（`levels` 配列が無い、`format` フィールドが欠損等）
  - 空ファイル（0バイト）
  - `levels` 配列が空（`[]`）
  - `prefixRegex` に不正な正規表現が指定されている
- 外部要素の応答パターン
  - ファイル読み込み自体は成功し、内容がデシリアライズ不可

**期待:**

- JSON 構文エラー: JsonException が発生する
- スキーマ不適合: デシリアライズ時または使用時に例外が発生する
- 不正な正規表現: 正規表現コンパイル時に例外が発生する
- 例外の `Exception.ToString()` がトレースログに出力される
- stderr にエラーメッセージが出力される
- 終了コード 1 で終了する
- フォールバック処理は行われない

---

### TP-090: 大量データの処理

**分類:** 負荷

**条件:**

- 入力パラメータパターン
  - 多数のスライド（50枚以上）を持つ .pptx
  - 1スライド内に多数の段落（20段落以上）を持つスライド
  - 空スライド（段落なし）が複数混在する .pptx
  - 番号付与対象外の自由配置図形が大量にあるスライド
- 手順
  - inspect: 大量データの TSV 出力が完走すること
  - apply: 大量データの番号付与が完走すること

**期待:**

- 処理が正常に完走する（OutOfMemoryException 等が発生しない）
- inspect の TSV 出力が全段落分出力される
- apply の番号付与が全対象段落に正しく適用される
- 空スライドが含まれていてもエラーにならずスキップされる
- 連番がスライドをまたいで正しく継続される（中間の空スライドで途切れない）

---

### TP-100: 連続実行・再実行

**分類:** 連続

**条件:**

- 手順パターン
  1. ルールA で apply → output-A.pptx 生成 → ルールB で apply（output-A.pptx → output-B.pptx）→ ルールB の番号体系で付番されること
  2. apply 実行後、同じ output.pptx に対して inspect 実行 → apply 結果が inspect の TSV で確認できること
  3. 同一入力に対して同一ルールで apply を3回連続実行 → 3回とも出力が同一（TP-050 の拡張）
- ルールJSON パターン
  - ルールA: `{H1}.` / `{H1}.{H2}` / `{H3})`
  - ルールB: `第{H1}章` / `{H1}-{H2}` / `({H3})`

**期待:**

- ルール切替時: 旧ルールの番号が prefixRegex で除去され、新ルールの番号に差し替えられる
- apply 後の inspect: TSV の Text 列が番号付与後のテキストを反映している
- 3回連続実行: 毎回同一の出力（前回の番号が正しく上書きされる）
- 前回の apply で使用した RunProperties が残存しても、新ルール適用に支障がない

---

## ブラックボックス入力パターン一覧

### CLI コマンドパターン

| # | パターン | 代表的カバー観点 |
|---|---|---|
| 1 | `inspect <valid.pptx>` | TP-010 |
| 2 | `apply <valid.pptx> <output.pptx> --rule <valid.json>` | TP-020, TP-030, TP-040 |
| 3 | `inspect <nonexistent.pptx>` | TP-060 |
| 4 | `apply <nonexistent.pptx> <out.pptx> --rule <r.json>` | TP-060 |
| 5 | `apply <in.pptx> <out.pptx> --rule <nonexistent.json>` | TP-060 |
| 6 | `apply <same.pptx> <same.pptx> --rule <r.json>` | TP-070 |
| 7 | `apply <in.pptx> <out.pptx> --rule <broken.json>` | TP-080 |
| 8 | （引数なし実行） | TP-070 |
| 9 | `apply <in.pptx> <out.pptx>`（--rule 未指定） | TP-070 |

### .pptx 内部構造パターン

| # | パターン | 代表的カバー観点 |
|---|---|---|
| A | title プレースホルダーのみ | TP-010, TP-020 |
| B | ctrTitle プレースホルダー（表紙） | TP-010, TP-020 |
| C | body プレースホルダー（level=0, 1 混在） | TP-010, TP-020, TP-030 |
| D | obj プレースホルダー | TP-010, TP-020 |
| E | 自由配置図形（プレースホルダーなし） | TP-010, TP-020 |
| F | 空テキスト段落 | TP-010, TP-040 |
| G | 段落レベル 2 以上 | TP-010, TP-020 |
| H | 先頭トークンが単一 Run | TP-040 |
| I | 先頭トークンが複数 Run にまたがる | TP-040 |
| J | 全角スペースセパレータ | TP-040 |
| K | 既存番号プレフィックスあり | TP-050 |
| L | 50枚以上のスライド | TP-090 |
| M | 空スライド混在 | TP-090 |
| N | 壊れた ZIP（非 .pptx） | TP-060 |

### JSON ルールパターン

| # | パターン | 代表的カバー観点 |
|---|---|---|
| i | 標準3レベル（H1/H2/H3） | TP-020, TP-030 |
| ii | insertWhenPrefixMissing=true | TP-040 |
| iii | insertWhenPrefixMissing=false | TP-040 |
| iv | JSON 構文エラー | TP-080 |
| v | 必須フィールド欠損 | TP-080 |
| vi | levels 配列が空 | TP-080 |
| vii | 不正な prefixRegex | TP-080 |
| viii | 空ファイル（0バイト） | TP-080 |
| ix | 異なる番号書式（`第{H1}章` 形式） | TP-100 |

---

## Plan トレーサビリティマトリクス

| Plan シナリオ / 要件 | カバーする観点 |
|---|---|
| S-010 inspect 正常系 | TP-010 |
| S-020 apply 正常系 | TP-020, TP-030, TP-040 |
| S-030 apply 冪等性 | TP-050 |
| S-040 inspect ファイル不存在 | TP-060 |
| S-050 apply 入出力パス同一 | TP-070 |
| S-060 apply JSON 不正 | TP-080 |
| FR: SlideWalker プレースホルダー種別判定 | TP-010, TP-020 |
| FR: HeadingCounter リセットロジック | TP-030 |
| FR: PrefixReplacer Run分割対応 | TP-040 |
| FR: 番号書式のJSON差し替え | TP-020, TP-100 |
| FR: エラーハンドリング方針（例外トレースログ） | TP-060, TP-070, TP-080 |
| Non-goal: 対象外図形の非変更 | TP-020（マッチしない段落のスキップ確認） |
