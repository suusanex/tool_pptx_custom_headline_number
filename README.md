# PowerPoint カスタム見出し連番ツール

`.pptx` の見出しプレフィックスを JSON ルールに従って付与・更新する .NET コンソールアプリです。

## コマンド

```bash
dotnet run --project src\PptxHeadlineNumbering -- inspect <input.pptx>
dotnet run --project src\PptxHeadlineNumbering -- apply <input.pptx> <output.pptx> --rule <rule.json>
```

## inspect / apply / rule.json の関係

- `inspect` は読み取り専用です。`.pptx` の段落構造を一覧化し、どの段落を番号付け対象にするかを確認します。
- `apply` は書き込み用です。`--rule` で指定した `rule.json` に従って、対象段落の先頭を番号文字列に差し替えて `output.pptx` を作成します。
- `rule.json` は `apply` の動作を決めます。`levels` で「どの段落が H1/H2/H3 か」、`format` で「番号の見た目」、`prefixRegex` で「既存番号の消し方」、`separator` で「番号と本文の区切り方」を指定します。
- さらに `excludedSlideRanges` を使うと、番号付与対象から外したいページ範囲を複数指定できます。表紙〜目次、巻末資料などの除外を想定しています。
- `inspect` で段落構造を確認したあとに `rule.json` を作り、最後に `apply` を実行する流れが基本です。
- ルールを変更した場合は、その都度 `rule.json` を差し替えて `apply` を再実行できます。必要なら事前に `inspect` で対象段落を見直してください。
- `PlaceholderType` が空の段落を対象にしたい場合は、`match.shapeNames` または `matches[].shapeNames` に `inspect` の `ShapeName` を指定します。
- レイアウト違いで同じ見出しレベルを複数条件の OR にしたい場合は、同じ `name` を重複させず、1つの level の中で `matches` 配列を使って複数条件を並べます。
- `excludedSlideRanges` の番号は **1-based** です。`inspect` の `SlideIndex` は 0-based なので、そのままではなく **`SlideIndex + 1`** で考えてください。

## ルールファイル

`sample-rule.json` をコピーして利用してください。主要プロパティ:

- `prefixRegex`: 既存プレフィックス検出正規表現
- `separator`: 新プレフィックスと本文の区切り文字
- `insertWhenPrefixMissing`: プレフィックス未検出時に先頭挿入するか
- `excludedSlideRanges`: apply 対象外にするスライド範囲（`startSlideNumber` / `endSlideNumber`、1-based、両端含む）
- `levels`: 見出しレベル定義（`match` または `matches`, `format`, `resetsOnNewLevel`）
- `match.placeholderTypes`: `title` / `ctrTitle` / `body` / `obj` で対象を絞る
- `match.shapeNames`: `inspect` の `ShapeName` を使って対象を絞る
- `matches`: 複数の `match` 条件を OR で並べる

## テスト

```bash
dotnet test ruby-larch.sln
```
