# PowerPoint カスタム見出し連番ツール

## Background / Goal

PowerPoint プレゼン資料の各スライドに、資料全体をまたいだ特殊な連番見出しを付与する必要がある。
番号の振り方は「タイトル = `1.` `2.` `3.`」「箇条書き1段目 = `1.1` `1.2`」「箇条書き2段目 = `1)` `2)`」のように統一性がなく、かつ毎回変わる。

このルールを **JSON ルールファイル** で外部化し、何度でも差し替えできる C# コンソールアプリケーションを Open XML SDK ベースで作成する。

主要モード:

| モード | 目的 |
| --- | --- |
| **inspect** | 全スライドの段落構造（Slide番号 / Shape名 / Placeholder種別 / 段落level / 現在テキスト）を一覧出力し、番号付与対象を目視で確認する |
| **apply** | JSON ルールファイルに従い、対象段落の先頭トークンを連番に差し替える |

## Non-goals

- SmartArt、表セル内テキスト、グループ化図形、ノート欄への番号付与
- PowerPoint の自動番号機能（ListStyle/BulletAutoNum）との連携
- Office Interop（COM オートメーション）の使用
- GUI の提供（コンソールアプリのみ）
- スライドの追加・削除・並べ替え
- 番号以外のテキスト内容の変更

## Current state summary

リポジトリは空の状態（README.md、LICENSE、.gitignore のみ）。
コードは存在しない。

## Proposed design / architecture delta

### プロジェクト構成

```
ruby-larch/
├── src/
│   └── PptxHeadlineNumbering/
│       ├── PptxHeadlineNumbering.csproj    # Console App (.NET 10 / net10.0)
│       ├── Program.cs                       # エントリポイント（CLI引数解析）
│       ├── InspectCommand.cs                # inspect モード実装
│       ├── ApplyCommand.cs                  # apply モード実装
│       ├── NumberingRule.cs                 # JSON ルールファイルのデシリアライズモデル
│       ├── SlideWalker.cs                   # スライド走査・段落分類のコアロジック
│       ├── ParagraphInfo.cs                 # 段落情報の値オブジェクト
│       ├── HeadingCounter.cs                # H1/H2/H3 カウンタ管理
│       └── PrefixReplacer.cs                # 先頭トークン検出・差し替えロジック
├── tests/
│   └── PptxHeadlineNumbering.Tests/
│       ├── PptxHeadlineNumbering.Tests.csproj
│       ├── SlideWalkerTests.cs
│       ├── HeadingCounterTests.cs
│       ├── PrefixReplacerTests.cs
│       ├── ApplyCommandTests.cs
│       └── TestData/                        # テスト用 .pptx ファイル
├── ruby-larch.sln
└── plans/
```

### NuGet 依存

| パッケージ | 用途 |
| --- | --- |
| `DocumentFormat.OpenXml` (3.x) | .pptx ファイルの読み書き |
| `System.CommandLine` (2.x) | CLI 引数パース |
| `System.Text.Json` | ルールファイル読み込み（SDK 組み込み） |
| `NUnit` / `NUnit3TestAdapter` / `Microsoft.NET.Test.Sdk` | テスト実行基盤 |

### コンポーネント設計

#### 1. Program.cs — CLI エントリポイント

`System.CommandLine` で以下のサブコマンドを定義する。

```
pptx-headline-numbering inspect <input.pptx>
pptx-headline-numbering apply <input.pptx> <output.pptx> --rule <rule.json>
```

- `inspect`: 読み取り専用。標準出力に TSV 形式で段落情報を出力する。
- `apply`: `input.pptx` を読み取り、番号を付与して `output.pptx` に書き出す。入力と出力が同じパスの場合はエラーとする。

#### 2. SlideWalker — スライド走査・段落分類

責務: PresentationDocument を受け取り、全スライドを順に走査して `ParagraphInfo` のシーケンスを返す。

```csharp
// 段落情報の値オブジェクト
public record ParagraphInfo(
    int SlideIndex,          // 0-based
    string ShapeName,        // sp:nvSpPr/nvCxnSpPr の name
    PlaceholderType? PhType, // title / ctrTitle / body / obj / null（自由配置図形）
    int ParagraphLevel,      // pPr/@lvl (既定=0)
    string CurrentText,      // 段落全体のプレーンテキスト
    // Open XML の段落要素への参照（apply 時に書き戻すため）
    DocumentFormat.OpenXml.Drawing.Paragraph ParagraphElement
);
```

判定ロジック:
- Shape の `nvSpPr > nvPr > ph` 要素から `type` 属性を取得 → `PlaceholderType`
- `ph/@type` が `title` または `ctrTitle` → 見出し1候補
- `ph/@type` が `body` または `obj`、かつ `pPr/@lvl` = 0 または未指定 → 見出し2候補
- `ph/@type` が `body` または `obj`、かつ `pPr/@lvl` = 1 → 見出し3候補
- 上記以外 → 番号付与対象外（inspect では出力するが apply ではスキップ）

#### 3. NumberingRule — ルールファイルモデル

```json
{
  "prefixRegex": "^[^\\s\\u3000]+(?:[\\s\\u3000]+)?",
  "separator": " ",
  "insertWhenPrefixMissing": true,
  "excludedSlideRanges": [
    { "startSlideNumber": 1, "endSlideNumber": 2 },
    { "startSlideNumber": 15, "endSlideNumber": 18 }
  ],
  "levels": [
    {
      "name": "H1",
      "match": { "placeholderTypes": ["title", "ctrTitle"] },
      "format": "{H1}.",
      "resetsOnNewLevel": []
    },
    {
      "name": "H2",
      "matches": [
        { "placeholderTypes": ["body", "obj"], "paragraphLevel": 0 },
        { "shapeNames": ["Content Placeholder 2", "コンテンツ プレースホルダー 2"], "paragraphLevel": 0 }
      ],
      "format": "{H1}.{H2}",
      "resetsOnNewLevel": ["H1"]
    },
    {
      "name": "H3",
      "match": { "placeholderTypes": ["body", "obj"], "paragraphLevel": 1 },
      "format": "{H3})",
      "resetsOnNewLevel": ["H1", "H2"]
    }
  ]
}
```

**改良点**:
- ユーザ提案の flat な `titleFormat` / `level1Format` / `level2Format` を `levels` 配列に一般化する。これにより 4段以上の見出しにも対応でき、マッチ条件も明示的になる。
- `match.shapeNames` を追加し、`PlaceholderType` が取れない段落でも `inspect` で見えた Shape 名で対象指定できるようにする。
- `matches` 配列を追加し、同一レベル名の中で複数条件を OR で持てるようにする。`name` はカウンタ名であるため重複禁止のままとし、複数レイアウト対応は `matches` で表現する。
- `excludedSlideRanges` を追加し、複数のスライド範囲を apply 対象外にできるようにする。範囲は `startSlideNumber` / `endSlideNumber` の 1-based・両端含む指定とする。

`NumberingRule` クラスは `System.Text.Json` のソース生成を使い、`JsonSerializerContext` 経由でデシリアライズする。

#### 4. HeadingCounter — カウンタ管理

- 各レベル名（`H1`, `H2`, `H3`, …）に対して int カウンタを保持する。
- `Increment(string levelName)` → 該当カウンタをインクリメントし、その後で「`resetsOnNewLevel` に `levelName` を含む他レベル」のカウンタを 0 にリセットする。これにより次回インクリメント時に 1 から再開する。
- `Format(string template)` → `{H1}` `{H2}` `{H3}` をカウンタ値に置換した文字列を返す。

#### 5. PrefixReplacer — 先頭トークン差し替え

責務: Open XML の `A.Paragraph` 要素内の先頭テキストから、`prefixRegex` にマッチする部分を検出し、新しい番号文字列 + `separator` に差し替える。

**Run 分割問題への対応**:
- Open XML では書式変更ごとに `A.Run` が分割されるため、先頭トークンが複数 Run にまたがる可能性がある。
- 対策: 段落の先頭から Run を連結してプレーンテキストを構築し、正規表現マッチの文字数分だけ Run チェーンから削除。その後、最初の Run の先頭に新しい番号テキストを挿入する。挿入テキストには最初の Run の `RunProperties` を継承する。
- `insertWhenPrefixMissing = true` の場合、マッチしなくても段落先頭に番号 + separator を挿入する。

#### 6. InspectCommand

1. `SlideWalker` で全段落を走査
2. 各 `ParagraphInfo` を TSV 行として標準出力に書き出す

出力フォーマット:
```
SlideIndex	ShapeName	PlaceholderType	Level	Text
0	Title 1	title	0	はじめに
0	Content Placeholder 2	body	0	背景の説明
0	Content Placeholder 2	body	1	詳細項目A
1	Title 1	title	0	次のトピック
```

#### 7. ApplyCommand

1. `NumberingRule` を JSON から読み込む
2. `HeadingCounter` を初期化
3. `SlideWalker` で全段落を走査
4. 各段落について:
   a. `excludedSlideRanges` に含まれるスライドはスキップする
   b. `levels` 配列を順にチェックし、各 level の `match` / `matches` を評価して最初にマッチしたレベルを採用
   c. `HeadingCounter.Increment(levelName)` でカウンタを進める
   d. `HeadingCounter.Format(format)` で番号文字列を生成
   e. `PrefixReplacer` で段落先頭トークンを差し替え
5. マッチしない段落はスキップ（変更しない）
6. `output.pptx` に保存

**安全性**: `Path.GetFullPath` で正規化した入力・出力パスを比較し、同一パスの場合はファイルを開く前にエラーとする。入力ファイルは読み取り専用で開き、出力ファイルは別パスに書き出す。

### エラーハンドリング方針

copilot-instructions.md の規約に従い:
- 処理失敗時はフォールバックせずエラー／例外を返す
- 全ての例外は `Exception.ToString()` をトレースログ（`System.Diagnostics.Trace`）に出力する
- ファイル I/O エラー、JSON パースエラー、Open XML 構造エラーはすべて例外として上位に伝播させる

## Coarse interaction scenarios

### シナリオ S-010: inspect モード — 正常系

1. ユーザが `pptx-headline-numbering inspect presentation.pptx` を実行
2. Program.cs が CLI 引数を解析し、InspectCommand を呼び出す
3. InspectCommand が PresentationDocument を読み取り専用で開く
4. SlideWalker が全スライドの全 Shape の全段落を走査し、ParagraphInfo を生成
5. TSV 形式で標準出力に出力
6. 終了コード 0 で終了

### シナリオ S-020: apply モード — 正常系

1. ユーザが `pptx-headline-numbering apply input.pptx output.pptx --rule rule.json` を実行
2. Program.cs が CLI 引数を解析し、ApplyCommand を呼び出す
3. ApplyCommand が rule.json を読み込み NumberingRule をデシリアライズ
4. HeadingCounter を初期化（全カウンタ = 0）
5. input.pptx を読み取り、メモリ上にコピーを作成
6. SlideWalker で全段落を走査
7. 各段落について levels 配列を順にマッチング判定
8. マッチした段落: カウンタをインクリメント → フォーマット → PrefixReplacer で差し替え
9. 全段落処理後、output.pptx として保存
10. 終了コード 0 で終了

### シナリオ S-030: apply モード — 既存番号の上書き（2回目以降の実行）

1. 前回 apply 済みの output.pptx を input.pptx として渡す
2. prefixRegex で既存の番号トークンがマッチ → 削除して新番号に差し替え
3. 結果として、何度実行しても同じルールなら同じ結果になる（冪等性）

### シナリオ S-040: inspect モード — ファイルが存在しないエラー

1. 存在しないファイルパスを指定
2. FileNotFoundException が発生
3. トレースログに例外を出力
4. 終了コード 1 + エラーメッセージを stderr に出力

### シナリオ S-050: apply モード — 入出力パスが同一

1. input と output に同じパスを指定
2. ApplyCommand が正規化済みフルパス比較でエラーを検出
3. ArgumentException を投げる
4. 終了コード 1 + エラーメッセージ

### シナリオ S-060: apply モード — ルールファイルの JSON が不正

1. 壊れた JSON を --rule に指定
2. JsonException が発生
3. トレースログに例外を出力
4. 終了コード 1

## Impacted code / files / modules

すべて新規作成:

| ファイルパス | 種別 | 責務 |
| --- | --- | --- |
| `src/PptxHeadlineNumbering/PptxHeadlineNumbering.csproj` | プロジェクト | .NET 10 Console App (`net10.0`) |
| `src/PptxHeadlineNumbering/Program.cs` | エントリポイント | CLI 引数解析、サブコマンド登録 |
| `src/PptxHeadlineNumbering/InspectCommand.cs` | コマンド | inspect モード |
| `src/PptxHeadlineNumbering/ApplyCommand.cs` | コマンド | apply モード |
| `src/PptxHeadlineNumbering/NumberingRule.cs` | モデル | JSON ルールファイルのデシリアライズ |
| `src/PptxHeadlineNumbering/SlideWalker.cs` | コアロジック | スライド走査・段落分類 |
| `src/PptxHeadlineNumbering/ParagraphInfo.cs` | 値オブジェクト | 段落情報 |
| `src/PptxHeadlineNumbering/HeadingCounter.cs` | ロジック | カウンタ管理・フォーマット |
| `src/PptxHeadlineNumbering/PrefixReplacer.cs` | ロジック | 先頭トークン差し替え |
| `tests/PptxHeadlineNumbering.Tests/PptxHeadlineNumbering.Tests.csproj` | テストプロジェクト | NUnit + `dotnet test` |
| `tests/PptxHeadlineNumbering.Tests/SlideWalkerTests.cs` | テスト | SlideWalker の UnitTest |
| `tests/PptxHeadlineNumbering.Tests/HeadingCounterTests.cs` | テスト | HeadingCounter の UnitTest |
| `tests/PptxHeadlineNumbering.Tests/PrefixReplacerTests.cs` | テスト | PrefixReplacer の UnitTest |
| `tests/PptxHeadlineNumbering.Tests/ApplyCommandTests.cs` | テスト | E2E 統合テスト |
| `tests/PptxHeadlineNumbering.Tests/TestData/` | テストデータ | テスト用 .pptx |
| `ruby-larch.sln` | ソリューション | |
| `sample-rule.json` | サンプル | ルールファイルのサンプル |

## Verification design

> 詳細ドキュメント:
> - ランタイムエビデンス: [plans/pptx-headline-numbering-runtime-evidence.md](pptx-headline-numbering-runtime-evidence.md)
> - テスト観点: [plans/pptx-headline-numbering-integration-test-points.md](pptx-headline-numbering-integration-test-points.md)

テストプロジェクトは NUnit を採用し、`.NET 10 SDK` 上で `dotnet test` により UnitTest / 統合テストを実行する。

### UnitTest

| テスト対象 | テスト内容 |
| --- | --- |
| `HeadingCounter` | カウンタのインクリメント、リセット、フォーマット置換 |
| `HeadingCounter` | 異なるリセットルールの組み合わせ |
| `PrefixReplacer` | 先頭トークンが単一 Run の場合の差し替え |
| `PrefixReplacer` | 先頭トークンが複数 Run にまたがる場合の差し替え |
| `PrefixReplacer` | プレフィックスが存在しない場合の挿入（insertWhenPrefixMissing=true） |
| `PrefixReplacer` | プレフィックスが存在しない場合のスキップ（insertWhenPrefixMissing=false） |
| `PrefixReplacer` | 全角スペースを含むプレフィックスの検出 |
| `SlideWalker` | title プレースホルダーの検出 |
| `SlideWalker` | ctrTitle プレースホルダーの検出 |
| `SlideWalker` | body プレースホルダー内の段落レベル判定 |
| `SlideWalker` | obj プレースホルダー内の段落レベル判定 |
| `SlideWalker` | プレースホルダー無しの自由配置図形の扱い |
| `NumberingRule` | `shapeNames` を使う JSON のデシリアライズ |
| `NumberingRule` | `matches` 配列による OR 条件のデシリアライズ |
| `NumberingRule` | `excludedSlideRanges` のデシリアライズと検証 |
| `NumberingRule` | 不正な JSON でのエラー |

### 統合テスト（自動・CIで実行可能）

| テスト内容 | 方法 |
| --- | --- |
| apply の冪等性 | テスト用 .pptx に対して apply → 再度 apply → 結果が同一であることを検証 |
| 複数スライドの連番 | 3スライドの .pptx に apply → 全スライドを inspect 相当で検証 |
| shapeName 指定の番号付与 | PlaceholderType が無い Shape に対し `shapeNames` で番号付与できることを検証 |
| 同一レベルの複数 OR 条件 | 1つの level の `matches` に複数条件を入れて同じカウンタで採番されることを検証 |
| 除外スライド範囲 | 複数のページ範囲を apply 対象外にし、その間は採番も進まないことを検証 |
| 空のスライドへの apply | 段落が無いスライドでエラーにならないことを検証 |
| inspect の出力形式 | 標準出力が期待 TSV に一致することを検証 |

### テスト用 .pptx の作成方針

テストフィクスチャとして、Open XML SDK を使ってプログラマティックに .pptx を生成するヘルパーメソッドを `TestData` 配下に用意する。これにより:
- テストが自己完結する（外部ファイル依存なし）
- CI で確実に再現できる
- テストケースごとに必要な構造だけを含む最小限の .pptx を生成できる

## Traceability matrix

| 要件 / 振る舞い | シナリオ | 検証方法 |
| --- | --- | --- |
| タイトルに見出し1番号を付与 | S-020 | UnitTest (SlideWalker, HeadingCounter) + 統合テスト |
| 箇条書き1段目に見出し2番号を付与 | S-020 | UnitTest (SlideWalker) + 統合テスト |
| 箇条書き2段目に見出し3番号を付与 | S-020 | UnitTest (SlideWalker) + 統合テスト |
| JSON ルールで番号書式を変更可能 | S-020 | UnitTest (NumberingRule, HeadingCounter) |
| shapeName 指定で PlaceholderType 無しの図形を対象化できる | S-020 | UnitTest (NumberingRule) + 統合テスト |
| 同一レベルに複数の OR 条件を持てる | S-020 | UnitTest (NumberingRule) + 統合テスト |
| 複数のスライド範囲を対象外にできる | S-020 | UnitTest (NumberingRule) + 統合テスト |
| 既存番号を上書きできる（冪等性） | S-030 | 統合テスト (apply 2回実行) |
| inspect で段落構造を確認できる | S-010 | 統合テスト (出力 TSV 検証) |
| 存在しないファイルでエラー | S-040 | 統合テスト |
| 入出力パス同一でエラー | S-050 | UnitTest (ApplyCommand) |
| JSON 不正でエラー | S-060 | UnitTest (NumberingRule) |
| Run 分割された先頭トークンの差し替え | S-020 | UnitTest (PrefixReplacer) |
| obj プレースホルダーの箇条書き対応 | S-010, S-020 | UnitTest (SlideWalker) |
| 例外時のトレースログ出力 | S-040, S-050, S-060 | UnitTest |

## Definition of Done

1. `inspect` コマンドが .pptx の全スライドの段落情報を TSV で出力できる
2. `apply` コマンドが JSON ルールに従い、対象段落の先頭に連番を付与できる
3. `apply` は冪等（同じルールで2回実行しても結果が同一）
4. JSON ルールファイルの差し替えだけで番号書式を変更できる
5. `shapeNames` 指定で PlaceholderType が取れない Shape も番号付与対象にできる
6. 同一レベル内で `matches` による OR 条件を指定できる
7. `excludedSlideRanges` 指定で複数のスライド範囲を番号付与対象外にできる
8. NUnit UnitTest が `.NET 10 SDK` 上の `dotnet test` で通る
9. 統合テストが CI の `.NET 10 SDK` 環境で通る
10. 入出力ファイルパスが同一の場合にエラーを返す
11. 存在しないファイル、不正な JSON に対して適切なエラーメッセージを返す
12. 例外発生時に `Trace` ログへ `Exception.ToString()` が出力される
13. ソースコード上のコメント・XMLコメントは日本語、ログ出力・コード本体は英語

## Risks / rollout / rollback

| リスク | 影響 | 軽減策 |
| --- | --- | --- |
| Run 分割パターンが想定外 | 番号の書式（太字・色）が崩れる | PrefixReplacer で最初の Run の RunProperties を継承。inspect で事前に確認。 |
| 会社テンプレートで body 以外に本文がある | 番号が付与されない | ルールファイルの `placeholderTypes` でカバー + inspect で確認 |
| Shape 名が Office の言語設定やテンプレートで異なる | `shapeNames` 指定が一致せず番号が付与されない | 事前に `inspect` で ShapeName を確認し、対象テンプレートごとに JSON を調整 |
| 除外スライド範囲の番号基準を誤解する | 想定外のページに番号が付く / 付かない | `excludedSlideRanges` は 1-based と README / sample-rule に明記し、inspect の `SlideIndex` とは別物であることを案内 |
| Open XML SDK のバージョン差異 | API 差異でビルド失敗 | DocumentFormat.OpenXml 3.x 系に固定 |
| 大規模ファイル（100スライド超） | 処理時間 | Open XML はインメモリ操作のため通常は問題なし。必要なら測定。 |

ロールバック: 出力は別ファイルに書き出すため、元ファイルは常に保全される。

## Open questions / assumptions

### Assumptions

1. .NET 10 (LTS) / `net10.0` をターゲットとする
2. `System.CommandLine` は 2.x (プレリリース含む) を使用する
3. テストプロジェクトは NUnit を採用し、`dotnet test` で実行する
4. テスト用 .pptx はプログラマティックに生成する（手作業の .pptx をリポジトリにコミットしない）
5. 空テキストの段落（改行のみ等）は番号付与対象外とする
6. スライドの走査順序は物理的なスライド番号順（PresentationPart.Presentation.SlideIdList の順序）

### Open questions

なし
