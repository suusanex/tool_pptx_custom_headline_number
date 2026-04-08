# 実装カバレッジ: PowerPoint カスタム見出し連番ツール

> 元テスト観点: `plans/pptx-headline-numbering-integration-test-points.md`
> 対象実装: `src/PptxHeadlineNumbering/`
> 生成日: 2026-04

---

## サマリ

| 状態 | 件数 |
|---|---|
| Automated | 10 |
| RecordedButSkipped | 0 |
| ManualOnly | 0 |
| NotImplementedOrMismatch | 0 |
| **合計** | **10** |

---

## 各 ID の状態

### TP-010: inspect モード — TSV 出力の正確性

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `InspectCommandTests.Execute_WritesExpectedTsv`<br>`InspectCommandTests.UT_IT_010__Execute_OutputsCtrTitleAndEmptyParagraphAndHighLevel`<br>`SlideWalkerTests.Walk_ReturnsMetadataForAllParagraphKinds` |
| **理由** | TSV ヘッダー・各フィールド（SlideIndex/ShapeName/PlaceholderType/Level/Text）を検証済み。ctrTitle プレースホルダー・空テキスト段落・level≥2 の段落出力は `UT_IT_010__*` で追加確認。自由配置図形（null プレースホルダー）の出力も `Execute_WritesExpectedTsv` で確認済み。実装は `InspectCommand.cs` / `SlideWalker.cs` が存在し本物の実装を使用。 |

---

### TP-020: apply モード — 番号付与の正確性

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `ApplyCommandTests.Execute_AppliesNumberingAcrossSlides_AndIsIdempotent`<br>`ApplyCommandTests.UT_IT_020__Execute_AppliesFullWidthSeparator`<br>`ApplyCommandTests.UT_IT_020__Execute_NumbersObjPlaceholderAndSkipsFreeShape` |
| **理由** | 複数スライドにまたがる H1/H2/H3 連番・separator 動作を `Execute_AppliesNumberingAcrossSlides_AndIsIdempotent` で検証。全角スペース separator は `UT_IT_020__Execute_AppliesFullWidthSeparator` で確認。obj プレースホルダーの番号付与と自由配置図形のスキップは `UT_IT_020__Execute_NumbersObjPlaceholderAndSkipsFreeShape` で確認。実装は `ApplyCommand.cs` / `NumberingRule.cs` が存在し本物の実装を使用。 |

---

### TP-030: apply モード — カウンタリセットの正確性

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `HeadingCounterTests.Increment_ResetsDependentCounters`<br>`HeadingCounterTests.UT_IT_030__H2IncrementResetsH3`<br>`HeadingCounterTests.UT_IT_030__H1OnlyConsecutiveMonotonicallyIncreases`<br>`HeadingCounterTests.UT_IT_030__H3OnlyConsecutiveMonotonicallyIncreases`<br>`HeadingCounterTests.UT_IT_030__ConstructorThrowsWhenLevelsEmpty` |
| **理由** | H1 インクリメントで H2/H3 がリセットされることは既存テストで確認済み。H2 インクリメントで H3 のみリセット、H1 のみ単調増加、H3 のみ単調増加（他カウンタに影響しない）を追加テストで確認。`resetsOnNewLevel:[]` パターンも H1-only テストで実質カバー。実装は `HeadingCounter.cs` が本物の実装として存在。 |

---

### TP-040: apply モード — 先頭トークン差し替え（Run分割対応）

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `PrefixReplacerTests.Replace_ReplacesPrefixInSingleRun`<br>`PrefixReplacerTests.Replace_ReplacesPrefixAcrossMultipleRuns`<br>`PrefixReplacerTests.Replace_InsertsPrefixWhenMissing_IfConfigured`<br>`PrefixReplacerTests.Replace_DoesNothingWhenMissingAndInsertionDisabled`<br>`PrefixReplacerTests.Replace_SupportsFullWidthSeparator`<br>`PrefixReplacerTests.UT_IT_040__Replace_InsertsIntoEmptyParagraphWhenEnabled`<br>`PrefixReplacerTests.UT_IT_040__Replace_DoesNothingForEmptyParagraphWhenDisabled` |
| **理由** | 単一 Run / 複数 Run / insertWhenPrefixMissing=true/false / 全角セパレータの各ケースを既存テストで検証済み。テキストノードが 0 件（空段落）の分岐は `UT_IT_040__*` で追加確認。RunProperties 継承は `Replace_ReplacesPrefixAcrossMultipleRuns` で `RunProperties.Language` を確認済み。実装は `PrefixReplacer.cs` が本物の実装として存在。 |

---

### TP-050: apply モード — 冪等性

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `ApplyCommandTests.Execute_AppliesNumberingAcrossSlides_AndIsIdempotent` |
| **理由** | output1.pptx と output2.pptx（output1 を入力として再適用）の全段落テキストが完全一致することを検証済み。2回分の冪等性が証明できれば N 回も同様なため、3回連続は本テストで実質カバーされる（TP-100 の3回連続も同一）。本物の実装（`ApplyCommand` + `PrefixReplacer` + `HeadingCounter`）を使用。 |

---

### TP-060: ファイル I/O エラー

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `CliApplicationTests.Run_ReturnsErrorForMissingInputFile`（inspect 不存在ファイル → exit 1）<br>`InspectCommandTests.UT_IT_060__Execute_TracesExceptionOnFileNotFound`（trace ログ確認）<br>`ApplyCommandTests.UT_IT_060__Execute_ThrowsForNonexistentRuleFile`（rule ファイル不存在 → 例外）<br>`CliApplicationTests.UT_IT_060__Run_ReturnsErrorWhenApplyInputNotFound`（apply 不存在入力 → exit 1）<br>`CliApplicationTests.UT_IT_060__Run_ReturnsErrorForCorruptPptxFile`（壊れた ZIP → exit 1）<br>`ApplyCommandTests.Execute_ThrowsWhenInputAndOutputAreSame_AndWritesTrace`（trace ログ確認） |
| **理由** | ファイル不存在（inspect/apply 入力・rule）→ 例外発生・exit 1・trace ログ出力をそれぞれ確認。壊れた .pptx（非 ZIP バイト列）も exit 1 を確認。stderr への出力は CliApplicationTests.Run_ReturnsErrorForMissingInputFile で exit 1 と確認（stderr 内容は Run_Returns... で検証）。実装は `InspectCommand.cs` / `ApplyCommand.cs` に catch ブロックが存在。 |

---

### TP-070: CLI 引数・パス検証エラー

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `CliApplicationTests.Run_ReturnsErrorForMissingArguments`（引数なし → exit 1）<br>`CliApplicationTests.Run_ReturnsErrorWhenRuleOptionMissing`（--rule 未指定 → exit 1）<br>`ApplyCommandTests.Execute_ThrowsWhenInputAndOutputAreSame_AndWritesTrace`（同一パス → ArgumentException）<br>`CliApplicationTests.UT_IT_070__Run_ReturnsErrorForUnknownCommand`（不明コマンド → exit 1）<br>`CliApplicationTests.UT_IT_070__Run_ReturnsErrorWhenNormalizedApplyPathsAreSame`（パス正規化後同一 → exit 1） |
| **理由** | 引数なし・--rule 未指定・不明コマンドは exit 1 + エラーメッセージを確認。入出力パス同一（ArgumentException + trace ログ）確認済み。パス正規化（`./input.pptx` vs `input.pptx` → `Path.GetFullPath` 後同一）は `UT_IT_070__Run_ReturnsErrorWhenNormalizedApplyPathsAreSame` で確認。実装は `CliApplication.cs` / `ApplyCommand.cs` が本物の実装として存在。 |

---

### TP-080: JSON ルール不正

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `NumberingRuleTests.LoadFromFile_ThrowsForBrokenJson`（JSON 構文エラー → JsonException）<br>`NumberingRuleTests.UT_IT_080__LoadFromFile_ThrowsForEmptyFile`（空ファイル → JsonException）<br>`NumberingRuleTests.UT_IT_080__LoadFromFile_ThrowsForEmptyLevels`（levels=[] → InvalidDataException）<br>`NumberingRuleTests.UT_IT_080__LoadFromFile_ThrowsForInvalidPrefixRegex`（不正 regex → RegexParseException ⊂ ArgumentException）<br>`NumberingRuleTests.UT_IT_080__LoadFromFile_ThrowsWhenFormatMissing`（format 未指定 → InvalidDataException） |
| **理由** | JSON 構文エラー・空ファイル（JsonException）、levels=[]・format 欠損（InvalidDataException）、不正な prefixRegex（RegexParseException、ArgumentException のサブクラス）をそれぞれ確認。trace ログへの記録は `LoadFromFile` が全例外を catch してから re-throw する実装を確認（`Trace.WriteLine(ex.ToString())`）。実装は `NumberingRule.cs` の `LoadFromFile` + `Validate` + `BuildPrefixRegex` が存在。 |

---

### TP-090: 大量データの処理

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `SlideWalkerTests.UT_IT_090__Walk_HandlesLargeNumberOfSlides`（55枚スライド × 6段落 = 330段落を正常処理）<br>`SlideWalkerTests.UT_IT_090__Walk_HandlesEmptySlidesMixedIn`（空スライド混在 → SlideIndex が正しく継続）<br>`ApplyCommandTests.UT_IT_090__Execute_HandlesLargeNumberOfSlidesAndEmptySlides`（55枚スライド・空スライド混在で apply 完走、連番継続） |
| **理由** | 55枚スライド（1スライドあたり title 1段落 + body 5段落）を `PptxTestDocumentFactory` で生成し、全段落数が正しく取得できることを確認。空スライド混在時に SlideIndex が継続することも確認。さらに `UT_IT_090__Execute_HandlesLargeNumberOfSlidesAndEmptySlides` で、55枚スライド + 空スライド混在の入力に対して apply が正常完走し、最初と最後の連番が期待どおり継続することを確認。速度・メモリの閾値自体は仕様化されていないため、本 coverage では機能面の大量データ耐性として Automated とした。 |

---

### TP-100: 連続実行・再実行

| 項目 | 内容 |
|---|---|
| **状態** | Automated |
| **対応テスト** | `ApplyCommandTests.UT_IT_100__Execute_SwitchesRuleCorrectly`（ルールA→ルールB差し替え確認）<br>`ApplyCommandTests.UT_IT_100__Execute_ApplyResultVisibleInInspect`（apply 後 inspect で番号付きテキスト確認）<br>`ApplyCommandTests.UT_IT_100__Execute_RemainsStableAcrossThreeRuns`（同一ルール 3 回連続実行で出力不変） |
| **理由** | ルールA（`{H1}.` / `{H1}.{H2}`）でナンバリング後、ルールB（`第{H1}章` / `{H1}-{H2}`）で再適用すると旧番号が prefixRegex で除去され新番号に差し替わることを確認。apply 後に inspect すると TSV の Text 列に番号付きテキストが反映されることも確認。さらに `UT_IT_100__Execute_RemainsStableAcrossThreeRuns` で同一ルールの 3 回連続実行でも結果が変化しないことを直接確認した。実装は `ApplyCommand` + `PrefixReplacer` + `NumberingRule` が本物の実装として存在。 |

---

## 保留・未対応サマリ

### RecordedButSkipped (0件)

なし。

### ManualOnly (0件)

なし。なお TP-090 の速度・メモリ使用量の実測は、必要なら別途手動または性能テスト基盤で補う余地がある。

### NotImplementedOrMismatch (0件)

なし。全 ID について対応実装ファイルが `src/PptxHeadlineNumbering/` に存在することを確認。

---

## スタブ・本物の実装確認

本プロジェクトのすべての UnitTest は、モック・スタブ・InMemory 実装を使用せず、`src/PptxHeadlineNumbering/` の本物の実装クラス（`SlideWalker`, `InspectCommand`, `ApplyCommand`, `PrefixReplacer`, `HeadingCounter`, `NumberingRule`）を直接使用している。DI 配線も `CliApplication` コンストラクタ内でデフォルト具象クラスがインスタンス化されており、本番相当の経路から各責務に到達可能であることを確認済み。

---

## 変更した主なファイル

| ファイル | 変更内容 |
|---|---|
| `tests/.../HeadingCounterTests.cs` | UT_IT_030 × 4 テスト追加 |
| `tests/.../NumberingRuleTests.cs` | UT_IT_080 × 4 テスト追加 |
| `tests/.../PrefixReplacerTests.cs` | UT_IT_040 × 2 テスト追加 |
| `tests/.../InspectCommandTests.cs` | UT_IT_010 × 1 / UT_IT_060 × 1 テスト追加 |
| `tests/.../SlideWalkerTests.cs` | UT_IT_090 × 2 テスト追加 |
| `tests/.../ApplyCommandTests.cs` | UT_IT_020 × 2 / UT_IT_060 × 1 / UT_IT_090 × 1 / UT_IT_100 × 3 テスト追加 |
| `tests/.../CliApplicationTests.cs` | UT_IT_070 × 2 / UT_IT_060 × 2 テスト追加 |
