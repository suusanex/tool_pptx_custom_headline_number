# prefix 削除機能追加

## Background / Goal

`apply` コマンドで一度付与した prefix（番号文字列と区切り文字）を、番号変更ではなく**完全に削除**したいケースが発生している。

現状の `apply` コマンドは、各レベルの `format` を常に連番文字列に置換するため、prefix を取り除いて元の見出しテキストに戻すことができない。

## Non-goals

- prefix 削除専用の独立コマンド（`strip` サブコマンド等）の追加
- 削除後のテキストそのものの編集
- `inspect` コマンドへの変更
- `prefixRegex` の変更（既存の除去ロジックはそのまま利用する）

## Current state summary

### 変更対象となる主要な制約

| 箇所 | 現在の動作 | 問題点 |
|---|---|---|
| `NumberingRule.Validate()` | `format` 未指定と空文字列を区別せず、空文字列を一律で `InvalidDataException` にする | 明示的な削除指定を表現できない |
| `PrefixReplacer.Replace()` | 常に `newPrefix + separator` を先頭に挿入する | prefix を付けずに既存 prefix だけ除去できない |
| `ApplyCommand.Execute()` | `HeadingCounter.Format(level.Format)` を前提にしている | `format: ""` を許容すると空テンプレートで例外化する |

### 2案の比較

| 案 | 概要 | 変更クラス数 | 変更量 | 障害リスク |
|---|---|---|---|---|
| **Option 1** | `strip` CLIオプション追加 | 5 以上（CLI, コマンドクラス, ApplyCommand or 新コマンド等） | 大 | 中〜高 |
| **Option 2** | `"format": ""` で prefix 削除を表現 | 3（NumberingRule, PrefixReplacer, ApplyCommand） | 小 | 低 |

**Option 2 を採用する。** ルールファイルの書き心地と既存アーキテクチャへの侵食が最小限で、単一の `apply` コマンドで「付与」と「削除」を同一ルールファイルで一括指定できる。

## Proposed design / architecture delta

### ルールファイル（`rule.json`）

`levels` の各エントリで `"format": ""` を指定すると、そのレベルにマッチした段落の prefix を削除し、何も挿入しない。

削除される範囲は `PrefixReplacer` が `prefixRegex` の**先頭一致部分**として検出した文字列である。したがって、番号の直後にある区切り文字も削除したい場合は、既存の `sample-rule.json` と同様に `prefixRegex` 側で末尾の空白や全角空白まで含めてマッチさせる。

```json
{
  "levels": [
    {
      "name": "H1",
      "match": { "placeholderTypes": ["title", "ctrTitle"] },
      "format": "",
      "resetsOnNewLevel": []
    }
  ]
}
```

### `NumberingLevelRule.Format` / `NumberingRule.Validate()` の変更

`format` の意味を次の 3 値で明確化する。

- `null`（JSON で `format` 未指定）: 不正。設定漏れとして `InvalidDataException` を維持する
- `""`（長さ 0 の空文字）: 正常。prefix 削除指示として扱う
- 空白のみ（例: `" "`, `"　"`）: 不正。削除なのか空白 prefix 付与なのか曖昧なため `InvalidDataException`

この区別を実現するため、`NumberingLevelRule.Format` は nullable に変更し、`Validate()` で `null` / 空文字 / 空白のみを判定する。ルールの論理的整合性（`name` 必須・`match` 必須）は引き続き検証する。

```diff
- public string Format { get; init; } = string.Empty;
+ public string? Format { get; init; }

- if (level.Format is null)
- {
-     throw new InvalidDataException(
-         string.Format(CultureInfo.InvariantCulture, "format is required for level: {0}", level.Name));
- }
+
+ if (level.Format.Length > 0 && string.IsNullOrWhiteSpace(level.Format))
+ {
+     throw new InvalidDataException(
+         string.Format(CultureInfo.InvariantCulture, "format must be empty or non-whitespace for level: {0}", level.Name));
+ }
  level.Validate();
```

### `PrefixReplacer.Replace()` の変更

`newPrefix` が空文字列の場合は `separator` も付加せず、既存 prefix の除去のみ行う。`insertWhenPrefixMissing` が `false` かつ既存 prefix がない場合は何もしない（既存動作と同じ）。

```diff
  RemoveLeadingCharacters(textNodes, removeLength);
  var firstTextNode = textNodes[0];
- firstTextNode.Text = newPrefix + separator + firstTextNode.Text;
+ if (newPrefix.Length > 0)
+ {
+     firstTextNode.Text = newPrefix + separator + firstTextNode.Text;
+ }
```

テキストノードが 0 件（空段落）かつ `insertWhenPrefixMissing = true` でも、`newPrefix` が空の場合は何も挿入しない。

```diff
  if (textNodes.Count == 0)
  {
      if (!insertWhenPrefixMissing)
      {
          return false;
      }
+
+     if (newPrefix.Length == 0)
+     {
+         return false;
+     }
+
      var run = new A.Run();
```

### `ApplyCommand.Execute()` の変更

`format` が空文字列のレベルでも、見出し階層の構造は維持するため**カウンタ増分と `resetsOnNewLevel` は通常通り実行する**。ただし空テンプレートを `HeadingCounter.Format()` に渡すと例外になるため、フォーマット呼び出しだけをスキップする。

これにより、「H1 は prefix を消すが、H2 は H1 ごとに `1.1`, `2.1` と振り直す」のような混在ルールでも期待通りに動作する。

```diff
  var level = rule.MatchLevel(...);
  if (level is null) { continue; }

  counter.Increment(level.Name);

  var formattedPrefix = level.Format.Length == 0
      ? string.Empty
      : counter.Format(level.Format);
  _prefixReplacer.Replace(...);
```

## Coarse interaction scenarios

### S-DEL-010: prefix 削除の正常系

1. ユーザーが `format: ""` を含む `rule.json` を用意して `apply` を実行する
2. `NumberingRule.LoadFromFile` がルールをデシリアライズ・`Validate()` を通過する
3. 各段落に対して `MatchLevel` がレベルを返す
4. `level.Format.Length == 0` の場合でもカウンタ増分は行い、`formattedPrefix = ""` を渡す
5. `PrefixReplacer.Replace` が `prefixRegex` にマッチした部分を除去し、separator は付加せずに終了する
6. 対象段落は元の本文テキストのみに戻る

### S-DEL-015: 削除レベルと通常レベルの混在

1. H1 に `format: ""`、H2 に `format: "{H1}.{H2}"` を設定した `rule.json` を適用する
2. H1 段落では既存 prefix のみが削除され、表示上は素のタイトルテキストになる
3. H1 にマッチした時点で `HeadingCounter` の H1 カウンタは増分され、H2 のリセットも従来通り発生する
4. 後続の H2 段落では `1.1`, `1.2`, 次の H1 配下では `2.1` のように、非表示の H1 階層を基準とした連番が維持される

### S-DEL-020: prefix が存在しない段落への削除

- `prefixRegex` がマッチしない段落で `insertWhenPrefixMissing = false` の場合: 何も変更しない（既存動作と同じ）
- `insertWhenPrefixMissing = true` かつ `format = ""` の場合: 挿入するものがないため何もしない（新規動作）

### S-DEL-040: ルール不正の失敗系

- `format` 未指定: 設定漏れとして `InvalidDataException`
- `format` が空白のみ: 曖昧な設定として `InvalidDataException`
- `format` が空文字: 正常として削除モードで処理継続

### S-DEL-030: 冪等性の確保

一度 prefix を削除した `.pptx` に対して同じルールで再度 `apply` を実行した場合:
- `prefixRegex` がマッチしない → `insertWhenPrefixMissing = false` なら何も変わらない（冪等）
- `insertWhenPrefixMissing = true`・`format = ""` の場合も何も挿入されないためやはり冪等

## Impacted code / files / modules

| ファイル | 変更内容 |
|---|---|
| `src/PptxHeadlineNumbering/NumberingRule.cs` | `NumberingLevelRule.Format` を nullable 化し、未指定は不正・空文字は削除・空白のみは不正として検証 |
| `src/PptxHeadlineNumbering/PrefixReplacer.cs` | `newPrefix` が空の場合は separator も付加しない、空段落でも挿入しない |
| `src/PptxHeadlineNumbering/ApplyCommand.cs` | `format` が空でもカウンタ増分は維持しつつ、`HeadingCounter.Format()` 呼び出しだけをスキップ |
| `tests/PptxHeadlineNumbering.Tests/NumberingRuleTests.cs` | 空 format バリデーション通過の新テスト追加 |
| `tests/PptxHeadlineNumbering.Tests/PrefixReplacerTests.cs` | newPrefix="" 時の動作テスト追加 |
| `tests/PptxHeadlineNumbering.Tests/ApplyCommandTests.cs` | format="" レベルを含むルールでの統合テスト追加 |
| `README.md` | prefix 削除の使用方法を追記 |
| `sample-rule.json` | 変更不要（例示にはデフォルトの非空 format を使い続ける） |

## Verification design

### 単体テスト（新規追加）

| テストクラス | テスト名（案） | 検証内容 |
|---|---|---|
| `NumberingRuleTests` | `LoadFromFile_AllowsExplicitEmptyFormat` | `format: ""` のルールが `LoadFromFile()` / `Validate()` を通過する |
| `NumberingRuleTests` | `UT_IT_080__LoadFromFile_ThrowsWhenFormatMissing` | `format` 未指定は引き続き不正として拒否される |
| `NumberingRuleTests` | `LoadFromFile_ThrowsWhenFormatIsWhitespaceOnly` | `format: " "` や `"　"` は不正として拒否される |
| `PrefixReplacerTests` | `Replace_RemovesPrefixWithoutInserting_WhenNewPrefixIsEmpty` | 既存 prefix がある段落で `newPrefix=""` を渡すと prefix だけ除去される |
| `PrefixReplacerTests` | `Replace_DoesNothingForEmptyParagraph_WhenNewPrefixIsEmptyAndInsertEnabled` | 空段落で `insertWhenPrefixMissing=true` かつ `newPrefix=""` のとき何も挿入されない |
| `PrefixReplacerTests` | `Replace_MultipleRuns_RemovesPrefixOnlyWhenNewPrefixIsEmpty` | 複数 Run にまたがる prefix を `newPrefix=""` で除去できる |
| `ApplyCommandTests` | `Execute_RemovesPrefixWhenFormatIsEmpty` | `format: ""` のルールで apply すると、`prefixRegex` に含めた separator も含めて対象段落の prefix が削除された pptx が出力される |
| `ApplyCommandTests` | `Execute_PreservesHierarchyWhenDeletionAndNumberingAreMixed` | H1 削除 + H2 通常採番の混在ルールで H2 が H1 増分・リセットを正しく反映する |
| `ApplyCommandTests` | `Execute_IsIdempotent_WhenFormatIsEmpty` | 削除後の pptx に再度同ルールで apply しても変化しない（冪等） |

### 既存テストへの影響

- 既存の `UT_IT_080__LoadFromFile_ThrowsWhenFormatMissing` は維持する。`format` 未指定は引き続き不正であり、期待値は変えない
- 既存の `PrefixReplacerTests` は `newPrefix` が非空のケースなので影響なし

## Traceability matrix

| 要件 / 期待動作 | シナリオ | 検証方法 |
|---|---|---|
| `format: ""` のルールが JSON 検証を通過する | S-DEL-010 | `NumberingRuleTests.LoadFromFile_AllowsExplicitEmptyFormat` |
| `format` 未指定は設定漏れとして拒否される | S-DEL-040 | 既存テスト `UT_IT_080__LoadFromFile_ThrowsWhenFormatMissing` |
| `format` が空白のみのときは曖昧設定として拒否される | S-DEL-040 | `NumberingRuleTests.LoadFromFile_ThrowsWhenFormatIsWhitespaceOnly` |
| prefix がある段落で `format: ""` を指定すると prefix のみ除去される | S-DEL-010 | `PrefixReplacerTests.Replace_RemovesPrefixWithoutInserting_WhenNewPrefixIsEmpty` |
| `prefixRegex` が separator まで含む場合は区切り文字ごと削除される | S-DEL-010 | `ApplyCommandTests.Execute_RemovesPrefixWhenFormatIsEmpty` |
| prefix がマッチしない段落は変更されない | S-DEL-020 | 既存テスト `Replace_DoesNothingWhenMissingAndInsertionDisabled` |
| 空段落で `insertWhenPrefixMissing=true` かつ `format=""` のとき何も起こらない | S-DEL-020 | `PrefixReplacerTests.Replace_DoesNothingForEmptyParagraph_WhenNewPrefixIsEmptyAndInsertEnabled` |
| 削除レベルと通常レベルを混在させても階層カウンタは維持される | S-DEL-015 | `ApplyCommandTests.Execute_PreservesHierarchyWhenDeletionAndNumberingAreMixed` |
| 削除後のファイルへの再 apply が冪等 | S-DEL-030 | `ApplyCommandTests.Execute_IsIdempotent_WhenFormatIsEmpty` |

## Definition of Done

- [ ] `NumberingRule.Validate()` が `format: ""` を正常に通過する
- [ ] `NumberingRule.Validate()` が `format` 未指定を引き続き拒否し、空白のみも拒否する
- [ ] `PrefixReplacer.Replace()` が `newPrefix=""` のとき既存 prefix を除去し何も挿入しない
- [ ] `ApplyCommand.Execute()` が `format: ""` のレベルでも階層カウンタとリセットを維持する
- [ ] 上記に対応する単体テストが全てパスしている
- [ ] 既存テストが全てパスしている（デグレードなし）
- [ ] `README.md` に prefix 削除の使用方法と、separator ごと削除するには `prefixRegex` が末尾区切り文字を含む必要があることが追記されている

## Risks / rollout / rollback

| リスク | 対策 |
|---|---|
| `format` 未指定まで削除モードとして扱ってしまう設定事故 | `Format` を nullable 化し、未指定は引き続き `InvalidDataException` にする |
| `insertWhenPrefixMissing = true` かつ `format = ""` で意図せずテキストが崩れる | 空段落に対して何もしない分岐を `PrefixReplacer` に追加済み（設計通り） |
| separator ごとの削除を期待しているのに `prefixRegex` が separator を含まない | README とルール例で、削除範囲が `prefixRegex` 先頭一致部分であることを明記する |

ロールバックは `apply` の実行元 `.pptx` を保持しておけば再実行で復元可能（既存の設計方針と同じ）。

## Open questions / assumptions

なし。
