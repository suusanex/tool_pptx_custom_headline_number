using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxHeadlineNumbering.Tests;

public class PrefixReplacerTests
{
    private static readonly Regex NumberPrefixRegex = new(
        "^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
        RegexOptions.CultureInvariant);

    [Test]
    public void Replace_ReplacesPrefixInSingleRun()
    {
        var paragraph = CreateParagraph(["1. Intro"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, "2.", " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("2. Intro"));
        });
    }

    [Test]
    public void Replace_ReplacesPrefixAcrossMultipleRuns()
    {
        var paragraph = CreateParagraph(["1", ". ", "Intro"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, "3.", " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("3. Intro"));
            Assert.That(paragraph.Descendants<A.Run>().First().RunProperties?.Language?.Value, Is.EqualTo("en-US"));
        });
    }

    [Test]
    public void Replace_InsertsPrefixWhenMissing_IfConfigured()
    {
        var paragraph = CreateParagraph(["Intro"]);
        var numberingRegex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, numberingRegex, "1.", " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("1. Intro"));
        });
    }

    [Test]
    public void Replace_DoesNothingWhenMissingAndInsertionDisabled()
    {
        var paragraph = CreateParagraph(["Intro"]);
        var numberingRegex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, numberingRegex, "1.", " ", false);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.False);
            Assert.That(GetText(paragraph), Is.EqualTo("Intro"));
        });
    }

    [Test]
    public void Replace_RemovesPrefixWithoutInserting_WhenNewPrefixIsEmpty()
    {
        var paragraph = CreateParagraph(["1. Intro"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, string.Empty, " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("Intro"));
        });
    }

    [Test]
    public void Replace_MultipleRuns_RemovesPrefixOnlyWhenNewPrefixIsEmpty()
    {
        var paragraph = CreateParagraph(["1", ". ", "Intro"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, string.Empty, " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("Intro"));
        });
    }

    [Test]
    public void Replace_DoesNothingForEmptyParagraph_WhenNewPrefixIsEmptyAndInsertEnabled()
    {
        var paragraph = new A.Paragraph();
        var regex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, regex, string.Empty, " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.False);
            Assert.That(paragraph.Descendants<A.Text>().Any(), Is.False);
        });
    }

    [Test]
    public void UT_IT_040__Replace_RemovesPrefixOnlyWhenNewPrefixIsEmpty()
    {
        // format:"" のとき、既存プレフィックスのみ除去し separator は挿入しない（TP-040 prefix 削除分岐）
        var paragraph = CreateParagraph(["1. はじめに"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, string.Empty, " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("はじめに")); // prefix と後続スペースが除去され本文のみ残る
        });
    }

    [Test]
    public void UT_IT_040__Replace_DoesNothingWhenPrefixMissingAndNewPrefixIsEmptyEvenIfInsertEnabled()
    {
        // prefix 未検出かつ newPrefix="" の場合、insertWhenPrefixMissing=true でも何も挿入しない
        var paragraph = CreateParagraph(["Intro"]);
        var regex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, regex, string.Empty, " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.False);
            Assert.That(GetText(paragraph), Is.EqualTo("Intro"));
        });
    }

    [Test]
    public void Replace_SupportsFullWidthSeparator()
    {
        var paragraph = CreateParagraph(["1.　Intro"]);

        var changed = new PrefixReplacer().Replace(paragraph, NumberPrefixRegex, "2.", "　", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("2.　Intro"));
        });
    }

    [Test]
    public void UT_IT_040__Replace_InsertsIntoEmptyParagraphWhenEnabled()
    {
        // テキストノードが 0 件の段落で insertWhenPrefixMissing=true のとき、番号 Run が先頭に挿入される
        var paragraph = new A.Paragraph(); // 空の段落（Run なし）
        var regex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, regex, "1.", " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("1. "));
        });
    }

    [Test]
    public void UT_IT_040__Replace_DoesNothingForEmptyParagraphWhenDisabled()
    {
        // テキストノードが 0 件の段落で insertWhenPrefixMissing=false のとき、変更されない
        var paragraph = new A.Paragraph(); // 空の段落（Run なし）
        var regex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, regex, "1.", " ", false);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.False);
            Assert.That(paragraph.Descendants<A.Text>().Any(), Is.False);
        });
    }

    [Test]
    public void Replace_InsertsAfterParagraphPropertiesWhenPPrExists()
    {
        // a:pPr が存在する空段落で insertWhenPrefixMissing=true のとき、Run が a:pPr の後ろに挿入される
        var paragraph = new A.Paragraph(new A.ParagraphProperties { Level = 1 });
        var regex = new Regex("^\\d+[.)](?:[\\s\\u3000]+)?", RegexOptions.CultureInvariant);

        var changed = new PrefixReplacer().Replace(paragraph, regex, "1.", " ", true);

        Assert.Multiple(() =>
        {
            Assert.That(changed, Is.True);
            Assert.That(GetText(paragraph), Is.EqualTo("1. "));
            Assert.That(paragraph.GetFirstChild<A.ParagraphProperties>(), Is.Not.Null);
            Assert.That(paragraph.GetFirstChild<A.ParagraphProperties>()!.NextSibling<A.Run>(), Is.Not.Null);
        });
    }

    private static A.Paragraph CreateParagraph(IEnumerable<string> runs)
    {
        var paragraph = new A.Paragraph();
        foreach (var text in runs)
        {
            paragraph.Append(
                new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(text)));
        }

        return paragraph;
    }

    private static string GetText(A.Paragraph paragraph)
    {
        return string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text));
    }
}
