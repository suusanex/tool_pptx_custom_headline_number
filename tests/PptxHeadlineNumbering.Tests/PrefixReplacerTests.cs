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
