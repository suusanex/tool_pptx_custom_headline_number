using System.Diagnostics;
using PptxHeadlineNumbering.Tests.TestData;

namespace PptxHeadlineNumbering.Tests;

public class ApplyCommandTests
{
    private static readonly string RuleJson =
        """
        {
          "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
          "separator":" ",
          "insertWhenPrefixMissing":true,
          "levels":[
            {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]},
            {"name":"H2","match":{"placeholderTypes":["body","obj"],"paragraphLevel":0},"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]},
            {"name":"H3","match":{"placeholderTypes":["body","obj"],"paragraphLevel":1},"format":"{H3})","resetsOnNewLevel":["H1","H2"]}
          ]
        }
        """;

    [Test]
    public void Execute_AppliesNumberingAcrossSlides_AndIsIdempotent()
    {
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var output1Path = Path.Combine(tempDir, "output1.pptx");
        var output2Path = Path.Combine(tempDir, "output2.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Background"), new TestParagraph(1, "Detail A")),
                new TestShape("Free Text", null, new TestParagraph(0, "Free Text"))),
            new TestSlide(
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Topic")),
                new TestShape("Body 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Agenda"), new TestParagraph(1, "Detail B"))));
        File.WriteAllText(rulePath, RuleJson);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, output1Path, rulePath);
            command.Execute(output1Path, output2Path, rulePath);

            var output1Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output1Path);
            var output2Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output2Path);

            Assert.Multiple(() =>
            {
                Assert.That(
                    output1Texts,
                    Is.EqualTo(new[]
                    {
                        "1. Intro",
                        "1.1 Background",
                        "1) Detail A",
                        "Free Text",
                        "2. Topic",
                        "2.1 Agenda",
                        "1) Detail B",
                    }));
                Assert.That(output2Texts, Is.EqualTo(output1Texts));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_RemovesPrefixWhenFormatIsEmpty()
    {
        const string ruleJsonForRemoval =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"","resetsOnNewLevel":[]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "1. Intro"))),
            new TestSlide(
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "2. Topic"))));
        File.WriteAllText(rulePath, ruleJsonForRemoval);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(texts, Is.EqualTo(new[] { "Intro", "Topic" }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_020__Execute_RemovesPrefixFromCenteredTitleWhenFormatIsEmpty()
    {
        const string ruleJsonForRemoval =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"","resetsOnNewLevel":[]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Cover", DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle,
                    new TestParagraph(0, "1. Cover"))));
        File.WriteAllText(rulePath, ruleJsonForRemoval);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(texts, Is.EqualTo(new[] { "Cover" }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_PreservesHierarchyWhenDeletionAndNumberingAreMixed()
    {
        const string ruleJsonMixed =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"","resetsOnNewLevel":[]},
                {"name":"H2","match":{"placeholderTypes":["body","obj"],"paragraphLevel":0},"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "1. Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Agenda"))),
            new TestSlide(
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "2. Topic")),
                new TestShape("Body 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Summary"))));
        File.WriteAllText(rulePath, ruleJsonMixed);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(
                texts,
                Is.EqualTo(new[]
                {
                    "Intro",
                    "1.1 Agenda",
                    "Topic",
                    "2.1 Summary",
                }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_050__Execute_IsIdempotent_AfterPrefixRemoval()
    {
        // format:"" で prefix 削除後のファイルに再度同じルールを適用しても結果が変化しない（TP-050 冪等性、prefix 削除版）
        const string ruleJsonForRemoval =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"","resetsOnNewLevel":[]},
                {"name":"H2","match":{"placeholderTypes":["body","obj"],"paragraphLevel":0},"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var output1Path = Path.Combine(tempDir, "output1.pptx");
        var output2Path = Path.Combine(tempDir, "output2.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "1. Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Background"))),
            new TestSlide(
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "2. Topic")),
                new TestShape("Body 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Summary"))));
        File.WriteAllText(rulePath, ruleJsonForRemoval);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, output1Path, rulePath);
            command.Execute(output1Path, output2Path, rulePath);

            var output1Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output1Path);
            var output2Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output2Path);

            // H1 は prefix 削除、H2 は H1 カウンタを維持して採番される
            Assert.Multiple(() =>
            {
                Assert.That(
                    output1Texts,
                    Is.EqualTo(new[] { "Intro", "1.1 Background", "Topic", "2.1 Summary" }));
                Assert.That(output2Texts, Is.EqualTo(output1Texts)); // 2 回目適用後も同一
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_IsIdempotent_WhenFormatIsEmpty()
    {
        const string ruleJsonForRemoval =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"","resetsOnNewLevel":[]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var output1Path = Path.Combine(tempDir, "output1.pptx");
        var output2Path = Path.Combine(tempDir, "output2.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "1. Intro")),
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Topic"))));
        File.WriteAllText(rulePath, ruleJsonForRemoval);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, output1Path, rulePath);
            command.Execute(output1Path, output2Path, rulePath);

            var output1Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output1Path);
            var output2Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output2Path);

            Assert.Multiple(() =>
            {
                Assert.That(output1Texts, Is.EqualTo(new[] { "Intro", "Topic" }));
                Assert.That(output2Texts, Is.EqualTo(output1Texts));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_ThrowsWhenInputAndOutputAreSame_AndWritesTrace()
    {
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");
        File.WriteAllText(inputPath, "dummy");
        File.WriteAllText(rulePath, RuleJson);

        var traceWriter = new StringWriter();
        var listener = new TextWriterTraceListener(traceWriter);
        Trace.Listeners.Add(listener);
        Trace.AutoFlush = true;

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            Assert.Throws<ArgumentException>(() => command.Execute(inputPath, inputPath, rulePath));

            listener.Flush();
            Assert.That(traceWriter.ToString(), Does.Contain("ArgumentException"));
        }
        finally
        {
            Trace.Listeners.Remove(listener);
            listener.Dispose();
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_020__Execute_AppliesFullWidthSeparator()
    {
        // separator が全角スペースのとき、番号と本文テキストの間に全角スペースが挿入される（TP-020）
        // prefixRegex は数字プレフィックスのみにマッチさせ、本文テキストが消えないようにする
        const string ruleJsonFullWidth =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":"\u3000",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "はじめに"))));
        File.WriteAllText(rulePath, ruleJsonFullWidth);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(texts[0], Is.EqualTo("1.\u3000はじめに"));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_020__Execute_NumbersObjPlaceholderAndSkipsFreeShape()
    {
        // obj プレースホルダーは番号付与され、自由配置図形はスキップされる（TP-020）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro")),
                new TestShape("Object 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object,
                    new TestParagraph(0, "Obj Item")),
                new TestShape("Free Text", null,
                    new TestParagraph(0, "Free"))));
        File.WriteAllText(rulePath, RuleJson);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.Multiple(() =>
            {
                Assert.That(texts[0], Is.EqualTo("1. Intro"));    // title → H1
                Assert.That(texts[1], Is.EqualTo("1.1 Obj Item")); // obj lv0 → H2
                Assert.That(texts[2], Is.EqualTo("Free"));          // 自由配置図形 → 変更なし
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_AppliesNumberingByShapeNameWhenPlaceholderTypeIsMissing()
    {
        const string ruleJsonByShapeName =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]},
                {"name":"H2","match":{"shapeNames":["コンテンツ プレースホルダー 2"],"paragraphLevel":0},"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("タイトル 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro")),
                new TestShape("コンテンツ プレースホルダー 2", null,
                    new TestParagraph(0, "Profile")),
                new TestShape("Free Text", null,
                    new TestParagraph(0, "Free"))));
        File.WriteAllText(rulePath, ruleJsonByShapeName);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.Multiple(() =>
            {
                Assert.That(texts[0], Is.EqualTo("1. Intro"));
                Assert.That(texts[1], Is.EqualTo("1.1 Profile"));
                Assert.That(texts[2], Is.EqualTo("Free"));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_SupportsOrConditionsWithinSameLevelUsingMatches()
    {
        const string ruleJsonWithMatches =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]},
                {"name":"H2","matches":[{"placeholderTypes":["body"],"paragraphLevel":0},{"shapeNames":["コンテンツ プレースホルダー 2"],"paragraphLevel":0}],"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Body Item")),
                new TestShape("コンテンツ プレースホルダー 2", null,
                    new TestParagraph(0, "Name Item"))));
        File.WriteAllText(rulePath, ruleJsonWithMatches);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(
                texts,
                Is.EqualTo(new[]
                {
                    "1. Intro",
                    "1.1 Body Item",
                    "1.2 Name Item",
                }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void Execute_SkipsExcludedSlideRangesWithoutAdvancingCounters()
    {
        const string ruleJsonWithExcludedSlides =
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "excludedSlideRanges":[
                {"startSlideNumber":1,"endSlideNumber":2},
                {"startSlideNumber":5,"endSlideNumber":5}
              ],
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]},
                {"name":"H2","match":{"placeholderTypes":["body","obj"],"paragraphLevel":0},"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Cover")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Cover body"))),
            new TestSlide(
                new TestShape("Title 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Table of Contents")),
                new TestShape("Body 2", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Toc body"))),
            new TestSlide(
                new TestShape("Title 3", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Chapter 1")),
                new TestShape("Body 3", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Item 1"))),
            new TestSlide(
                new TestShape("Title 4", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Chapter 2")),
                new TestShape("Body 4", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Item 2"))),
            new TestSlide(
                new TestShape("Title 5", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Appendix")),
                new TestShape("Body 5", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Appendix item"))));
        File.WriteAllText(rulePath, ruleJsonWithExcludedSlides);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            Assert.That(
                texts,
                Is.EqualTo(new[]
                {
                    "Cover",
                    "Cover body",
                    "Table of Contents",
                    "Toc body",
                    "1. Chapter 1",
                    "1.1 Item 1",
                    "2. Chapter 2",
                    "2.1 Item 2",
                    "Appendix",
                    "Appendix item",
                }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_060__Execute_ThrowsForNonexistentRuleFile()
    {
        // 存在しないルールファイルを指定したとき例外が発生する（TP-060）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("T", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "X"))));

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            Assert.Catch<Exception>(() =>
                command.Execute(inputPath, Path.Combine(tempDir, "out.pptx"), "nonexistent-rule.json"));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_100__Execute_SwitchesRuleCorrectly()
    {
        // ルールA で付与した番号を prefixRegex が除去し、ルールB の番号体系に差し替えられる（TP-100）
        const string ruleJsonB =
            """
            {
              "prefixRegex":"^[^\\s\\u3000]+(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"\u7b2c{H1}\u7ae0","resetsOnNewLevel":[]},
                {"name":"H2","match":{"placeholderTypes":["body","obj"],"paragraphLevel":0},"format":"{H1}-{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """;

        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var output1Path = Path.Combine(tempDir, "output1.pptx");
        var output2Path = Path.Combine(tempDir, "output2.pptx");
        var rulePathA = Path.Combine(tempDir, "ruleA.json");
        var rulePathB = Path.Combine(tempDir, "ruleB.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Background"))));
        File.WriteAllText(rulePathA, RuleJson);
        File.WriteAllText(rulePathB, ruleJsonB);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());

            // ステップ1: ルールA を適用
            command.Execute(inputPath, output1Path, rulePathA);
            var output1Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output1Path);
            Assert.That(output1Texts, Is.EqualTo(new[] { "1. Intro", "1.1 Background" }));

            // ステップ2: ルールA の出力を入力として、ルールB を適用
            command.Execute(output1Path, output2Path, rulePathB);
            var output2Texts = PptxTestDocumentFactory.ReadAllParagraphTexts(output2Path);
            Assert.That(output2Texts, Is.EqualTo(new[] { "第1章 Intro", "1-1 Background" }));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_100__Execute_ApplyResultVisibleInInspect()
    {
        // apply 後の output.pptx を inspect すると、番号付与後のテキストが TSV に反映される（TP-100）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var outputPath = Path.Combine(tempDir, "output.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro"))));
        File.WriteAllText(rulePath, RuleJson);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var inspectOutput = new StringWriter();
            new InspectCommand(new SlideWalker(), inspectOutput).Execute(outputPath);

            var lines = inspectOutput.ToString()
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            // TSV の Text 列に番号付与後のテキストが反映されていること
            Assert.That(lines[1], Does.Contain("1. Intro"));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_090__Execute_HandlesLargeNumberOfSlidesAndEmptySlides()
    {
        // 50枚超のスライドと空スライドが混在しても apply が完走し、連番が継続する（TP-090）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input-large.pptx");
        var outputPath = Path.Combine(tempDir, "output-large.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        var slides = Enumerable.Range(0, 55)
            .Select(i =>
                i % 11 == 5
                    ? new TestSlide()
                    : new TestSlide(
                        new TestShape($"Title {i}", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                            new TestParagraph(0, $"Heading {i}")),
                        new TestShape($"Body {i}", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                            new TestParagraph(0, $"Item {i}"))))
            .ToArray();

        PptxTestDocumentFactory.Create(inputPath, slides);
        File.WriteAllText(rulePath, RuleJson);

        try
        {
            new ApplyCommand(new SlideWalker(), new PrefixReplacer()).Execute(inputPath, outputPath, rulePath);

            var texts = PptxTestDocumentFactory.ReadAllParagraphTexts(outputPath);
            var nonEmptySlideCount = slides.Count(slide => slide.Shapes.Length > 0);

            Assert.Multiple(() =>
            {
                Assert.That(texts, Has.Count.EqualTo(nonEmptySlideCount * 2));
                Assert.That(texts[0], Is.EqualTo("1. Heading 0"));
                Assert.That(texts[1], Is.EqualTo("1.1 Item 0"));
                Assert.That(texts[^2], Is.EqualTo($"{nonEmptySlideCount}. Heading 54"));
                Assert.That(texts[^1], Is.EqualTo($"{nonEmptySlideCount}.1 Item 54"));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_100__Execute_RemainsStableAcrossThreeRuns()
    {
        // 同一ルールで 3 回連続実行しても結果が変化しない（TP-100）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var inputPath = Path.Combine(tempDir, "input.pptx");
        var output1Path = Path.Combine(tempDir, "output1.pptx");
        var output2Path = Path.Combine(tempDir, "output2.pptx");
        var output3Path = Path.Combine(tempDir, "output3.pptx");
        var rulePath = Path.Combine(tempDir, "rule.json");

        PptxTestDocumentFactory.Create(
            inputPath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                    new TestParagraph(0, "Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, "Background"),
                    new TestParagraph(1, "Detail"))));
        File.WriteAllText(rulePath, RuleJson);

        try
        {
            var command = new ApplyCommand(new SlideWalker(), new PrefixReplacer());
            command.Execute(inputPath, output1Path, rulePath);
            command.Execute(output1Path, output2Path, rulePath);
            command.Execute(output2Path, output3Path, rulePath);

            var texts1 = PptxTestDocumentFactory.ReadAllParagraphTexts(output1Path);
            var texts2 = PptxTestDocumentFactory.ReadAllParagraphTexts(output2Path);
            var texts3 = PptxTestDocumentFactory.ReadAllParagraphTexts(output3Path);

            Assert.Multiple(() =>
            {
                Assert.That(texts2, Is.EqualTo(texts1));
                Assert.That(texts3, Is.EqualTo(texts1));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }
}
