using System.Diagnostics;
using PptxHeadlineNumbering.Tests.TestData;

namespace PptxHeadlineNumbering.Tests;

public class InspectCommandTests
{
    [Test]
    public void Execute_WritesExpectedTsv()
    {
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "inspect.pptx");
        PptxTestDocumentFactory.Create(
            filePath,
            new TestSlide(
                new TestShape("Title 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title, new TestParagraph(0, "Intro")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body, new TestParagraph(0, "Background"), new TestParagraph(1, "Detail A")),
                new TestShape("Free", null, new TestParagraph(0, "Memo"))));

        try
        {
            var output = new StringWriter();
            var command = new InspectCommand(new SlideWalker(), output);
            command.Execute(filePath);

            var lines = output.ToString()
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            Assert.Multiple(() =>
            {
                Assert.That(lines[0], Is.EqualTo("SlideIndex\tShapeName\tPlaceholderType\tLevel\tText"));
                Assert.That(lines[1], Is.EqualTo("0\tTitle 1\ttitle\t0\tIntro"));
                Assert.That(lines[2], Is.EqualTo("0\tBody 1\tbody\t0\tBackground"));
                Assert.That(lines[3], Is.EqualTo("0\tBody 1\tbody\t1\tDetail A"));
                Assert.That(lines[4], Is.EqualTo("0\tFree\t\t0\tMemo"));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_010__Execute_OutputsCtrTitleAndEmptyParagraphAndHighLevel()
    {
        // ctrTitle プレースホルダー、空テキスト段落、level>=2 の段落が正しく TSV 出力される
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "inspect2.pptx");
        PptxTestDocumentFactory.Create(
            filePath,
            new TestSlide(
                new TestShape("Cover", DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle,
                    new TestParagraph(0, "Cover Title")),
                new TestShape("Body 1", DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                    new TestParagraph(0, ""),          // 空テキスト段落
                    new TestParagraph(2, "Level 2 text")))); // level >= 2

        try
        {
            var output = new StringWriter();
            new InspectCommand(new SlideWalker(), output).Execute(filePath);

            var lines = output.ToString()
                .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            Assert.Multiple(() =>
            {
                Assert.That(lines, Has.Length.EqualTo(4)); // ヘッダー + 3 段落
                Assert.That(lines[1], Is.EqualTo("0\tCover\tctrTitle\t0\tCover Title"));
                Assert.That(lines[2], Is.EqualTo("0\tBody 1\tbody\t0")); // 空テキスト（TrimEntries でトレーリングタブが除去される）
                Assert.That(lines[3], Is.EqualTo("0\tBody 1\tbody\t2\tLevel 2 text")); // level=2
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_060__Execute_TracesExceptionOnFileNotFound()
    {
        // ファイル不存在時に例外のトレースログが出力される
        var traceWriter = new StringWriter();
        var listener = new TextWriterTraceListener(traceWriter);
        Trace.Listeners.Add(listener);
        Trace.AutoFlush = true;

        try
        {
            var output = new StringWriter();
            var command = new InspectCommand(new SlideWalker(), output);
            Assert.Catch<Exception>(() => command.Execute("nonexistent-file-xyz-inspect.pptx"));

            listener.Flush();
            Assert.That(traceWriter.ToString(), Does.Contain("Exception"));
        }
        finally
        {
            Trace.Listeners.Remove(listener);
            listener.Dispose();
        }
    }
}
