using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using PptxHeadlineNumbering.Tests.TestData;

namespace PptxHeadlineNumbering.Tests;

public class SlideWalkerTests
{
    [Test]
    public void Walk_ReturnsMetadataForAllParagraphKinds()
    {
        var filePath = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.pptx");
        PptxTestDocumentFactory.Create(
            filePath,
            new TestSlide(
                new TestShape("Title 1", PlaceholderValues.Title, new TestParagraph(0, "Intro")),
                new TestShape("Body 1", PlaceholderValues.Body, new TestParagraph(0, "Background"), new TestParagraph(1, "Detail A")),
                new TestShape("Object 1", PlaceholderValues.Object, new TestParagraph(0, "Obj 0"), new TestParagraph(1, "Obj 1")),
                new TestShape("Free Text", null, new TestParagraph(2, "No numbering target"))),
            new TestSlide(
                new TestShape("Centered", PlaceholderValues.CenteredTitle, new TestParagraph(0, "Cover"))));

        try
        {
            using var document = PresentationDocument.Open(filePath, false);
            var paragraphs = new SlideWalker().Walk(document).ToList();

            Assert.That(paragraphs, Has.Count.EqualTo(7));
            Assert.Multiple(() =>
            {
                Assert.That(paragraphs[0].SlideIndex, Is.EqualTo(0));
                Assert.That(paragraphs[0].ShapeName, Is.EqualTo("Title 1"));
                Assert.That(paragraphs[0].PlaceholderType, Is.EqualTo(PlaceholderValues.Title));
                Assert.That(paragraphs[0].ParagraphLevel, Is.EqualTo(0));

                Assert.That(paragraphs[2].PlaceholderType, Is.EqualTo(PlaceholderValues.Body));
                Assert.That(paragraphs[2].ParagraphLevel, Is.EqualTo(1));

                Assert.That(paragraphs[5].PlaceholderType, Is.Null);
                Assert.That(paragraphs[5].ParagraphLevel, Is.EqualTo(2));

                Assert.That(paragraphs[6].SlideIndex, Is.EqualTo(1));
                Assert.That(paragraphs[6].PlaceholderType, Is.EqualTo(PlaceholderValues.CenteredTitle));
            });
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Test]
    public void UT_IT_090__Walk_HandlesLargeNumberOfSlides()
    {
        // 50枚以上のスライドを含む .pptx を正常処理できること（TP-090）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "large.pptx");

        const int slideCount = 55;
        const int paragraphsPerSlide = 6; // title 1 + body 5

        var slides = Enumerable.Range(0, slideCount)
            .Select(i => new TestSlide(
                new TestShape($"Title {i}", PlaceholderValues.Title,
                    new TestParagraph(0, $"Heading {i}")),
                new TestShape($"Body {i}", PlaceholderValues.Body,
                    Enumerable.Range(0, 5)
                        .Select(j => new TestParagraph(0, $"Item {i}-{j}"))
                        .ToArray())))
            .ToArray();

        PptxTestDocumentFactory.Create(filePath, slides);

        try
        {
            using var document = PresentationDocument.Open(filePath, false);
            var paragraphs = new SlideWalker().Walk(document);

            Assert.Multiple(() =>
            {
                Assert.That(paragraphs.Count, Is.EqualTo(slideCount * paragraphsPerSlide));
                Assert.That(paragraphs[0].SlideIndex, Is.EqualTo(0));
                Assert.That(paragraphs[^1].SlideIndex, Is.EqualTo(slideCount - 1));
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_090__Walk_HandlesEmptySlidesMixedIn()
    {
        // 空スライド（図形なし）が混在しても正常にスキップし、SlideIndex が継続する（TP-090）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "empty-slides.pptx");

        PptxTestDocumentFactory.Create(
            filePath,
            new TestSlide(
                new TestShape("Title 0", PlaceholderValues.Title, new TestParagraph(0, "Heading 0"))),
            new TestSlide(), // 空スライド（図形なし）
            new TestSlide(
                new TestShape("Title 2", PlaceholderValues.Title, new TestParagraph(0, "Heading 2"))));

        try
        {
            using var document = PresentationDocument.Open(filePath, false);
            var paragraphs = new SlideWalker().Walk(document).ToList();

            Assert.Multiple(() =>
            {
                Assert.That(paragraphs, Has.Count.EqualTo(2));
                Assert.That(paragraphs[0].SlideIndex, Is.EqualTo(0));
                Assert.That(paragraphs[1].SlideIndex, Is.EqualTo(2)); // 空スライドで index が継続
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }
}
