using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxHeadlineNumbering.Tests.TestData;

internal sealed record TestParagraph(int Level, params string[] Runs);

internal sealed record TestShape(string Name, P.PlaceholderValues? PlaceholderType, params TestParagraph[] Paragraphs);

internal sealed record TestSlide(params TestShape[] Shapes);

internal static class PptxTestDocumentFactory
{
    public static void Create(string path, params TestSlide[] slides)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = File.Create(path);
        Create(stream, slides);
    }

    public static void Create(Stream stream, IReadOnlyList<TestSlide> slides)
    {
        using var document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        var slideIdList = new P.SlideIdList();
        presentationPart.Presentation = new P.Presentation(
            slideIdList,
            new P.SlideSize { Cx = 9144000, Cy = 6858000, Type = P.SlideSizeValues.Screen4x3 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 });

        uint slideIdValue = 256;
        foreach (var slide in slides)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(
                    new P.ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.GroupShapeProperties(
                            new A.TransformGroup(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 0L, Cy = 0L },
                                new A.ChildOffset { X = 0L, Y = 0L },
                                new A.ChildExtents { Cx = 0L, Cy = 0L })))),
                new P.ColorMapOverride(new A.MasterColorMapping()));

            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("ShapeTree was not created.");

            uint shapeId = 2;
            foreach (var shape in slide.Shapes)
            {
                var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();
                if (shape.PlaceholderType is not null)
                {
                    applicationNonVisualDrawingProperties.Append(new P.PlaceholderShape { Type = shape.PlaceholderType.Value });
                }

                var textBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
                foreach (var paragraph in shape.Paragraphs)
                {
                    var openXmlParagraph = new A.Paragraph();
                    if (paragraph.Level > 0)
                    {
                        openXmlParagraph.ParagraphProperties = new A.ParagraphProperties { Level = paragraph.Level };
                    }

                    foreach (var runText in paragraph.Runs)
                    {
                        var run = new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text(runText));
                        openXmlParagraph.Append(run);
                    }

                    textBody.Append(openXmlParagraph);
                }

                var openXmlShape = new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = shapeId++, Name = shape.Name },
                        new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                        applicationNonVisualDrawingProperties),
                    new P.ShapeProperties(),
                    textBody);

                shapeTree.Append(openXmlShape);
            }

            slidePart.Slide.Save();
            var relationshipId = presentationPart.GetIdOfPart(slidePart);
            slideIdList.Append(new P.SlideId { Id = slideIdValue++, RelationshipId = relationshipId });
        }

        presentationPart.Presentation.Save();
    }

    public static List<string> ReadAllParagraphTexts(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var slideWalker = new SlideWalker();
        return slideWalker.Walk(document).Select(paragraph => paragraph.CurrentText).ToList();
    }
}
