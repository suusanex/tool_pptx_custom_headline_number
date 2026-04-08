using System.IO;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxHeadlineNumbering;

public sealed class SlideWalker
{
    public IReadOnlyList<ParagraphInfo> Walk(PresentationDocument presentationDocument)
    {
        ArgumentNullException.ThrowIfNull(presentationDocument);

        var presentationPart = presentationDocument.PresentationPart
            ?? throw new InvalidDataException("PresentationPart is missing.");
        var slideIdList = presentationPart.Presentation?.SlideIdList
            ?? throw new InvalidDataException("SlideIdList is missing.");

        var result = new List<ParagraphInfo>();
        var slideIndex = 0;
        foreach (var slideId in slideIdList.Elements<P.SlideId>())
        {
            var relationshipId = slideId.RelationshipId?.Value
                ?? throw new InvalidDataException("Slide relationship ID is missing.");
            var slidePart = presentationPart.GetPartById(relationshipId) as SlidePart
                ?? throw new InvalidDataException($"Slide part was not found: {relationshipId}.");
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree is null)
            {
                slideIndex++;
                continue;
            }

            foreach (var shape in shapeTree.Elements<P.Shape>())
            {
                var shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
                var placeholderType = shape
                    .NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>()?
                    .Type?
                    .Value;
                var textBody = shape.TextBody;
                if (textBody is null)
                {
                    continue;
                }

                foreach (var paragraph in textBody.Elements<A.Paragraph>())
                {
                    var level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
                    var text = string.Concat(paragraph.Descendants<A.Text>().Select(x => x.Text));
                    result.Add(new ParagraphInfo(slideIndex, shapeName, placeholderType, level, text, paragraph));
                }
            }

            slideIndex++;
        }

        return result;
    }
}
