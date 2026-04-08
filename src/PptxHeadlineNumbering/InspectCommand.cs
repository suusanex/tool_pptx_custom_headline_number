using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;

namespace PptxHeadlineNumbering;

public sealed class InspectCommand
{
    private readonly SlideWalker _slideWalker;
    private readonly TextWriter _output;

    public InspectCommand(SlideWalker slideWalker, TextWriter output)
    {
        _slideWalker = slideWalker ?? throw new ArgumentNullException(nameof(slideWalker));
        _output = output ?? throw new ArgumentNullException(nameof(output));
    }

    public void Execute(string inputPath)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
            var fullPath = Path.GetFullPath(inputPath);

            using var presentationDocument = PresentationDocument.Open(fullPath, false);
            var paragraphs = _slideWalker.Walk(presentationDocument);

            _output.WriteLine("SlideIndex\tShapeName\tPlaceholderType\tLevel\tText");
            foreach (var paragraph in paragraphs)
            {
                _output.WriteLine(
                    $"{paragraph.SlideIndex}\t{Sanitize(paragraph.ShapeName)}\t{FormatPlaceholderType(paragraph.PlaceholderType)}\t{paragraph.ParagraphLevel}\t{Sanitize(paragraph.CurrentText)}");
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine(ex.ToString());
            throw;
        }
    }

    private static string FormatPlaceholderType(DocumentFormat.OpenXml.Presentation.PlaceholderValues? placeholderType)
    {
        if (placeholderType is null)
        {
            return string.Empty;
        }

        if (placeholderType.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title)
        {
            return "title";
        }

        if (placeholderType.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle)
        {
            return "ctrTitle";
        }

        if (placeholderType.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body)
        {
            return "body";
        }

        if (placeholderType.Value == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object)
        {
            return "obj";
        }

        return placeholderType.Value.ToString();
    }

    private static string Sanitize(string value)
    {
        return (value ?? string.Empty)
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ');
    }
}
