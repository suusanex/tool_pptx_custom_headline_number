using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;

namespace PptxHeadlineNumbering;

public sealed class ApplyCommand
{
    private readonly SlideWalker _slideWalker;
    private readonly PrefixReplacer _prefixReplacer;

    public ApplyCommand(SlideWalker slideWalker, PrefixReplacer prefixReplacer)
    {
        _slideWalker = slideWalker ?? throw new ArgumentNullException(nameof(slideWalker));
        _prefixReplacer = prefixReplacer ?? throw new ArgumentNullException(nameof(prefixReplacer));
    }

    public void Execute(string inputPath, string outputPath, string rulePath)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
            ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
            ArgumentException.ThrowIfNullOrWhiteSpace(rulePath);

            var inputFullPath = Path.GetFullPath(inputPath);
            var outputFullPath = Path.GetFullPath(outputPath);
            if (string.Equals(inputFullPath, outputFullPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Input and output paths must be different.");
            }

            var rule = NumberingRule.LoadFromFile(rulePath);
            var prefixRegex = rule.BuildPrefixRegex();
            var counter = new HeadingCounter(rule.Levels);

            using var sourceStream = File.OpenRead(inputFullPath);
            using var editableStream = new MemoryStream();
            sourceStream.CopyTo(editableStream);
            editableStream.Position = 0;

            using (var presentationDocument = PresentationDocument.Open(editableStream, true))
            {
                var paragraphs = _slideWalker.Walk(presentationDocument);
                foreach (var paragraph in paragraphs)
                {
                    if (rule.IsExcludedSlide(paragraph.SlideIndex))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(paragraph.CurrentText))
                    {
                        continue;
                    }

                    var level = rule.MatchLevel(paragraph.PlaceholderType, paragraph.ShapeName, paragraph.ParagraphLevel);
                    if (level is null)
                    {
                        continue;
                    }

                    counter.Increment(level.Name);
                    var format = level.Format
                        ?? throw new InvalidDataException(
                            string.Format(
                                System.Globalization.CultureInfo.InvariantCulture,
                                "format is required for level: {0}",
                                level.Name));
                    var formattedPrefix = format.Length == 0
                        ? string.Empty
                        : counter.Format(format);
                    _prefixReplacer.Replace(
                        paragraph.ParagraphElement,
                        prefixRegex,
                        formattedPrefix,
                        rule.Separator,
                        rule.InsertWhenPrefixMissing);
                }

                presentationDocument.PresentationPart?.Presentation?.Save();
            }

            editableStream.Position = 0;
            using var outputStream = File.Create(outputFullPath);
            editableStream.CopyTo(outputStream);
        }
        catch (Exception ex)
        {
            Trace.WriteLine(ex.ToString());
            throw;
        }
    }
}
