using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxHeadlineNumbering;

public sealed class NumberingRule
{
    public string PrefixRegex { get; init; } = "^[^\\s\\u3000]+(?:[\\s\\u3000]+)?";

    public string Separator { get; init; } = " ";

    public bool InsertWhenPrefixMissing { get; init; } = true;

    public IReadOnlyList<ExcludedSlideRange> ExcludedSlideRanges { get; init; } = Array.Empty<ExcludedSlideRange>();

    public IReadOnlyList<NumberingLevelRule> Levels { get; init; } = Array.Empty<NumberingLevelRule>();

    public static NumberingRule LoadFromFile(string path)
    {
        try
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);

            var fullPath = Path.GetFullPath(path);
            var json = File.ReadAllText(fullPath);
            var rule = JsonSerializer.Deserialize(json, NumberingRuleJsonContext.Default.NumberingRule)
                ?? throw new JsonException("Rule file is empty.");

            rule.Validate();
            return rule;
        }
        catch (Exception ex)
        {
            Trace.WriteLine(ex.ToString());
            throw;
        }
    }

    public Regex BuildPrefixRegex()
    {
        try
        {
            return new Regex(PrefixRegex, RegexOptions.CultureInvariant);
        }
        catch (Exception ex)
        {
            Trace.WriteLine(ex.ToString());
            throw;
        }
    }

    public NumberingLevelRule? MatchLevel(P.PlaceholderValues? placeholderType, string shapeName, int paragraphLevel)
    {
        var placeholderToken = ToPlaceholderToken(placeholderType);
        return Levels.FirstOrDefault(level => level.IsMatch(placeholderToken, shapeName, paragraphLevel));
    }

    public bool IsExcludedSlide(int slideIndex)
    {
        var slideNumber = slideIndex + 1;
        return (ExcludedSlideRanges ?? Array.Empty<ExcludedSlideRange>()).Any(range => range.Contains(slideNumber));
    }

    private void Validate()
    {
        if (Levels.Count == 0)
        {
            throw new InvalidDataException("At least one level must be defined.");
        }

        if (string.IsNullOrWhiteSpace(PrefixRegex))
        {
            throw new InvalidDataException("prefixRegex is required.");
        }

        _ = BuildPrefixRegex();
        foreach (var range in ExcludedSlideRanges ?? Array.Empty<ExcludedSlideRange>())
        {
            range.Validate();
        }

        var knownLevels = new HashSet<string>(StringComparer.Ordinal);
        foreach (var level in Levels)
        {
            if (string.IsNullOrWhiteSpace(level.Name))
            {
                throw new InvalidDataException("level.name is required.");
            }

            if (!knownLevels.Add(level.Name))
            {
                throw new InvalidDataException(
                    string.Format(CultureInfo.InvariantCulture, "Duplicate level name: {0}", level.Name));
            }

            if (level.Format is null)
            {
                throw new InvalidDataException(
                    string.Format(CultureInfo.InvariantCulture, "format is required for level: {0}", level.Name));
            }

            if (level.Format.Length > 0 && string.IsNullOrWhiteSpace(level.Format))
            {
                throw new InvalidDataException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "format must be empty or non-whitespace for level: {0}",
                        level.Name));
            }

            level.Validate();
        }

        foreach (var level in Levels)
        {
            foreach (var dependency in level.ResetsOnNewLevel)
            {
                if (!knownLevels.Contains(dependency))
                {
                    throw new InvalidDataException(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "Unknown resetsOnNewLevel reference '{0}' in level '{1}'.",
                            dependency,
                            level.Name));
                }
            }
        }
    }

    internal static string? ToPlaceholderToken(P.PlaceholderValues? placeholderType)
    {
        if (placeholderType is null)
        {
            return null;
        }

        if (placeholderType.Value == P.PlaceholderValues.Title)
        {
            return "title";
        }

        if (placeholderType.Value == P.PlaceholderValues.CenteredTitle)
        {
            return "ctrTitle";
        }

        if (placeholderType.Value == P.PlaceholderValues.Body)
        {
            return "body";
        }

        if (placeholderType.Value == P.PlaceholderValues.Object)
        {
            return "obj";
        }

        return placeholderType.Value.ToString();
    }
}

public sealed class NumberingLevelRule
{
    public string Name { get; init; } = string.Empty;

    public NumberingMatchRule? Match { get; init; }

    public IReadOnlyList<NumberingMatchRule> Matches { get; init; } = Array.Empty<NumberingMatchRule>();

    public string? Format { get; init; }

    public IReadOnlyList<string> ResetsOnNewLevel { get; init; } = Array.Empty<string>();

    public bool IsMatch(string? placeholderType, string shapeName, int paragraphLevel)
    {
        return GetEffectiveMatches().Any(match => match.IsMatch(placeholderType, shapeName, paragraphLevel));
    }

    public void Validate()
    {
        var effectiveMatches = GetEffectiveMatches().ToArray();
        if (effectiveMatches.Length == 0)
        {
            throw new InvalidDataException($"match or matches is required for level: {Name}.");
        }

        foreach (var match in effectiveMatches)
        {
            match.Validate(Name);
        }
    }

    private IEnumerable<NumberingMatchRule> GetEffectiveMatches()
    {
        if (Match is not null)
        {
            yield return Match;
        }

        foreach (var match in Matches ?? Array.Empty<NumberingMatchRule>())
        {
            if (match is not null)
            {
                yield return match;
            }
        }
    }
}

public sealed class ExcludedSlideRange
{
    public int StartSlideNumber { get; init; }

    public int EndSlideNumber { get; init; }

    public bool Contains(int slideNumber)
    {
        return slideNumber >= StartSlideNumber && slideNumber <= EndSlideNumber;
    }

    public void Validate()
    {
        if (StartSlideNumber < 1)
        {
            throw new InvalidDataException("startSlideNumber must be greater than or equal to 1.");
        }

        if (EndSlideNumber < StartSlideNumber)
        {
            throw new InvalidDataException("endSlideNumber must be greater than or equal to startSlideNumber.");
        }
    }
}

public sealed class NumberingMatchRule
{
    public IReadOnlyList<string> PlaceholderTypes { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> ShapeNames { get; init; } = Array.Empty<string>();

    public int? ParagraphLevel { get; init; }

    public bool IsMatch(string? placeholderType, string shapeName, int paragraphLevel)
    {
        var placeholderTypes = PlaceholderTypes ?? Array.Empty<string>();
        var shapeNames = ShapeNames ?? Array.Empty<string>();

        if (placeholderTypes.Count > 0)
        {
            if (string.IsNullOrWhiteSpace(placeholderType))
            {
                return false;
            }

            if (!placeholderTypes.Any(expected => string.Equals(expected, placeholderType, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }
        }

        if (shapeNames.Count > 0
            && !shapeNames.Any(expected => string.Equals(expected, shapeName, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        return !ParagraphLevel.HasValue || ParagraphLevel.Value == paragraphLevel;
    }

    public void Validate(string levelName)
    {
        if ((PlaceholderTypes?.Count ?? 0) == 0 && (ShapeNames?.Count ?? 0) == 0)
        {
            throw new InvalidDataException($"placeholderTypes or shapeNames is required for level: {levelName}.");
        }
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(NumberingRule))]
[JsonSerializable(typeof(NumberingLevelRule))]
[JsonSerializable(typeof(ExcludedSlideRange))]
[JsonSerializable(typeof(NumberingMatchRule))]
internal partial class NumberingRuleJsonContext : JsonSerializerContext
{
}
