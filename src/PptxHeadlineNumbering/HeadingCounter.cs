using System.Globalization;
using System.Text.RegularExpressions;

namespace PptxHeadlineNumbering;

public sealed class HeadingCounter
{
    private readonly Dictionary<string, int> _counts;
    private readonly Dictionary<string, List<string>> _resetTargetsByTrigger;

    public HeadingCounter(IReadOnlyList<NumberingLevelRule> levels)
    {
        ArgumentNullException.ThrowIfNull(levels);
        if (levels.Count == 0)
        {
            throw new ArgumentException("At least one level must be defined.", nameof(levels));
        }

        _counts = levels.ToDictionary(level => level.Name, _ => 0, StringComparer.Ordinal);
        _resetTargetsByTrigger = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (var level in levels)
        {
            foreach (var resetTrigger in level.ResetsOnNewLevel)
            {
                if (!_counts.ContainsKey(resetTrigger))
                {
                    throw new KeyNotFoundException($"Unknown reset trigger level: {resetTrigger}");
                }

                if (!_resetTargetsByTrigger.TryGetValue(resetTrigger, out var targets))
                {
                    targets = new List<string>();
                    _resetTargetsByTrigger.Add(resetTrigger, targets);
                }

                if (!targets.Contains(level.Name, StringComparer.Ordinal))
                {
                    targets.Add(level.Name);
                }
            }
        }
    }

    public void Increment(string levelName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(levelName);
        if (!_counts.ContainsKey(levelName))
        {
            throw new KeyNotFoundException($"Unknown level: {levelName}");
        }

        _counts[levelName]++;
        if (!_resetTargetsByTrigger.TryGetValue(levelName, out var resetTargets))
        {
            return;
        }

        foreach (var resetTarget in resetTargets.Where(target => !string.Equals(target, levelName, StringComparison.Ordinal)))
        {
            _counts[resetTarget] = 0;
        }
    }

    public string Format(string template)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(template);

        return Regex.Replace(
            template,
            "\\{([A-Za-z0-9_]+)\\}",
            match =>
            {
                var counterName = match.Groups[1].Value;
                if (!_counts.TryGetValue(counterName, out var value))
                {
                    throw new KeyNotFoundException($"Unknown counter in template: {counterName}");
                }

                return value.ToString(CultureInfo.InvariantCulture);
            });
    }

    public int GetCount(string levelName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(levelName);
        return _counts.TryGetValue(levelName, out var value)
            ? value
            : throw new KeyNotFoundException($"Unknown level: {levelName}");
    }
}
