using System.Text.RegularExpressions;
using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Search;

/// <summary>
/// Стратегия поиска на основе регулярных выражений
/// </summary>
public class RegexSearchStrategy : ITextSearchStrategy
{
    private readonly List<RegexPattern> _patterns;

    public string StrategyName { get; }

    public RegexSearchStrategy(string name, params RegexPattern[] patterns)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _patterns = patterns?.ToList() ?? throw new ArgumentNullException(nameof(patterns));
            
        if (_patterns.Count == 0)
            throw new ArgumentException("Необходимо указать хотя бы один паттерн", nameof(patterns));
    }

    public IEnumerable<TextMatch> FindMatches(string text)
    {
        if (string.IsNullOrEmpty(text))
            yield break;

        var foundMatches = new HashSet<string>();

        foreach (var pattern in _patterns)
        {
            var matches = Regex.Matches(text, pattern.Pattern, pattern.Options);

            foreach (Match match in matches)
            {
                if (!match.Success || foundMatches.Contains(match.Value))
                    continue;

                foundMatches.Add(match.Value);

                yield return new TextMatch
                {
                    Value = match.Value,
                    StartIndex = match.Index,
                    Length = match.Length,
                    MatchType = pattern.Name,
                    Metadata = new Dictionary<string, object>
                    {
                        ["Pattern"] = pattern.Pattern,
                        ["PatternName"] = pattern.Name
                    }
                };
            }
        }
    }
}