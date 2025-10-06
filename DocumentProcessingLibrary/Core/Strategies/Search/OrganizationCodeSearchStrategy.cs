using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Search;

/// <summary>
/// Стратегия поиска кодов организаций на основе предоставленного списка
/// </summary>
public class OrganizationCodeSearchStrategy : ITextSearchStrategy
{
    private readonly HashSet<string> _codes;
    public string StrategyName => "OrganizationCodes";
    public OrganizationCodeSearchStrategy(IEnumerable<string> codes)
    {
        _codes = new HashSet<string>(codes ?? Enumerable.Empty<string>());
    }
    public IEnumerable<TextMatch> FindMatches(string text)
    {
        if (string.IsNullOrEmpty(text) || _codes.Count == 0)
            yield break;
        foreach (var code in _codes)
        {
            if (string.IsNullOrEmpty(code))
                continue;
            var startIndex = 0;
            while ((startIndex = text.IndexOf(code, startIndex)) != -1)
            {
                var isValidMatch = true;
                if (startIndex > 0)
                {
                    var prevChar = text[startIndex - 1];
                    if (char.IsLetterOrDigit(prevChar))
                    {
                        isValidMatch = false;
                    }
                }
                if (isValidMatch && startIndex + code.Length < text.Length)
                {
                    var nextChar = text[startIndex + code.Length];
                    if (char.IsLetterOrDigit(nextChar) && nextChar != '.')
                    {
                        isValidMatch = false;
                    }
                }
                if (isValidMatch)
                {
                    yield return new TextMatch
                    {
                        Value = code, StartIndex = startIndex, Length = code.Length, MatchType = "OrganizationCode", Metadata = new Dictionary<string, object>
                        {
                            ["Code"] = code, ["IsStandaloneCode"] = true
                        }
                    };
                }
                startIndex += code.Length;
            }
        }
    }
    /// <summary>
    /// Добавляет новые коды в стратегию поиска
    /// </summary>
    public void AddCodes(IEnumerable<string> codes)
    {
        foreach (var code in codes)
        {
            if (!string.IsNullOrEmpty(code))
            {
                _codes.Add(code);
            }
        }
    }
}