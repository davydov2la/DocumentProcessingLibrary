using System.Text.RegularExpressions;

namespace DocumentProcessingLibrary.Core.Strategies.Search;

/// <summary>
/// Паттерн регулярного выражения с метаданными
/// </summary>
public class RegexPattern
{
    public string Name { get; set; }
    public string Pattern { get; set; }
    public RegexOptions Options { get; set; } = RegexOptions.None;

    public RegexPattern(string name, string pattern, RegexOptions options = RegexOptions.None)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Pattern = pattern ?? throw new ArgumentNullException(nameof(pattern));
        Options = options;
    }
}