namespace DocumentProcessingLibrary.Core.Models;

/// <summary>
/// Результат поиска с метаданными
/// </summary>
public class TextMatch
{
    public string Value { get; set; } = string.Empty;
    public int StartIndex { get; set; }
    public int Length { get; set; }
    public string MatchType { get; set; } = string.Empty;
    public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
}