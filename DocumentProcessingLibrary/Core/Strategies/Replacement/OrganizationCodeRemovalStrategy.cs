using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;
using DocumentProcessingLibrary.Core.Utilities;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия замены для десятичных обозначений с извлечением и удалением кодов организаций
/// </summary>
public class OrganizationCodeRemovalStrategy : ITextReplacementStrategy
{
    private readonly HashSet<string> _extractedCodes = [];

    public string StrategyName => "OrganizationCodeRemoval";

    public string Replace(TextMatch match)
    {
        var value = match.Value;
        var dotIndex = value.IndexOf('.');

        if (dotIndex <= 0 || dotIndex >= value.Length - 1)
            return value;

        var code = OrganizationCodeExtractor.ExtractCode(value);
        if (!string.IsNullOrEmpty(code))
        {
            _extractedCodes.Add(code);
        }

        return value.Replace(code!, new string('*', code!.Length));
    }

    /// <summary>
    /// Получает все извлеченные коды организаций
    /// </summary>
    public IReadOnlyCollection<string> GetExtractedCodes()
    {
        return _extractedCodes;
    }

    /// <summary>
    /// Очищает список извлеченных кодов
    /// </summary>
    public void ClearExtractedCodes()
    {
        _extractedCodes.Clear();
    }
}