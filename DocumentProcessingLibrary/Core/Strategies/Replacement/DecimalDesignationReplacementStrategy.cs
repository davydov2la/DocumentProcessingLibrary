using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия замены для десятичных обозначений.
/// Удаляет код организации (все до первой точки включительно)
/// </summary>
public class DecimalDesignationReplacementStrategy : ITextReplacementStrategy
{
    public string StrategyName => "DecimalDesignation";

    public string Replace(TextMatch match)
    {
        var value = match.Value;
        var dotIndex = value.IndexOf('.');

        if (dotIndex <= 0 || dotIndex >= value.Length - 1)
            return value;

        var result = value.Substring(dotIndex + 1);
            
        return result;
    }
}