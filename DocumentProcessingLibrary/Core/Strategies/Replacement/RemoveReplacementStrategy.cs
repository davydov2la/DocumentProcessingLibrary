using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия полного удаления найденного текста
/// </summary>
public class RemoveReplacementStrategy : ITextReplacementStrategy
{
    public string StrategyName => "Remove";
        
    public string Replace(TextMatch match) => string.Empty;
}