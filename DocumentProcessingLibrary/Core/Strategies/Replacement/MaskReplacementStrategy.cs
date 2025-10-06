using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия замены текста на звездочки
/// </summary>
public class MaskReplacementStrategy : ITextReplacementStrategy
{
    public string StrategyName => "Mask";
        
    public string Replace(TextMatch match) => new string('*', match.Length);
}