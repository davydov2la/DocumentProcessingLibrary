using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия замены на константное значение
/// </summary>
public class ConstantReplacementStrategy : ITextReplacementStrategy
{
    private readonly string _replacement;

    public string StrategyName { get; }

    public ConstantReplacementStrategy(string name, string replacement)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _replacement = replacement ?? string.Empty;
    }

    public string Replace(TextMatch match) => _replacement;
}