using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Стратегия замены с кастомной трансформацией
/// </summary>
public class TransformReplacementStrategy : ITextReplacementStrategy
{
    private readonly Func<TextMatch, string> _transformer;

    public string StrategyName { get; }

    public TransformReplacementStrategy(string name, Func<TextMatch, string> transformer)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _transformer = transformer ?? throw new ArgumentNullException(nameof(transformer));
    }

    public string Replace(TextMatch match) => _transformer(match);
}