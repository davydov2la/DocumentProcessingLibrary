using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Strategies.Replacement;

/// <summary>
/// Комбинированная стратегия замены с условием
/// </summary>
public class CompositeReplacementStrategy : ITextReplacementStrategy
{
    private readonly Func<TextMatch, bool> _condition;
    private readonly ITextReplacementStrategy _trueStrategy;
    private readonly ITextReplacementStrategy _falseStrategy;

    public string StrategyName { get; }

    public CompositeReplacementStrategy(
        string name,
        Func<TextMatch, bool> condition,
        ITextReplacementStrategy trueStrategy,
        ITextReplacementStrategy falseStrategy)
    {
        StrategyName = name ?? throw new ArgumentNullException(nameof(name));
        _condition = condition ?? throw new ArgumentNullException(nameof(condition));
        _trueStrategy = trueStrategy ?? throw new ArgumentNullException(nameof(trueStrategy));
        _falseStrategy = falseStrategy ?? throw new ArgumentNullException(nameof(falseStrategy));
    }

    public string Replace(TextMatch match)
    {
        return _condition(match)
            ? _trueStrategy.Replace(match)
            : _falseStrategy.Replace(match);
    }
}