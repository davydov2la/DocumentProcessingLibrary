using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Interfaces;

/// <summary>
/// Стратегия поиска текстовых фрагментов в документе
/// </summary>
public interface ITextSearchStrategy
{
    /// <summary>
    /// Находит все совпадения в тексте
    /// </summary>
    IEnumerable<TextMatch> FindMatches(string text);

    /// <summary>
    /// Название стратегии для логирования и отладки
    /// </summary>
    string StrategyName { get; }
}