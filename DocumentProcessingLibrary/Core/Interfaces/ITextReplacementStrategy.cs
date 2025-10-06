using DocumentProcessingLibrary.Core.Models;

namespace DocumentProcessingLibrary.Core.Interfaces;

/// <summary>
/// Стратегия замены найденного текста
/// </summary>
public interface ITextReplacementStrategy
{
    /// <summary>
    /// Выполняет замену найденного фрагмента
    /// </summary>
    string Replace(TextMatch match);

    /// <summary>
    /// Название стратегии
    /// </summary>
    string StrategyName { get; }
}