using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;
using DocumentProcessingLibrary.Processing.Interfaces;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Processing.Handlers;

/// <summary>
/// Базовый класс для обработчиков элементов документа
/// Реализует Chain of Responsibility
/// </summary>
public abstract class BaseDocumentElementHandler<TContext> : IDocumentElementHandler<TContext>
{
    private IDocumentElementHandler<TContext>? _nextHandler;
    protected ILogger? Logger { get; }

    public abstract string HandlerName { get; }

    protected BaseDocumentElementHandler(ILogger? logger = null)
    {
        Logger = logger;
    }

    public IDocumentElementHandler<TContext> SetNext(IDocumentElementHandler<TContext> handler)
    {
        _nextHandler = handler;
        return handler;
    }

    public ProcessingResult Handle(TContext context, ProcessingConfiguration config)
    {
        Logger?.LogDebug("Начало обработки в {HandlerName}", HandlerName);
        
        var result = ProcessElement(context, config);
        
        Logger?.LogDebug("Завершение обработки в {HandlerName}: найдено {Found}, обработано {Processed}",
            HandlerName, result.MatchesFound, result.MatchesProcessed);

        if (_nextHandler != null)
        {
            var nextResult = _nextHandler.Handle(context, config);
            return MergeResults(result, nextResult);
        }

        return result;
    }

    /// <summary>
    /// Обрабатывает конкретный элемент документа
    /// </summary>
    protected abstract ProcessingResult ProcessElement(TContext context, ProcessingConfiguration config);

    /// <summary>
    /// Находит все совпадения в тексте используя настроенные стратегии
    /// </summary>
    protected IEnumerable<TextMatch> FindAllMatches(string text, ProcessingConfiguration config)
    {
        if (string.IsNullOrEmpty(text))
            return Enumerable.Empty<TextMatch>();

        var allMatches = new List<TextMatch>();

        foreach (var strategy in config.SearchStrategies)
        {
            var matches = strategy.FindMatches(text);
            allMatches.AddRange(matches.Where(m => m.Length >= config.Options.MinMatchLength));
        }

        return allMatches;
    }

    /// <summary>
    /// Выполняет замену текста
    /// </summary>
    protected string ReplaceText(string originalText, IEnumerable<TextMatch> matches, ITextReplacementStrategy strategy)
    {
        if (string.IsNullOrEmpty(originalText) || strategy == null)
            return originalText;

        var result = originalText;
        var sortedMatches = matches.OrderByDescending(m => m.StartIndex).ToList();

        foreach (var match in sortedMatches)
        {
            try
            {
                var replacement = strategy.Replace(match);
                result = result.Remove(match.StartIndex, match.Length)
                    .Insert(match.StartIndex, replacement);
            }
            catch (Exception ex)
            {
                Logger?.LogError(ex, "Ошибка замены в {HandlerName} на позиции {Position}", HandlerName, match.StartIndex);
            }
        }

        return result;
    }

    /// <summary>
    /// Объединяет результаты обработки
    /// ИСПРАВЛЕНО: Корректная логика Success
    /// </summary>
    private ProcessingResult MergeResults(ProcessingResult first, ProcessingResult second)
    {
        // Success = true только если оба успешны И нет критических ошибок
        var mergedSuccess = first.Success && second.Success;
        
        return new ProcessingResult
        {
            Success = mergedSuccess, 
            MatchesFound = first.MatchesFound + second.MatchesFound, 
            MatchesProcessed = first.MatchesProcessed + second.MatchesProcessed, 
            Errors = first.Errors.Concat(second.Errors).Distinct().ToList(), 
            Warnings = first.Warnings.Concat(second.Warnings).Distinct().ToList(), 
            Metadata = first.Metadata.Concat(second.Metadata)
                .GroupBy(kvp => kvp.Key)
                .ToDictionary(g => g.Key, g => g.First().Value)
        };
    }
}