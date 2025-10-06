using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Processing.Interfaces;

/// <summary>
/// Обработчик элементов документа (Chain of Responsibility)
/// </summary>
public interface IDocumentElementHandler<TContext>
{
    /// <summary>
    /// Устанавливает следующий обработчик в цепочке
    /// </summary>
    IDocumentElementHandler<TContext> SetNext(IDocumentElementHandler<TContext> handler);

    /// <summary>
    /// Обрабатывает элементы документа
    /// </summary>
    ProcessingResult Handle(TContext context, ProcessingConfiguration config);

    /// <summary>
    /// Название обработчика для логирования
    /// </summary>
    string HandlerName { get; }
}