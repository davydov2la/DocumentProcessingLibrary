using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Strategies.Replacement;

namespace DocumentProcessingLibrary.Processing.Models;

/// <summary>
/// Конфигурация обработки документа
/// </summary>
public class ProcessingConfiguration
{
    public List<ITextSearchStrategy> SearchStrategies { get; set; } = [];
    public ITextReplacementStrategy ReplacementStrategy { get; set; } = new RemoveReplacementStrategy();
    public ProcessingOptions Options { get; set; } = new ProcessingOptions();
}