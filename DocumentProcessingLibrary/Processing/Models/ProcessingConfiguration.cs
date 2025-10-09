using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Strategies.Replacement;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Processing.Models;

/// <summary>
/// Конфигурация обработки документа
/// </summary>
public class ProcessingConfiguration
{
    public List<ITextSearchStrategy> SearchStrategies { get; set; } = [];
    public ITextReplacementStrategy ReplacementStrategy { get; set; } = new RemoveReplacementStrategy();
    public ProcessingOptions Options { get; set; } = new ProcessingOptions();
    
    /// <summary>
    /// Логгер для процесса обработки (опционально)
    /// Использует стандартный Microsoft.Extensions.Logging.ILogger
    /// </summary>
    public ILogger? Logger { get; set; }
}