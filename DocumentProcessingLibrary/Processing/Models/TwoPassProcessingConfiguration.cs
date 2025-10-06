using DocumentProcessingLibrary.Core.Strategies.Replacement;

namespace DocumentProcessingLibrary.Processing.Models;

/// <summary>
/// Конфигурация для двухпроходной обработки документа
/// (сначала обозначения с извлечением кодов, затем удаление кодов)
/// </summary>
public class TwoPassProcessingConfiguration
{
    /// <summary>
    /// Конфигурация первого прохода (обработка обозначений)
    /// </summary>
    public ProcessingConfiguration FirstPassConfiguration { get; set; } = new ProcessingConfiguration();

    /// <summary>
    /// Конфигурация второго прохода (удаление кодов)
    /// </summary>
    public ProcessingConfiguration SecondPassConfiguration { get; set; } = new ProcessingConfiguration();

    /// <summary>
    /// Стратегия для первого прохода, которая будет извлекать коды
    /// </summary>
    public OrganizationCodeRemovalStrategy? CodeExtractionStrategy { get; set; }
}