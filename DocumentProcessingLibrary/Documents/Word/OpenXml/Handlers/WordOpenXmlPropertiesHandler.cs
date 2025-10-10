using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик свойств Word документа через OpenXML
/// </summary>
public class WordOpenXmlPropertiesHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlProperties";
    
    public WordOpenXmlPropertiesHandler(ILogger? logger = null) : base(logger) { }
    
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        
        try
        {
            var totalMatches = 0;
            var processed = 0;
            var propErrors = 0;
            
            var coreProps = context.Document.PackageProperties;
            if (coreProps != null)
            {
                try
                {
                    Logger?.LogDebug("Очистка встроенных свойств документа");
                    coreProps.Creator = "";
                    coreProps.Title = "";
                    coreProps.Subject = "";
                    coreProps.Keywords = "";
                    coreProps.Description = "";
                    coreProps.LastModifiedBy = "";
                    coreProps.Category = "";
                    processed += 7;
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось очистить встроенные свойства");
                    propErrors++;
                }
            }
            
            var customProps = context.Document.CustomFilePropertiesPart;
            if (customProps != null)
            {
                var properties = customProps.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>().ToList();
                Logger?.LogDebug("Найдено пользовательских свойств: {Count}", properties.Count);
                
                foreach (var prop in properties)
                {
                    try
                    {
                        var propValue = prop.InnerText;
                        if (!string.IsNullOrEmpty(propValue))
                        {
                            var matches = FindAllMatches(propValue, config).ToList();
                            if (matches.Any())
                            {
                                totalMatches += matches.Count;
                                var newValue = ReplaceText(propValue, matches, config.ReplacementStrategy);

                                prop.RemoveAllChildren();

                                if (!string.IsNullOrEmpty(newValue))
                                    prop.AppendChild(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(newValue));

                                processed += matches.Count;
                                Logger?.LogDebug("Обработано совпадений в свойстве '{Name}': {Count}", prop.Name,
                                    matches.Count);
                            }
                        }
                        else
                            prop.RemoveAllChildren();
                    }
                    catch (Exception ex)
                    {
                        Logger?.LogWarning(ex, "Не удалось обработать свойство");
                        propErrors++;
                    }
                }
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка свойств завершена");
            
            if (propErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {propErrors} свойств", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}", Logger, ex);
        }
    }
}