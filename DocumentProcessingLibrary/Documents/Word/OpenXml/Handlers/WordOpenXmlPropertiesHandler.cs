using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик свойств Word документа через OpenXML
/// </summary>
public class WordOpenXmlPropertiesHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlProperties";
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        try
        {
            var totalMatches = 0;
            var processed = 0;
            var coreProps = context.Document.PackageProperties;
            if (coreProps != null)
            {
                try
                {
                    coreProps.Creator = "";
                    coreProps.Title = "";
                    coreProps.Subject = "";
                    coreProps.Keywords = "";
                    coreProps.Description = "";
                    coreProps.LastModifiedBy = "";
                    coreProps.Category = "";
                    processed += 7;
                }
                catch { }
            }
            var customProps = context.Document.CustomFilePropertiesPart;
            if (customProps != null && customProps.Properties != null)
            {
                var properties = customProps.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>().ToList();
                
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
                                {
                                    prop.AppendChild(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(newValue));
                                }
                                
                                processed += matches.Count;
                            }
                        }
                        else
                        {
                            prop.RemoveAllChildren();
                        }
                    }
                    catch { }
                }
            }
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}");
        }
    }
}