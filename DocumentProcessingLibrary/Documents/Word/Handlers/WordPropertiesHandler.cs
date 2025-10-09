using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.Handlers;

public class WordPropertiesHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    public override string HandlerName => "WordProperties";
    
    public WordPropertiesHandler(ILogger? logger = null) : base(logger) { }
    
    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        try
        {
            Logger?.LogDebug("Обработка встроенных свойств");
            ProcessBuiltInProperties(context);
            
            Logger?.LogDebug("Обработка пользовательских свойств");
            ProcessCustomProperties(context, config, ref totalMatches, ref processed);
            
            return ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка свойств завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}", Logger, ex);
        }
    }
    
    private void ProcessBuiltInProperties(WordDocumentContext context)
    {
        dynamic builtins = context.Document.BuiltInDocumentProperties;
        if (builtins == null)
            return;
        
        try
        {
            for (var i = builtins.Count; i >= 1; i--)
            {
                var prop = builtins[i];
                try
                {
                    prop.Value = "";
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось обработать встроенное свойство №{Index}", (int)i);
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(builtins);
        }
    }
    private void ProcessCustomProperties(WordDocumentContext context, ProcessingConfiguration config, 
        ref int totalMatches, ref int processed)
    {
        dynamic customs = context.Document.CustomDocumentProperties;
        if (customs == null)
            return;
        
        try
        {
            Logger?.LogDebug("Найдено пользовательских свойств: {Count}", (int)customs.Count);
            
            for (var i = customs.Count; i >= 1; i--)
            {
                var prop = customs[i];
                try
                {
                    var propValue = prop.Value as string;
                    if (!string.IsNullOrEmpty(propValue))
                    {
                        var matches = FindAllMatches(propValue, config).ToList();
                        if (matches.Any())
                        {
                            totalMatches += matches.Count;
                            var newValue = ReplaceText(propValue, matches, config.ReplacementStrategy);
                            prop.Value = newValue;
                            processed += matches.Count;

                            Logger?.LogDebug("Обработано совпадений в свойстве #{Index}: {Count}", (int)i,
                                matches.Count);
                        }
                        else
                            prop.Value = "";
                    }
                    else
                        prop.Value = "";
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось обработать пользовательское свойство #{Index}", (int)i);
                }
                finally
                {
                    Marshal.ReleaseComObject(prop);
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(customs);
        }
    }
}