using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик свойств SolidWorks документов
/// </summary>
public class SolidWorksPropertiesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksProperties";
    
    public SolidWorksPropertiesHandler(ILogger? logger = null) : base(logger) { }

    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties || context.Model == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var propErrors = 0;
        
        try
        {
            Logger?.LogDebug("Обработка свойств модели");

            ClearCustomProperties(context.Model, "", ref totalMatches, ref processed, ref propErrors, config);

            if (context.Drawing?.GetSheetNames() is string[] sheetNames)
            {
                Logger?.LogDebug("Обработка свойств листов: {Count}", sheetNames.Length);

                foreach (var sheetName in sheetNames)
                    ClearCustomProperties(context.Model, sheetName, ref totalMatches, ref processed, ref propErrors, config);
            }

            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка свойств завершена");
            
            
            if (propErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {propErrors} свойств", Logger);

            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}", Logger, ex);
        }
    }
    private void ClearCustomProperties(ModelDoc2 model, string configName, ref int totalMatches, 
        ref int processed, ref int errors, ProcessingConfiguration config)
    {
        CustomPropertyManager? cusProps = null;
        
        try
        {
            cusProps = model.Extension.CustomPropertyManager[configName];
            var namesObj = cusProps?.GetNames();
            var names = namesObj as string[];
            
            if (names == null || names.Length == 0)
                return;
            
            Logger?.LogDebug("Обработка {Count} свойств для конфигурации '{Config}'", 
                names.Length, string.IsNullOrEmpty(configName) ? "default" : configName);
            
            foreach (var name in names)
            {
                if (cusProps == null) continue;
                try
                {
                    cusProps.Get5(name, false, out var value, out _, out _);
                    
                    if (!string.IsNullOrEmpty(value))
                    {
                        var matches = FindAllMatches(value, config).ToList();
                        if (matches.Any())
                        {
                            totalMatches += matches.Count;
                            var newValue = ReplaceText(value, matches, config.ReplacementStrategy);
                            cusProps.Set2(name, newValue);
                            processed += matches.Count;
                        }
                        else
                            cusProps.Set2(name, "");
                    }
                    else
                        cusProps.Set2(name, "");
                }
                catch (Exception ex)
                {
                    Logger?.LogWarning(ex, "Не удалось обработать свойство '{PropertyName}'", name);
                    errors++;
                }
            }
        }
        finally
        {
            if (cusProps != null)
                Marshal.ReleaseComObject(cusProps);
        }
    }
}