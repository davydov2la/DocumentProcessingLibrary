using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик свойств SolidWorks документов
/// </summary>
public class SolidWorksPropertiesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksProperties";
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties || context.Model == null)
            return ProcessingResult.Successful(0, 0);
        int totalMatches = 0;
        int processed = 0;
        try
        {
            ClearCustomProperties(context.Model, "", ref totalMatches, ref processed, config);
            if (context.Drawing?.GetSheetNames() is string[] sheetNames)
                foreach (var sheetName in sheetNames)
                {
                    ClearCustomProperties(context.Model, sheetName, ref totalMatches, ref processed, config);
                }
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}");
        }
    }
    private void ClearCustomProperties(ModelDoc2 model, string configName, ref int totalMatches, ref int processed, ProcessingConfiguration config)
    {
        CustomPropertyManager? cusProps = null;
        try
        {
            cusProps = model.Extension.CustomPropertyManager[configName];
            var namesObj = cusProps?.GetNames();
            var names = namesObj as string[];
            if (names == null || names.Length == 0)
                return;
            foreach (var name in names)
            {
                if (cusProps == null) continue;
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
                    {
                        cusProps.Set2(name, "");
                    }
                }
                else
                {
                    cusProps.Set2(name, "");
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