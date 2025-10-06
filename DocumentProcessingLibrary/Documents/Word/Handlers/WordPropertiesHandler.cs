using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using InteropWord =  Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace DocumentProcessingLibrary.Documents.Word.Handlers;

public class WordPropertiesHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    public override string HandlerName => "WordProperties";
    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties)
            return ProcessingResult.Successful(0, 0);
        var totalMatches = 0;
        var processed = 0;
        try
        {
            ProcessBuiltInProperties(context, ref processed);
            ProcessCustomProperties(context, config, ref totalMatches, ref processed);
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки свойств: {ex.Message}");
        }
    }
    private void ProcessBuiltInProperties(WordDocumentContext context, ref int processed)
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
                catch { }
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
                        }
                        else
                        {
                            prop.Value = "";
                        }
                    }
                    else
                    {
                        prop.Value = "";
                    }
                }
                catch { }
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