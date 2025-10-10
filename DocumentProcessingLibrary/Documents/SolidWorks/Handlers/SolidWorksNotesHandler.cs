using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик заметок в SolidWorks чертежах
/// </summary>
public class SolidWorksNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksNotes";
    
    public SolidWorksNotesHandler(ILogger? logger = null) : base(logger) { }
    
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Drawing == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var sheetErrors = 0;
        
        try
        {
            if (context.Drawing.GetSheetNames() is string[] sheetNames)
            {
                Logger?.LogDebug("Найдено листов: {Count}", sheetNames.Length);
                
                foreach (var sheetName in sheetNames)
                {
                    context.Drawing.ActivateSheet(sheetName);
                    var view = context.Drawing.GetFirstView() as View;
                    
                    while (view != null)
                    {
                        try
                        {
                            if (view.GetNotes() is object[] notes)
                            {
                                Logger?.LogDebug("Обработка вида '{ViewName}': найдено заметок {Count}", 
                                    view.Name, notes.Length);
                                
                                foreach (var noteObj in notes)
                                {
                                    if (noteObj is Note note)
                                    {
                                        try
                                        {
                                            var result = ProcessNote(note, config);
                                            totalMatches += result.MatchesFound;
                                            processed += result.MatchesProcessed;
                                            
                                            if (!result.Success)
                                                sheetErrors++;
                                        }
                                        finally
                                        {
                                            Marshal.ReleaseComObject(note);
                                        }
                                    }
                                }
                            }

                            view = (View)view.GetNextView();
                        }
                        catch (Exception ex)
                        {
                            Logger?.LogError(ex, "Ошибка обработки вида на листе '{SheetName}'", sheetName);
                            sheetErrors++;
                        }
                    }
                }
            }

            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка заметок завершена");
            
            if (sheetErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {sheetErrors} видов/заметок", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки заметок: {ex.Message}", Logger, ex);
        }
    }
    
    private ProcessingResult ProcessNote(Note note, ProcessingConfiguration config)
    {
        try
        {
            var text = note.GetText();
            if (string.IsNullOrEmpty(text))
                return ProcessingResult.Successful(0, 0);
            
            var matches = FindAllMatches(text, config).ToList();
            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);
            
            var newText = ReplaceText(text, matches, config.ReplacementStrategy);
            note.SetText(newText);
            
            return ProcessingResult.Successful(matches.Count, matches.Count);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки заметки");
            return ProcessingResult.PartialSuccess(0, 0, 
                $"Ошибка обработки заметки: {ex.Message}", Logger);
        }
    }
}