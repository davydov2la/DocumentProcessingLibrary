using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик заметок в SolidWorks чертежах
/// </summary>
public class SolidWorksNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksNotes";
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Drawing == null)
            return ProcessingResult.Successful(0, 0);
        var totalMatches = 0;
        var processed = 0;
        try
        {
            if (context.Drawing.GetSheetNames() is string[] sheetNames)
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
                                foreach (var noteObj in notes)
                                {
                                    if (noteObj is Note note)
                                    {
                                        try
                                        {
                                            var result = ProcessNote(note, config);
                                            totalMatches += result.MatchesFound;
                                            processed += result.MatchesProcessed;
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
                            Console.WriteLine($"Ошибка обработки вида: {ex.Message}");
                        }
                    }
                }

            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки заметок: {ex.Message}");
        }
    }
    private ProcessingResult ProcessNote(Note note, ProcessingConfiguration config)
    {
        try
        {
            string text = note.GetText();
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
            Console.WriteLine($"Ошибка обработки заметки: {ex.Message}");
            return ProcessingResult.Failed(ex.Message);
        }
    }
}