using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик заметок в блоках SolidWorks
/// </summary>
public class SolidWorksBlockNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksBlockNotes";
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Model == null)
            return ProcessingResult.Successful(0, 0);
        var totalMatches = 0;
        var processed = 0;
        try
        {
            var sketchMgr = context.Model.SketchManager;
            if (sketchMgr == null)
                return ProcessingResult.Successful(0, 0);
            try
            {
                var blocks = sketchMgr.GetSketchBlockDefinitions() as object[];
                if (blocks != null)
                {
                    foreach (var blockObj in blocks)
                    {
                        var block = blockObj as SketchBlockDefinition;
                        if (block != null)
                        {
                            try
                            {
                                var blockNotes = block.GetNotes() as object[];
                                if (blockNotes != null)
                                {
                                    foreach (var noteObj in blockNotes)
                                    {
                                        var note = noteObj as Note;
                                        if (note != null)
                                        {
                                            try
                                            {
                                                var text = note.GetText();
                                                if (!string.IsNullOrEmpty(text))
                                                {
                                                    var matches = FindAllMatches(text, config).ToList();
                                                    if (matches.Any())
                                                    {
                                                        totalMatches += matches.Count;
                                                        var newText = ReplaceText(text, matches, config.ReplacementStrategy);
                                                        note.SetText(newText);
                                                        processed += matches.Count;
                                                    }
                                                }
                                            }
                                            finally
                                            {
                                                Marshal.ReleaseComObject(note);
                                            }
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(block);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (sketchMgr != null)
                    Marshal.ReleaseComObject(sketchMgr);
            }
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки блоков: {ex.Message}");
        }
    }
}