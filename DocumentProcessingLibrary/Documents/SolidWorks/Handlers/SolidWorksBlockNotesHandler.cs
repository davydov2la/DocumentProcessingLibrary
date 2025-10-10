using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;
using SolidWorks.Interop.sldworks;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик заметок в блоках SolidWorks
/// </summary>
public class SolidWorksBlockNotesHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksBlockNotes";
    
    public SolidWorksBlockNotesHandler(ILogger? logger = null) : base(logger) { }

    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessNotes || context.Model == null)
            return ProcessingResult.Successful(0, 0);
        
        var totalMatches = 0;
        var processed = 0;
        var blockErrors = 0;
        
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
                    Logger?.LogDebug("Найдено блоков: {Count}", blocks.Length);
                    
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
                                            catch (Exception ex)
                                            {
                                                Logger?.LogWarning(ex, "Не удалось обработать заметку в блоке");
                                                blockErrors++;
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
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, 
                "Обработка блоков завершена");
            
            if (blockErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {blockErrors} заметок в блоках", Logger);
            
            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки блоков: {ex.Message}",  Logger, ex);
        }
    }
}