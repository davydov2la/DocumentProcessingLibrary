using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using InteropWord =  Microsoft.Office.Interop.Word;

namespace DocumentProcessingLibrary.Documents.Word.Handlers;

public class WordShapesHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    public override string HandlerName => "WordShapes";

    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessTextBoxes)
            return ProcessingResult.Successful(0, 0);

        var totalMatches = 0;
        var processed = 0;

        try
        {
            foreach (InteropWord.Shape shape in context.Document.Shapes)
            {
                try
                {
                    if (shape.TextFrame?.HasText != 0)
                    {
                        var text = shape.TextFrame.TextRange.Text;
                        var matches = FindAllMatches(text, config).ToList();

                        if (matches.Any())
                        {
                            totalMatches += matches.Count;
                            var newText = ReplaceText(text, matches, config.ReplacementStrategy);
                            shape.TextFrame.TextRange.Text = newText;
                            processed += matches.Count;
                        }
                    }
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                }
            }

            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки фигур: {ex.Message}");
        }
    }
}