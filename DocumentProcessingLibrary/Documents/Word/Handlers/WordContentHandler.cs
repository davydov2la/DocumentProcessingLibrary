using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessingLibrary.Documents.Word.Handlers;

/// <summary>
/// Обработчик основного содержимого Word документа
/// </summary>
public class WordContentHandler : BaseDocumentElementHandler<WordDocumentContext>
{
    public override string HandlerName => "WordContent";

    protected override ProcessingResult ProcessElement(WordDocumentContext context, ProcessingConfiguration config)
    {
        try
        {
            var content = context.Document.Content;
            var text = content.Text;
            var matches = FindAllMatches(text, config).ToList();

            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);

            foreach (var match in matches)
            {
                var replacement = config.ReplacementStrategy.Replace(match);
                var find = content.Find;
                try
                {
                    find.Execute(
                        FindText: match.Value,
                        MatchCase: config.Options.CaseSensitive,
                        ReplaceWith: replacement,
                        Replace: InteropWord.WdReplace.wdReplaceAll
                    );
                }
                finally
                {
                    if (find != null) Marshal.ReleaseComObject(find);
                }
            }

            return ProcessingResult.Successful(matches.Count, matches.Count);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки содержимого: {ex.Message}");
        }
    }
}