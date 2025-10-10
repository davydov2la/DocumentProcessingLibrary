using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Models;
using DocumentProcessingLibrary.Processing.Models;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;

public class ParagraphProcessor
{
    public static ProcessingResult ProcessParagraphWithReplacement(
        Paragraph paragraph,
        ProcessingConfiguration config,
        Func<string, ProcessingConfiguration, IEnumerable<TextMatch>> findMatches,
        Func<string, IEnumerable<TextMatch>, ITextReplacementStrategy, string> replaceText,
        ILogger? logger = null)
    {
        var found = 0;
        var processed = 0;
    
        try
        {
            var textElements = paragraph.Descendants<Text>().ToList();
            if (!textElements.Any()) return ProcessingResult.Successful(0, 0);
        
            var fullText = TextRunHelper.CollectText(textElements);
            if (string.IsNullOrEmpty(fullText)) return ProcessingResult.Successful(0, 0);
        
            var matches = findMatches(fullText, config).ToList();
            if (!matches.Any()) return ProcessingResult.Successful(0, 0);
        
            found = matches.Count;
        
            foreach (var match in matches.OrderByDescending(m => m.StartIndex))
            {
                var replacement = config.ReplacementStrategy.Replace(match);
                var currentTextElements = paragraph.Descendants<Text>().ToList();
                var currentElementMap = TextRunHelper.MapTextElements(currentTextElements);
            
                var result = TextRunHelper.ReplaceTextInRange(
                    currentElementMap, 
                    match.StartIndex, 
                    match.Length, 
                    replacement);
                
                if (result.Success) processed++;
                logger?.LogWarning("Не удалось заменить текст в позиции {Position}: {Error}",
                    match.StartIndex, result.ErrorMessage);
            }
        
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            logger?.LogError(ex, "Ошибка обработки параграфа");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки параграфа: {ex.Message}", logger);
        }
    }
}