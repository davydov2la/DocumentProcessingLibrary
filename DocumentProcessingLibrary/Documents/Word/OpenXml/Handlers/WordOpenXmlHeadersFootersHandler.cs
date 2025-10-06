using DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик колонтитулов в Word документе через OpenXML
/// </summary>
public class WordOpenXmlHeadersFootersHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlHeadersFooters";
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessHeaders && !config.Options.ProcessFooters)
            return ProcessingResult.Successful(0, 0);
        try
        {
            var mainPart = context.Document.MainDocumentPart;
            if (mainPart == null)
                return ProcessingResult.Failed("Не удалось получить основную часть документа");
            var totalMatches = 0;
            var processed = 0;
            if (config.Options.ProcessHeaders)
            {
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    var result = ProcessHeaderPart(headerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            if (config.Options.ProcessFooters)
            {
                foreach (var footerPart in mainPart.FooterParts)
                {
                    var result = ProcessFooterPart(footerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки колонтитулов: {ex.Message}");
        }
    }
    private ProcessingResult ProcessHeaderPart(HeaderPart headerPart, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        try
        {
            var paragraphs = headerPart.Header.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                found += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка обработки заголовка: {ex.Message}");
            return ProcessingResult.Successful(found, processed);
        }
    }
    private ProcessingResult ProcessFooterPart(FooterPart footerPart, ProcessingConfiguration config)
    {
        int found = 0;
        int processed = 0;
        try
        {
            var paragraphs = footerPart.Footer.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                found += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка обработки подвала: {ex.Message}");
            return ProcessingResult.Successful(found, processed);
        }
    }
    private ProcessingResult ProcessParagraph(Paragraph paragraph, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        try
        {
            var textElements = paragraph.Descendants<Text>().ToList();
            
            if (!textElements.Any())
                return ProcessingResult.Successful(0, 0);
            string fullText = TextRunHelper.CollectText(textElements);
            
            if (string.IsNullOrEmpty(fullText))
                return ProcessingResult.Successful(0, 0);
            var matches = FindAllMatches(fullText, config).ToList();
            
            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);
            found = matches.Count;
            var elementMap = TextRunHelper.MapTextElements(textElements);
            foreach (var match in matches.OrderByDescending(m => m.StartIndex))
            {
                var replacement = config.ReplacementStrategy.Replace(match);
                
                var result = TextRunHelper.ReplaceTextInRange(
                    elementMap,
                    match.StartIndex,
                    match.Length,
                    replacement );
                if (result.Success)
                {
                    processed++;
                }
            }
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка обработки параграфа: {ex.Message}");
            return ProcessingResult.Successful(found, processed);
        }
    }
}