using DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик колонтитулов в Word документе через OpenXML
/// </summary>
public class WordOpenXmlHeadersFootersHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlHeadersFooters";
    
    public WordOpenXmlHeadersFootersHandler(ILogger? logger = null) : base(logger) { } 
    
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (config.Options is { ProcessHeaders: false, ProcessFooters: false })
            return ProcessingResult.Successful(0, 0);
        
        try
        {
            var mainPart = context.Document.MainDocumentPart;
            if (mainPart == null)
                return ProcessingResult.Failed("Не удалось получить основную часть документа", Logger);
            
            var totalMatches = 0;
            var processed = 0;
            
            if (config.Options.ProcessHeaders)
            {
                Logger?.LogDebug("Обработка верхних колонтитулов: найдено {Count}", mainPart.HeaderParts.Count());
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    var result = ProcessHeaderPart(headerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            
            if (config.Options.ProcessFooters)
            {
                Logger?.LogDebug("Обработка нижних колонтитулов: найдено {Count}", mainPart.FooterParts.Count());
                foreach (var footerPart in mainPart.FooterParts)
                {
                    var result = ProcessFooterPart(footerPart, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            
            return ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка колонтитулов завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки колонтитулов: {ex.Message}", Logger, ex);
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
            Logger?.LogError(ex, "Ошибка обработки верхнего колонтитула");
            return ProcessingResult.Successful(found, processed);
        }
    }
    
    private ProcessingResult ProcessFooterPart(FooterPart footerPart, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        
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
            Logger?.LogError(ex, "Ошибка обработки нижнего колонтитула");
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
            
            var fullText = TextRunHelper.CollectText(textElements);
            
            if (string.IsNullOrEmpty(fullText))
                return ProcessingResult.Successful(0, 0);
            
            var matches = FindAllMatches(fullText, config).ToList();
            
            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);
            
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
                    replacement );
                if (result.Success)
                    processed++;
                else
                    Logger?.LogWarning("Не удалось заменить текст в колонтитуле на позиции {Position}: {Error}",
                        match.StartIndex, result.ErrorMessage);
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки параграфа в колонтитуле");
            return ProcessingResult.Successful(found, processed);
        }
    }
}