using DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Обработчик основного содержимого Word документа через OpenXML
/// </summary>
public class WordOpenXmlContentHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlContent";
    
    public WordOpenXmlContentHandler(ILogger? logger = null) :  base(logger) { }
    
    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        try
        {
            var body = context.Document.MainDocumentPart?.Document.Body;
            if (body == null)
                return ProcessingResult.Failed("Не удалось получить тело документа", Logger);
            
            var totalMatches = 0;
            var processed = 0;
            
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Logger?.LogDebug("Найдено параграфов: {Count}", paragraphs.Count);
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            
            var tables = body.Descendants<Table>().ToList();
            Logger?.LogDebug("Найдено таблиц: {Count}",  tables.Count);
            
            foreach (var table in tables)
            {
                var result = ProcessTable(table, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;
            }
            
            return ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка содержимого завершена");
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки содержимого: {ex.Message}", Logger, ex);
        }
    }
    
    /// <summary>
    /// Обрабатывает параграф (собирает текст из всех Run элементов)
    /// </summary>
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
            Logger?.LogDebug("В параграфе найдено совпадений: {Count}", found);
            
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
                    Logger?.LogWarning("Не удалось заменить текст в позиции {Position}: {Error}",
                        match.StartIndex, result.ErrorMessage);
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки параграфа");
            return ProcessingResult.Successful(found, processed);
        }
    }
    
    /// <summary>
    /// Обрабатывает таблицу (обрабатывает каждую ячейку как параграфы)
    /// </summary>
    private ProcessingResult ProcessTable(Table table, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;
        
        try
        {
            var cells = table.Descendants<TableCell>().ToList();
            
            foreach (var cell in cells)
            {
                var paragraphs = cell.Descendants<Paragraph>().ToList();
                
                foreach (var paragraph in paragraphs)
                {
                    var result = ProcessParagraph(paragraph, config);
                    found += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            
            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки таблицы");
            return ProcessingResult.Successful(found, processed);
        }
    }
}