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
            var paragraphErrors = 0;
            var tableErrors = 0;
            
            var paragraphs = body.Descendants<Paragraph>().ToList();
            Logger?.LogDebug("Найдено параграфов: {Count}", paragraphs.Count);
            
            foreach (var paragraph in paragraphs)
            {
                var result = ProcessParagraph(paragraph, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;

                if (!result.Success)
                    paragraphErrors++;
            }
            
            var tables = body.Descendants<Table>().ToList();
            Logger?.LogDebug("Найдено таблиц: {Count}",  tables.Count);
            
            foreach (var table in tables)
            {
                var result = ProcessTable(table, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;
                
                if (!result.Success)
                    tableErrors++;
            }
            
            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка содержимого завершена");
            
            if (paragraphErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {paragraphErrors} параграфов", Logger);
            
            if (tableErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {tableErrors} таблиц", Logger);
            
            return finalResult;
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
        return ParagraphProcessor.ProcessParagraphWithReplacement(
            paragraph, 
            config, 
            FindAllMatches, 
            ReplaceText, 
            Logger);
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
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки таблицы: {ex.Message}", Logger);
        }
    }
}