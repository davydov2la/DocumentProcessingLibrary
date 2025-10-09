using DocumentProcessingLibrary.Core.Strategies.Search;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml;

/// <summary>
/// Процессор для обработки Word документов через OpenXML
/// Не требует установленного Microsoft Word
/// </summary>
public class WordOpenXmlDocumentProcessor : ITwoPassDocumentProcessor
{
    private bool _disposed;
    private readonly ILogger? _logger;
    
    public string ProcessorName => "WordOpenXmlDocumentProcessor";
    public IEnumerable<string> SupportedExtensions => [".docx", ".docm"];

    public WordOpenXmlDocumentProcessor(ILogger? logger = null)
    {
        _logger = logger;
    }
    
    public bool CanProcess(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return extension is ".docx" or ".docm";
    }
    
    public ProcessingResult Process(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}", _logger);
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}", _logger);
        
        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало обработки документа: {FilePath}", request.InputFilePath);

        var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + Path.GetExtension(request.InputFilePath));
        
        try
        {
            File.Copy(request.InputFilePath, tempFilePath, true);
            ProcessingResult result;
            
            using (var doc = WordprocessingDocument.Open(tempFilePath, true))
            {
                if (doc == null)
                    return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

                var context = new WordOpenXmlDocumentContext
                {
                    Document = doc, 
                    FilePath = tempFilePath
                };

                var contentHandler = new WordOpenXmlContentHandler(logger);
                var headersFootersHandler = new WordOpenXmlHeadersFootersHandler(logger);
                var propertiesHandler = new WordOpenXmlPropertiesHandler(logger);

                contentHandler
                    .SetNext(headersFootersHandler)
                    .SetNext(propertiesHandler);

                result = contentHandler.Handle(context, request.Configuration);
                doc.MainDocumentPart?.Document?.Save();
            }

            if (request.ExportOptions.SaveModified)
                SaveProcessedDocument(request, tempFilePath, logger);

            if (request.ExportOptions.ExportToPdf)
                result.AddWarning("OpenXML процессор не поддерживает конвертацию в PDF. Используйте Interop процессор для экспорта в PDF.", logger);
            
            logger?.LogInformation("Обработка завершена успешно: найдено {Found}, обработано {Processed}",
                result.MatchesFound, result.MatchesProcessed);

            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки документа: {ex.Message}", logger, ex);
        }
        finally
        {
            try
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "Не удалось удалить временный файл");
            }
        }
    }
    
    /// <summary>
    /// Двухпроходная обработка для OpenXML
    /// </summary>
    public ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        
        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало двухпроходной обработки документа: {FilePath}", request.InputFilePath);

        var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + Path.GetExtension(request.InputFilePath));
        
        try
        {
            File.Copy(request.InputFilePath, tempFilePath, true);
            ProcessingResult firstPassResult;
            
            logger?.LogInformation("=== ПЕРВЫЙ ПРОХОД ===");
            using (var doc = WordprocessingDocument.Open(tempFilePath, true))
            {
                if (doc == null)
                    return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

                var context = new WordOpenXmlDocumentContext
                {
                    Document = doc, 
                    FilePath = tempFilePath
                };

                var contentHandler = new WordOpenXmlContentHandler(logger);
                var headersFootersHandler = new WordOpenXmlHeadersFootersHandler(logger);
                var propertiesHandler = new WordOpenXmlPropertiesHandler(logger);

                contentHandler
                    .SetNext(headersFootersHandler)
                    .SetNext(propertiesHandler);

                firstPassResult = contentHandler.Handle(context, twoPassConfig.FirstPassConfiguration);
                doc.MainDocumentPart?.Document?.Save();
            }

            var extractedCodes = twoPassConfig.CodeExtractionStrategy?.GetExtractedCodes() ?? new List<string>();
            logger?.LogInformation("Извлечено кодов организаций: {Count}", extractedCodes.Count);
            if (extractedCodes.Count > 0)
            {
                logger?.LogInformation("=== ВТОРОЙ ПРОХОД ===");
                
                using (var doc = WordprocessingDocument.Open(tempFilePath, true))
                {
                    var codeSearchStrategy = new OrganizationCodeSearchStrategy(extractedCodes);
                    twoPassConfig.SecondPassConfiguration.SearchStrategies.Add(codeSearchStrategy);

                    var context = new WordOpenXmlDocumentContext
                    {
                        Document = doc, 
                        FilePath = tempFilePath
                    };

                    var secondPassContentHandler = new WordOpenXmlContentHandler(logger);
                    var secondPassHeadersFootersHandler = new WordOpenXmlHeadersFootersHandler(logger);

                    secondPassContentHandler
                        .SetNext(secondPassHeadersFootersHandler);

                    var secondPassResult = secondPassContentHandler.Handle(context, twoPassConfig.SecondPassConfiguration);
                    
                    firstPassResult.MatchesFound += secondPassResult.MatchesFound;
                    firstPassResult.MatchesProcessed += secondPassResult.MatchesProcessed;
                    firstPassResult.Warnings.AddRange(secondPassResult.Warnings);
                    firstPassResult.Metadata["CodesRemoved"] = extractedCodes.Count;
                    
                    doc.MainDocumentPart?.Document?.Save();
                }
            }

            if (request.ExportOptions.SaveModified)
                SaveProcessedDocument(request, tempFilePath, logger);

            if (request.ExportOptions.ExportToPdf)
                firstPassResult.AddWarning("OpenXML процессор не поддерживает конвертацию в PDF.", logger);
            
            logger?.LogInformation("Двухпроходная обработка завершена: найдено {Found}, обработано {Processed}", 
                firstPassResult.MatchesFound, firstPassResult.MatchesProcessed);

            return firstPassResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка двухпроходной обработки: {ex.Message}", logger, ex);
        }
        finally
        {
            try
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "Не удалось удалить временный файл");
            }
        }
    }

    /// <summary>
    /// Сохраняет обработанный документ с учетом флага PreserveOriginal
    /// </summary>
    private void SaveProcessedDocument(DocumentProcessingRequest request, string tempFilePath, ILogger? logger)
    {
        if (request.PreserveOriginal)
        {
            var fileName = Path.GetFileNameWithoutExtension(request.InputFilePath);
            var extension = Path.GetExtension(request.InputFilePath);
            var processedFileName = $"{fileName}_processed{extension}";
            var outputPath = Path.Combine(request.OutputDirectory, processedFileName);

            File.Copy(tempFilePath, outputPath, true);
            logger?.LogInformation("Обработанный документ сохранен: {Path}", outputPath);
            
        }
        else
        {
            File.Copy(tempFilePath, request.InputFilePath, true);
            logger?.LogInformation("Оригинальный документ перезаписан: {Path}", request.InputFilePath);
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;
        _disposed = true;
    }

    ~WordOpenXmlDocumentProcessor()
    {
        Dispose(false);
    }
}