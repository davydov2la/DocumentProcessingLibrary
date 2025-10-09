using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.Word.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentProcessingLibrary.Core.Strategies.Search;
using Microsoft.Extensions.Logging;
using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessingLibrary.Documents.Word;

/// <summary>
/// Процессор для обработки Word документов
/// </summary>
public class WordDocumentProcessor : ITwoPassDocumentProcessor
{
    private InteropWord.Application? _wordApp;
    private bool _disposed;
    private readonly ILogger? _logger;
    
    public string ProcessorName => "WordDocumentProcessor";
    public IEnumerable<string> SupportedExtensions => [".doc", ".docx", ".docm"];
    
    public WordDocumentProcessor(bool visible = false, ILogger? logger = null)
    {
        _wordApp = new InteropWord.Application
        {
            Visible = visible, 
            DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone
        };
        _logger = logger;
    }
    
    public bool CanProcess(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return extension == ".doc" || extension == ".docx" || extension == ".docm";
    }
    
    public ProcessingResult Process(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_wordApp == null)
            return ProcessingResult.Failed("Word Application не инициализирован", _logger);
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}", _logger);
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}", _logger);

        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало обработки документа: {FilePath}", request.InputFilePath);
        
        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
                File.Copy(request.InputFilePath, workingFilePath, true);

            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var contentHandler = new WordContentHandler(logger);
            var shapesHandler = new WordShapesHandler(logger);
            var propertiesHandler = new WordPropertiesHandler(logger);

            contentHandler
                .SetNext(shapesHandler)
                .SetNext(propertiesHandler);

            var result = contentHandler.Handle(context, request.Configuration);
            
            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
                doc.Save();

            if (request.ExportOptions.ExportToPdf)
            {
                logger?.LogInformation("Экспорт в формат PDF");
                var pdfFileName = GetPdfFileName(request);
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            
            logger?.LogInformation("Обработка завершена: найдено {Found}, обработано {Processed}",
                result.MatchesFound, result.MatchesProcessed);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки документа: {ex.Message}", logger, ex);
        }
        finally
        {
            if (doc != null)
            {
                Marshal.ReleaseComObject(doc);
            }
        }
    }
    
    /// <summary>
    /// Обрабатывает документ в два прохода: сначала обозначения, затем коды
    /// </summary>
    public ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_wordApp == null)
            return ProcessingResult.Failed("Word Application не инициализирован");
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        
        var logger = request.Configuration.Logger ?? _logger;
        logger?.LogInformation("Начало двухпроходной обработки документа: {FilePath}", request.InputFilePath);

        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
                File.Copy(request.InputFilePath, workingFilePath, true);
            
            logger?.LogInformation("=== ПЕРВЫЙ ПРОХОД ===");
            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}", logger);

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var firstPassHandler = new WordContentHandler(logger);
            var firstPassShapesHandler = new WordShapesHandler(logger);
            var firstPassPropertiesHandler = new WordPropertiesHandler(logger);

            firstPassHandler
                .SetNext(firstPassShapesHandler)
                .SetNext(firstPassPropertiesHandler);

            var firstPassResult = firstPassHandler.Handle(context, twoPassConfig.FirstPassConfiguration);

            var extractedCodes = twoPassConfig.CodeExtractionStrategy?.GetExtractedCodes() ?? new List<string>();
            logger?.LogInformation("Извлечено кодов организаций: {Count}", extractedCodes.Count);
            
            if (extractedCodes.Count > 0)
            {
                logger?.LogInformation("=== ВТОРОЙ ПРОХОД ===");
                
                var codeSearchStrategy = new OrganizationCodeSearchStrategy(extractedCodes);
                twoPassConfig.SecondPassConfiguration.SearchStrategies.Add(codeSearchStrategy);

                var secondPassHandler = new WordContentHandler(logger);
                var secondPassShapesHandler = new WordShapesHandler(logger);

                secondPassHandler.SetNext(secondPassShapesHandler);

                var secondPassResult = secondPassHandler.Handle(context, twoPassConfig.SecondPassConfiguration);
                
                firstPassResult.MatchesFound += secondPassResult.MatchesFound;
                firstPassResult.MatchesProcessed += secondPassResult.MatchesProcessed;
                firstPassResult.Warnings.AddRange(secondPassResult.Warnings);
                firstPassResult.Metadata["CodesRemoved"] = extractedCodes.Count;
            }

            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
                doc.Save();

            if (request.ExportOptions.ExportToPdf)
            {
                logger?.LogInformation("Экспорт в формат PDF");
                var pdfFileName = GetPdfFileName(request);
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
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
            if (doc != null)
            {
                Marshal.ReleaseComObject(doc);
            }
        }
    }

    /// <summary>
    /// Определяет путь к файлу для работы (оригинал или копия)
    /// </summary>
    private string GetWorkingFilePath(DocumentProcessingRequest request)
    {
        if (request.PreserveOriginal)
        {
            var fileName = Path.GetFileNameWithoutExtension(request.InputFilePath);
            var extension = Path.GetExtension(request.InputFilePath);
            var processedFileName = $"{fileName}_processed{extension}";
            return Path.Combine(request.OutputDirectory, processedFileName);
        }
        
        return request.InputFilePath;
    }

    /// <summary>
    /// Формирует имя PDF файла
    /// </summary>
    private string GetPdfFileName(DocumentProcessingRequest request)
    {
        if (!string.IsNullOrEmpty(request.ExportOptions.PdfFileName))
            return request.ExportOptions.PdfFileName;

        var docName = Path.GetFileNameWithoutExtension(
            request.PreserveOriginal 
                ? Path.GetFileName(GetWorkingFilePath(request)) 
                : request.InputFilePath
        );
        
        return Path.Combine(request.OutputDirectory, docName + ".pdf");
    }

    private void ExportToPdf(InteropWord.Document doc, string filename, PdfQuality quality)
    {
        var exportFormat = InteropWord.WdExportFormat.wdExportFormatPDF;
        var optimizeFor = quality switch
        {
            PdfQuality.Draft => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
            PdfQuality.Standard => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
            PdfQuality.HighQuality => InteropWord.WdExportOptimizeFor.wdExportOptimizeForPrint,
            _ => InteropWord.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        };

        doc.ExportAsFixedFormat(
            filename,
            exportFormat,
            false,
            optimizeFor
        );
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

        if (disposing && _wordApp != null)
        {
            try
            {
                _wordApp.Quit(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            }
            finally
            {
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
            }
        }

        _disposed = true;
    }
    
    ~WordDocumentProcessor()
    {
        Dispose(false);
    }
}