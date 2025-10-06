using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.Word.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentProcessingLibrary.Core.Strategies.Search;
using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessingLibrary.Documents.Word;

/// <summary>
/// Процессор для обработки Word документов
/// </summary>
public class WordDocumentProcessor : ITwoPassDocumentProcessor
{
    private InteropWord.Application? _wordApp;
    private bool _disposed;
    
    public string ProcessorName => "WordDocumentProcessor";
    public IEnumerable<string> SupportedExtensions => [".doc", ".docx", ".docm"];
    
    public WordDocumentProcessor(bool visible = false)
    {
        _wordApp = new InteropWord.Application
        {
            Visible = visible, 
            DisplayAlerts = InteropWord.WdAlertLevel.wdAlertsNone
        };
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
            return ProcessingResult.Failed("Word Application не инициализирован");
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}");

        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
            {
                File.Copy(request.InputFilePath, workingFilePath, true);
            }

            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}");

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var contentHandler = new WordContentHandler();
            var shapesHandler = new WordShapesHandler();
            var propertiesHandler = new WordPropertiesHandler();

            contentHandler
                .SetNext(shapesHandler)
                .SetNext(propertiesHandler);

            var result = contentHandler.Handle(context, request.Configuration);
            
            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
            {
                doc.Save();
            }

            if (request.ExportOptions.ExportToPdf)
            {
                var pdfFileName = GetPdfFileName(request);
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки документа: {ex.Message}");
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

        InteropWord.Document? doc = null;
        
        var workingFilePath = GetWorkingFilePath(request);
        object path = workingFilePath;
        object readOnly = false;
        object isVisible = false;
        
        try
        {
            if (request.PreserveOriginal && workingFilePath != request.InputFilePath)
            {
                File.Copy(request.InputFilePath, workingFilePath, true);
            }

            doc = _wordApp.Documents.Open(ref path, ReadOnly: ref readOnly, Visible: ref isVisible);
            if (doc == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}");

            var context = new WordDocumentContext
            {
                Document = doc, 
                Application = _wordApp
            };

            var firstPassHandler = new WordContentHandler();
            var firstPassShapesHandler = new WordShapesHandler();
            var firstPassPropertiesHandler = new WordPropertiesHandler();

            firstPassHandler
                .SetNext(firstPassShapesHandler)
                .SetNext(firstPassPropertiesHandler);

            var firstPassResult = firstPassHandler.Handle(context, twoPassConfig.FirstPassConfiguration);

            var extractedCodes = twoPassConfig.CodeExtractionStrategy?.GetExtractedCodes() ?? new List<string>();
            if (extractedCodes.Count > 0)
            {
                var codeSearchStrategy = new OrganizationCodeSearchStrategy(extractedCodes);
                twoPassConfig.SecondPassConfiguration.SearchStrategies.Add(codeSearchStrategy);

                var secondPassHandler = new WordContentHandler();
                var secondPassShapesHandler = new WordShapesHandler();

                secondPassHandler.SetNext(secondPassShapesHandler);

                var secondPassResult = secondPassHandler.Handle(context, twoPassConfig.SecondPassConfiguration);
                
                firstPassResult.MatchesFound += secondPassResult.MatchesFound;
                firstPassResult.MatchesProcessed += secondPassResult.MatchesProcessed;
                firstPassResult.Warnings.AddRange(secondPassResult.Warnings);
                firstPassResult.Metadata["CodesRemoved"] = extractedCodes.Count;
            }

            doc.Fields.Update();

            if (request.ExportOptions.SaveModified)
            {
                doc.Save();
            }

            if (request.ExportOptions.ExportToPdf)
            {
                var pdfFileName = GetPdfFileName(request);
                ExportToPdf(doc, pdfFileName, request.ExportOptions.Quality);
            }

            doc.Close(InteropWord.WdSaveOptions.wdDoNotSaveChanges);
            return firstPassResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка двухпроходной обработки: {ex.Message}");
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
            optimizeFor,
            InteropWord.WdExportRange.wdExportAllDocument
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