using DocumentProcessingLibrary.Core.Strategies.Search;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml;

/// <summary>
/// Процессор для обработки Word документов через OpenXML
/// Не требует установленного Microsoft Word
/// </summary>
public class WordOpenXmlDocumentProcessor : ITwoPassDocumentProcessor
{
    private bool _disposed;
    public string ProcessorName => "WordOpenXmlDocumentProcessor";
    public IEnumerable<string> SupportedExtensions => [".docx", ".docm"];
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
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}");
        var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(request.InputFilePath));
        
        try
        {
            File.Copy(request.InputFilePath, tempFilePath, true);
            ProcessingResult result;
            using (var doc = WordprocessingDocument.Open(tempFilePath, true))
            {
                if (doc == null)
                    return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}");
                var context = new WordOpenXmlDocumentContext
                {
                    Document = doc, FilePath = tempFilePath
                };
                var contentHandler = new WordOpenXmlContentHandler();
                // var shapesHandler = new WordOpenXmlShapesHandler();
                var headersFootersHandler = new WordOpenXmlHeadersFootersHandler();
                var propertiesHandler = new WordOpenXmlPropertiesHandler();
                contentHandler
                    // .SetNext(shapesHandler)
                    .SetNext(headersFootersHandler)
                    .SetNext(propertiesHandler);
                result = contentHandler.Handle(context, request.Configuration);
                doc.MainDocumentPart?.Document?.Save();
            }
            if (request.ExportOptions.SaveModified)
            {
                File.Copy(tempFilePath, request.InputFilePath, true);
            }
            if (request.ExportOptions.ExportToPdf)
            {
                result.Warnings.Add("OpenXML процессор не поддерживает конвертацию в PDF. Используйте Interop процессор для экспорта в PDF.");
            }
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки документа: {ex.Message}");
        }
        finally
        {
            try
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }
            catch { }
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
        var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(request.InputFilePath));
        
        try
        {
            File.Copy(request.InputFilePath, tempFilePath, true);
            ProcessingResult firstPassResult;
            
            using (var doc = WordprocessingDocument.Open(tempFilePath, true))
            {
                if (doc == null)
                    return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}");
                var context = new WordOpenXmlDocumentContext
                {
                    Document = doc, FilePath = tempFilePath
                };
                var contentHandler = new WordOpenXmlContentHandler();
                // var shapesHandler = new WordOpenXmlShapesHandler();
                var headersFootersHandler = new WordOpenXmlHeadersFootersHandler();
                var propertiesHandler = new WordOpenXmlPropertiesHandler();
                contentHandler
                    // .SetNext(shapesHandler)
                    .SetNext(headersFootersHandler)
                    .SetNext(propertiesHandler);
                firstPassResult = contentHandler.Handle(context, twoPassConfig.FirstPassConfiguration);
                doc.MainDocumentPart?.Document?.Save();
            }
            var extractedCodes = twoPassConfig.CodeExtractionStrategy?.GetExtractedCodes() ?? new List<string>();
            if (extractedCodes.Count > 0)
            {
                using (var doc = WordprocessingDocument.Open(tempFilePath, true))
                {
                    var codeSearchStrategy = new OrganizationCodeSearchStrategy(extractedCodes);
                    twoPassConfig.SecondPassConfiguration.SearchStrategies.Add(codeSearchStrategy);
                    var context = new WordOpenXmlDocumentContext
                    {
                        Document = doc, FilePath = tempFilePath
                    };
                    var secondPassContentHandler = new WordOpenXmlContentHandler();
                    // var secondPassShapesHandler = new WordOpenXmlShapesHandler();
                    var secondPassHeadersFootersHandler = new WordOpenXmlHeadersFootersHandler();
                    secondPassContentHandler
                        // .SetNext(secondPassShapesHandler)
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
            {
                File.Copy(tempFilePath, request.InputFilePath, true);
            }
            if (request.ExportOptions.ExportToPdf)
            {
                firstPassResult.Warnings.Add("OpenXML процессор не поддерживает конвертацию в PDF.");
            }
            return firstPassResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка двухпроходной обработки: {ex.Message}");
        }
        finally
        {
            try
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }
            catch { }
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