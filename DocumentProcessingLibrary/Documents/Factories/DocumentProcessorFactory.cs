using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.SolidWorks;
using DocumentProcessingLibrary.Documents.Word;
using DocumentProcessingLibrary.Documents.Word.OpenXml;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingLibrary.Documents.Factories;

/// <summary>
/// Фабрика для создания процессоров документов
/// </summary>
public class DocumentProcessorFactory : IDisposable
{
    private readonly List<IDocumentProcessor> _processors = new List<IDocumentProcessor>();
    private readonly bool _visible;
    private readonly bool _useOpenXml;
    private readonly ILogger? _logger;
    private bool _disposed;
    
    /// <summary>
    /// Создает фабрику процессоров
    /// </summary>
    /// <param name="visible">Видимость приложений (для Interop)</param>
    /// <param name="useOpenXml">Использовать OpenXML вместо Interop для Word (по умолчанию true)</param>
    /// <param name="logger">Логгер для процессов обработки</param>
    public DocumentProcessorFactory(bool visible = false, bool useOpenXml = true, ILogger? logger = null)
    {
        _visible = visible;
        _useOpenXml = useOpenXml;
        _logger = logger;
    }
    
    /// <summary>
    /// Создает процессор для конкретного файла
    /// </summary>
    public IDocumentProcessor CreateProcessor(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));
        
        _logger?.LogDebug("Создание процессора для файла: {FilePath}", filePath);
        
        IDocumentProcessor? processor;
        
        if (_useOpenXml)
        {
            processor = TryCreateWordOpenXmlProcessor(filePath);
            if (processor != null)
            {
                _processors.Add(processor);
                _logger?.LogInformation("Создан OpenXML процессор для: {FileName}", Path.GetFileName(filePath));
                return processor;
            }
        }
        
        processor = TryCreateWordInteropProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            _logger?.LogInformation("Создан Word Interop процессор для: {FileName}", Path.GetFileName(filePath));
            return processor;
        }
        
        processor = TryCreateSolidWorksProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            _logger?.LogInformation("Создан SolidWorks процессор для: {FileName}", Path.GetFileName(filePath));
            return processor;
        }
        
        _logger?.LogError("Не найден процессор для файла: {FilePath}", filePath);
        throw new NotSupportedException($"Не найден процессор для файла: {filePath}");
    }
    
    /// <summary>
    /// Получает все поддерживаемые расширения
    /// </summary>
    public IEnumerable<string> GetSupportedExtensions()
    {
        var wordExtensions = new[] { ".doc", ".docx", ".docm" };
        var solidWorksExtensions = new[] { ".slddrw", ".sldprt", ".sldasm" };
        return wordExtensions.Concat(solidWorksExtensions);
    }
    
    /// <summary>
    /// Проверяет, поддерживается ли файл
    /// </summary>
    public bool IsSupported(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return GetSupportedExtensions().Contains(extension);
    }
    
    private IDocumentProcessor? TryCreateWordOpenXmlProcessor(string filePath)
    {
        var processor = new WordOpenXmlDocumentProcessor();
        if (processor.CanProcess(filePath))
            return processor;
        
        processor.Dispose();
        return null;
    }
    
    private IDocumentProcessor? TryCreateWordInteropProcessor(string filePath)
    {
        try
        {
            var processor = new WordDocumentProcessor(_visible);
            if (processor.CanProcess(filePath))
                return processor;
            
            processor.Dispose();
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Не удалось создать Word Interop процессор");
            return null;
        }
    }
    
    private IDocumentProcessor? TryCreateSolidWorksProcessor(string filePath)
    {
        try
        {
            var processor = new SolidWorksDocumentProcessor(_visible);
            if (processor.CanProcess(filePath))
                return processor;
            
            processor.Dispose();
            return null;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Не удалось создать SolidWorks процессор");
            return null;
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
        
        if (disposing)
        {
            _logger?.LogDebug("Освобождение ресурсов фабрики. Активных процессоров: {Count}", _processors.Count);

            foreach (var processor in _processors)
            {
                try
                {
                    processor?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Ошибка при освобождении процессора");
                }
            }
            _processors.Clear();
        }
        
        _disposed = true;
    }
}