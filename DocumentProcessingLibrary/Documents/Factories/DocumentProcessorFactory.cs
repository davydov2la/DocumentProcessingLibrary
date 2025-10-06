using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.SolidWorks;
using DocumentProcessingLibrary.Documents.Word;
using DocumentProcessingLibrary.Documents.Word.OpenXml;

namespace DocumentProcessingLibrary.Documents.Factories;

/// <summary>
/// Фабрика для создания процессоров документов
/// </summary>
public class DocumentProcessorFactory : IDisposable
{
    private readonly List<IDocumentProcessor> _processors = new List<IDocumentProcessor>();
    private readonly bool _visible;
    private readonly bool _useOpenXml;
    private bool _disposed;
    /// <summary>
    /// Создает фабрику процессоров
    /// </summary>
    /// <param name="visible">Видимость приложений (для Interop)</param>
    /// <param name="useOpenXml">Использовать OpenXML вместо Interop для Word (по умолчанию true)</param>
    public DocumentProcessorFactory(bool visible = false, bool useOpenXml = true)
    {
        _visible = visible;
        _useOpenXml = useOpenXml;
    }
    /// <summary>
    /// Создает процессор для конкретного файла
    /// </summary>
    public IDocumentProcessor CreateProcessor(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));
        IDocumentProcessor? processor;
        if (_useOpenXml)
        {
            processor = TryCreateWordOpenXmlProcessor(filePath);
            if (processor != null)
            {
                _processors.Add(processor);
                return processor;
            }
        }
        processor = TryCreateWordInteropProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            return processor;
        }
        processor = TryCreateSolidWorksProcessor(filePath);
        if (processor != null)
        {
            _processors.Add(processor);
            return processor;
        }
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
        var extension = System.IO.Path.GetExtension(filePath)?.ToLowerInvariant();
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
        catch
        {
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
        catch
        {
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
            foreach (var processor in _processors)
            {
                try
                {
                    processor?.Dispose();
                }
                catch { }
            }
            _processors.Clear();
        }
        _disposed = true;
    }
}