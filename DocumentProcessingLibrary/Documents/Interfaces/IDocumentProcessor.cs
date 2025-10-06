using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Documents.Interfaces;

/// <summary>
/// Процессор документа конкретного типа
/// </summary>
public interface IDocumentProcessor : IDisposable
{
    /// <summary>
    /// Поддерживаемые расширения файлов
    /// </summary>
    IEnumerable<string> SupportedExtensions { get; }

    /// <summary>
    /// Проверяет, может ли процессор обработать файл
    /// </summary>
    bool CanProcess(string filePath);

    /// <summary>
    /// Обрабатывает документ
    /// </summary>
    ProcessingResult Process(DocumentProcessingRequest request);

    /// <summary>
    /// Название процессора
    /// </summary>
    string ProcessorName { get; }
}