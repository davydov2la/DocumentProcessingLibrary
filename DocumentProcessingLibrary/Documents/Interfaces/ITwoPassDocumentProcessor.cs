using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Documents.Interfaces;

/// <summary>
/// Процессор документа с поддержкой двухпроходной обработки
/// </summary>
public interface ITwoPassDocumentProcessor : IDocumentProcessor
{
    /// <summary>
    /// Обрабатывает документ в два прохода
    /// </summary>
    ProcessingResult ProcessTwoPass(DocumentProcessingRequest request, TwoPassProcessingConfiguration twoPassConfig);
}