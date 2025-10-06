using InteropWord = Microsoft.Office.Interop.Word;

namespace DocumentProcessingLibrary.Documents.Word.Handlers;

/// <summary>
/// Контекст обработки Word документа
/// </summary>
public class WordDocumentContext
{
    public InteropWord.Document Document { get; set; } = null!;
    public InteropWord.Application Application { get; set; } = null!;
}