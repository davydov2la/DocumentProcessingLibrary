using DocumentFormat.OpenXml.Packaging;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

/// <summary>
/// Контекст обработки Word документа через OpenXML
/// </summary>
public class WordOpenXmlDocumentContext
{
    public WordprocessingDocument Document { get; set; } = null!;
    public string FilePath { get; set; } = string.Empty;
}