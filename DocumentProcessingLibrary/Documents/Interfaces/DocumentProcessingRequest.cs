using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Documents.Interfaces;

/// <summary>
/// Запрос на обработку документа
/// </summary>
public class DocumentProcessingRequest
{
    public string InputFilePath { get; set; } = string.Empty;
    public string OutputDirectory { get; set; } = string.Empty;
    public ProcessingConfiguration Configuration { get; set; } = new ProcessingConfiguration();
    public ExportOptions ExportOptions { get; set; } = new ExportOptions();
    public bool PreserveOriginal { get; set; } = true;
}