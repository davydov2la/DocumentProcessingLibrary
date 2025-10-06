namespace DocumentProcessingLibrary.Documents.Interfaces;

/// <summary>
/// Опции экспорта документа
/// </summary>
public class ExportOptions
{
    public bool ExportToPdf { get; set; } = true;
    public bool SaveModified { get; set; } = true;
    public string? PdfFileName { get; set; }
    public PdfQuality Quality { get; set; } = PdfQuality.Standard;
}