namespace DocumentProcessingLibrary.Processing.Models;

/// <summary>
/// Опции обработки документа
/// </summary>
public class ProcessingOptions
{
    public bool ProcessProperties { get; set; } = true;
    public bool ProcessTextBoxes { get; set; } = true;
    public bool ProcessNotes { get; set; } = true;
    public bool ProcessHeaders { get; set; } = true;
    public bool ProcessFooters { get; set; } = true;
    public int MinMatchLength { get; set; } = 8;
    public bool CaseSensitive { get; set; } = false;
}