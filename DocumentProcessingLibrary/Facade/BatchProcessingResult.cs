namespace DocumentProcessingLibrary.Facade;

/// <summary>
/// Результат пакетной обработки документов
/// </summary>
public class BatchProcessingResult
{
    public int TotalFiles { get; set; }
    public int SuccessfulFiles { get; set; }
    public int FailedFiles { get; set; }
    public List<FileProcessingResult> Results { get; set; } = new List<FileProcessingResult>();
}