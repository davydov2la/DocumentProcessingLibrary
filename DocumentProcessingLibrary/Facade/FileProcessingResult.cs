namespace DocumentProcessingLibrary.Facade;

/// <summary>
/// Результат обработки одного файла
/// </summary>
public class FileProcessingResult
{
    public string FilePath { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public bool Success { get; set; }
    public int MatchesFound { get; set; }
    public int MatchesProcessed { get; set; }
    public string? Error { get; set; }
}