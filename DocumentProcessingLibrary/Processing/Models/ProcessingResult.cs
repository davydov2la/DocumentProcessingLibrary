namespace DocumentProcessingLibrary.Processing.Models;

/// <summary>
/// Результат обработки документа
/// </summary>
public class ProcessingResult
{
    public bool Success { get; set; }
    public int MatchesFound { get; set; }
    public int MatchesProcessed { get; set; }
    public List<string> Errors { get; set; } = [];
    public List<string> Warnings { get; set; } = [];
    public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

    public static ProcessingResult Successful(int found, int processed)
    {
        return new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };
    }

    public static ProcessingResult Failed(string error)
    {
        return new ProcessingResult
        {
            Success = false,
            Errors = [error]
        };
    }
}