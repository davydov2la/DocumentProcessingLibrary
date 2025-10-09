using Microsoft.Extensions.Logging;
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
    public Dictionary<string, object> Metadata { get; set; } = new();

    public static ProcessingResult Successful(int found, int processed)
    {
        return new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };
    }
    
    public static ProcessingResult Successful(int found, int processed, ILogger? logger = null, string? message = null)
    {
        var result = new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };

        logger?.LogInformation(message ?? "Обработка успешно завершена: {Совпадений найдено}/{Обработано}", found, processed);

        return result;
    }

    public static ProcessingResult Failed(string error)
    {
        return new ProcessingResult
        {
            Success = false,
            Errors = [error]
        };
    }

    public static ProcessingResult Failed(string error, ILogger? logger = null, Exception? ex = null)
    {
        var result = new ProcessingResult
        {
            Success = false
        };
        
        if (!string.IsNullOrWhiteSpace(error))
            result.Errors.Add(error);
        
        logger?.LogError(ex, "Ошибка при обработке: {Error}", error);
        return result;
    }
    
    public static ProcessingResult PartialSuccess(int found, int processed, string warning, ILogger? logger = null)
    {
        var result = new ProcessingResult
        {
            Success = true,
            MatchesFound = found,
            MatchesProcessed = processed
        };

        if (!string.IsNullOrWhiteSpace(warning))
            result.Warnings.Add(warning);

        logger?.LogWarning("Обработка завершена с предупреждением: {Warning}. Найдено {Found}, обработано {Processed}",
            warning, found, processed);

        return result;
    }

    public void AddWarning(string warning, ILogger? logger = null)
    {
        if (!string.IsNullOrWhiteSpace(warning) && !Warnings.Contains(warning))
        {
            Warnings.Add(warning);
            logger?.LogWarning("{Warning}", warning);
        }
    }
    
    public void AddError(string error, ILogger? logger = null, Exception? ex = null)
    {
        if (!string.IsNullOrWhiteSpace(error) && !Errors.Contains(error))
        {
            Errors.Add(error);
            Success = false;
            logger?.LogError(ex, "{Error}", error);
        }
    }
}