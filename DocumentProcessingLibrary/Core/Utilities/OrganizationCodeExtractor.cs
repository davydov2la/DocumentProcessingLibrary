namespace DocumentProcessingLibrary.Core.Utilities;

/// <summary>
/// Утилита для извлечения кода организации из обозначения
/// </summary>
public static class OrganizationCodeExtractor
{
    /// <summary>
    /// Извлекает код организации (часть до первой точки)
    /// </summary>
    public static string? ExtractCode(string? designation)
    {
        if (string.IsNullOrEmpty(designation))
            return null;

        var dotIndex = designation.IndexOf('.');
            
        return dotIndex <= 0 ? null : designation[..dotIndex];
    }
}