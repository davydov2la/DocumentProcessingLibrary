using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;

/// <summary>
/// Вспомогательный класс для работы с текстовыми элементами в OpenXML
/// </summary>
public class TextRunHelper
{
    public class TextElementInfo
    {
        public Text TextElement { get; set; } = null!;
        public int StartIndex { get; set; }
        public int Length { get; set; }
        public string Content { get; set; } = string.Empty;
    }
    public class ReplacementResult
    {
        public int ElementsModified { get; set; }
        public bool Success { get; set; }
    }
    /// <summary>
    /// Собирает текст из всех Text элементов в контейнере
    /// </summary>
    public static string CollectText(IEnumerable<Text> textElements)
    {
        var sb = new StringBuilder();
        foreach (var text in textElements)
        {
            if (!string.IsNullOrEmpty(text.Text))
            {
                sb.Append(text.Text);
            }
        }
        return sb.ToString();
    }
    /// <summary>
    /// Создает карту текстовых элементов с их позициями в общем тексте
    /// </summary>
    public static List<TextElementInfo> MapTextElements(IEnumerable<Text> textElements)
    {
        var map = new List<TextElementInfo>();
        int currentPosition = 0;
        foreach (var text in textElements)
        {
            var content = text.Text ?? string.Empty;
            
            map.Add(new TextElementInfo
            {
                TextElement = text, StartIndex = currentPosition, Length = content.Length, Content = content
            });
            currentPosition += content.Length;
        }
        return map;
    }
    /// <summary>
    /// Заменяет текст в указанном диапазоне
    /// </summary>
    public static ReplacementResult ReplaceTextInRange(
        List<TextElementInfo> elementMap,
        int startIndex,
        int length,
        string replacement)
    {
        if (elementMap == null || elementMap.Count == 0)
            return new ReplacementResult { Success = false };
        var endIndex = startIndex + length;
        var elementsModified = 0;
        try
        {
            var affectedElements = elementMap
                .Where(e => e.StartIndex < endIndex && (e.StartIndex + e.Length) > startIndex)
                .ToList();
            if (!affectedElements.Any())
                return new ReplacementResult { Success = false };
            if (affectedElements.Count == 1)
            {
                var element = affectedElements[0];
                var relativeStart = startIndex - element.StartIndex;
                
                if (relativeStart < 0 || relativeStart + length > element.Content.Length)
                    return new ReplacementResult { Success = false };
                var newText = element.Content.Remove(relativeStart, length)
                    .Insert(relativeStart, replacement);
                
                element.TextElement.Space = SpaceProcessingModeValues.Preserve;
                element.TextElement.Text = newText;
                element.Content = newText;
                element.Length = newText.Length;
                
                elementsModified = 1;
            }
            else
            {
                var firstElement = affectedElements[0];
                var lastElement = affectedElements[^1];
                var firstElementCutStart = startIndex - firstElement.StartIndex;
                var lastElementCutEnd = (lastElement.StartIndex + lastElement.Length) - endIndex;
                if (firstElementCutStart < 0 || firstElementCutStart > firstElement.Content.Length)
                    return new ReplacementResult { Success = false };
                if (lastElementCutEnd < 0 || lastElementCutEnd > lastElement.Content.Length)
                    return new ReplacementResult { Success = false };
                var textBefore = firstElement.Content[..firstElementCutStart];
                var textAfter = lastElement.Content[^lastElementCutEnd..];
                firstElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                firstElement.TextElement.Text = textBefore + replacement;
                firstElement.Content = firstElement.TextElement.Text;
                firstElement.Length = firstElement.Content.Length;
                elementsModified++;
                for (var i = 1; i < affectedElements.Count - 1; i++)
                {
                    affectedElements[i].TextElement.Space = SpaceProcessingModeValues.Preserve;
                    affectedElements[i].TextElement.Text = string.Empty;
                    affectedElements[i].Content = string.Empty;
                    affectedElements[i].Length = 0;
                    elementsModified++;
                }
                if (affectedElements.Count > 1)
                {
                    lastElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    lastElement.TextElement.Text = textAfter;
                    lastElement.Content = textAfter;
                    lastElement.Length = textAfter.Length;
                    elementsModified++;
                }
                var indexShift = replacement.Length - length;
                var lastElementIndex = elementMap.IndexOf(lastElement);
                for (var i = lastElementIndex + 1; i < elementMap.Count; i++)
                {
                    elementMap[i].StartIndex += indexShift;
                }
            }
            return new ReplacementResult
            {
                Success = true, ElementsModified = elementsModified
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка замены текста: {ex.Message}");
            return new ReplacementResult { Success = false };
        }
    }
}