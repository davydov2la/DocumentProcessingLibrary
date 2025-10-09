using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;

/// <summary>
/// Вспомогательный класс для работы с текстовыми элементами в OpenXML
/// ИСПРАВЛЕНО: правильное удаление пустых Run элементов
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
        public string? ErrorMessage { get; set; }
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
                TextElement = text, 
                StartIndex = currentPosition, 
                Length = content.Length, 
                Content = content
            });

            currentPosition += content.Length;
        }

        return map;
    }

    /// <summary>
    /// Заменяет текст в указанном диапазоне
    /// ИСПРАВЛЕНО: корректное удаление пустых Run элементов
    /// </summary>
    public static ReplacementResult ReplaceTextInRange(
        List<TextElementInfo> elementMap,
        int startIndex,
        int length,
        string replacement)
    {
        if (elementMap == null || elementMap.Count == 0)
        {
            return new ReplacementResult 
            { 
                Success = false,
                ErrorMessage = "Карта элементов пуста"
            };
        }

        var endIndex = startIndex + length;
        var elementsModified = 0;

        try
        {
            var affectedElements = elementMap
                .Where(e => e.StartIndex < endIndex && (e.StartIndex + e.Length) > startIndex)
                .ToList();

            if (!affectedElements.Any())
            {
                return new ReplacementResult 
                { 
                    Success = false,
                    ErrorMessage = "Не найдено затронутых элементов"
                };
            }

            if (affectedElements.Count == 1)
            {
                // Простой случай: замена в одном элементе
                var element = affectedElements[0];
                var relativeStart = startIndex - element.StartIndex;
                
                if (relativeStart < 0 || relativeStart + length > element.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Выход за границы элемента"
                    };
                }

                var newText = element.Content.Remove(relativeStart, length)
                    .Insert(relativeStart, replacement);
                
                // ИСПРАВЛЕНИЕ: Если текст стал пустым, удаляем весь Run
                if (string.IsNullOrEmpty(newText))
                {
                    RemoveEmptyRun(element.TextElement);
                }
                else
                {
                    element.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    element.TextElement.Text = newText;
                }
                
                element.Content = newText;
                element.Length = newText.Length;
                elementsModified = 1;
            }
            else
            {
                // Сложный случай: замена через несколько элементов
                var firstElement = affectedElements[0];
                var lastElement = affectedElements[^1];

                var firstElementCutStart = startIndex - firstElement.StartIndex;
                var lastElementCutEnd = (lastElement.StartIndex + lastElement.Length) - endIndex;

                if (firstElementCutStart < 0 || firstElementCutStart > firstElement.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Неверная позиция в первом элементе"
                    };
                }

                if (lastElementCutEnd < 0 || lastElementCutEnd > lastElement.Content.Length)
                {
                    return new ReplacementResult 
                    { 
                        Success = false,
                        ErrorMessage = $"Неверная позиция в последнем элементе"
                    };
                }

                var textBefore = firstElement.Content[..firstElementCutStart];
                var textAfter = lastElement.Content[^lastElementCutEnd..];

                // Устанавливаем текст в первый элемент
                var firstElementNewText = textBefore + replacement;
                
                // ИСПРАВЛЕНИЕ: Если первый элемент стал пустым, удаляем его Run
                if (string.IsNullOrEmpty(firstElementNewText))
                {
                    RemoveEmptyRun(firstElement.TextElement);
                }
                else
                {
                    firstElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    firstElement.TextElement.Text = firstElementNewText;
                }
                
                firstElement.Content = firstElementNewText;
                firstElement.Length = firstElementNewText.Length;
                elementsModified++;

                // ИСПРАВЛЕНИЕ: Удаляем средние элементы полностью (вместе с Run)
                for (var i = 1; i < affectedElements.Count - 1; i++)
                {
                    RemoveEmptyRun(affectedElements[i].TextElement);
                    affectedElements[i].Content = string.Empty;
                    affectedElements[i].Length = 0;
                    elementsModified++;
                }

                // Последний элемент
                if (affectedElements.Count > 1)
                {
                    // ИСПРАВЛЕНИЕ: Если последний элемент стал пустым, удаляем его Run
                    if (string.IsNullOrEmpty(textAfter))
                    {
                        RemoveEmptyRun(lastElement.TextElement);
                    }
                    else
                    {
                        lastElement.TextElement.Space = SpaceProcessingModeValues.Preserve;
                        lastElement.TextElement.Text = textAfter;
                    }
                    
                    lastElement.Content = textAfter;
                    lastElement.Length = textAfter.Length;
                    elementsModified++;
                }
            }

            return new ReplacementResult
            {
                Success = true, 
                ElementsModified = elementsModified
            };
        }
        catch (Exception ex)
        {
            return new ReplacementResult 
            { 
                Success = false,
                ErrorMessage = $"Исключение при замене: {ex.Message}"
            };
        }
    }

    /// <summary>
    /// НОВЫЙ МЕТОД: Удаляет пустой Run элемент полностью
    /// </summary>
    private static void RemoveEmptyRun(Text textElement)
    {
        try
        {
            // Находим родительский Run элемент
            var run = textElement.Ancestors<Run>().FirstOrDefault();
            
            if (run != null)
            {
                // ВАЖНО: Удаляем весь Run, а не только Text
                run.Remove();
            }
            else
            {
                // Если Run не найден, удаляем хотя бы сам Text элемент
                textElement.Remove();
            }
        }
        catch
        {
            // В крайнем случае просто очищаем текст
            textElement.Text = string.Empty;
        }
    }
}