using DocumentProcessingLibrary.Documents.Word.OpenXml.Utilities;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocumentProcessingLibrary.Documents.Word.OpenXml.Handlers;

public class WordOpenXmlShapesHandler : BaseDocumentElementHandler<WordOpenXmlDocumentContext>
{
    public override string HandlerName => "WordOpenXmlShapes";

    public WordOpenXmlShapesHandler(ILogger? logger = null) : base(logger) { }

    protected override ProcessingResult ProcessElement(WordOpenXmlDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessTextBoxes)
            return ProcessingResult.Successful(0, 0);

        try
        {
            var body = context.Document.MainDocumentPart?.Document?.Body;
            if (body == null)
                return ProcessingResult.Failed("Не удалось получить тело документа", Logger);

            var totalMatches = 0;
            var processed = 0;
            var shapeErrors = 0;

            var vmlShapes = body.Descendants<DocumentFormat.OpenXml.Vml.Shape>().ToList();
            Logger?.LogDebug("Найдено VML Shapes: {Count}", vmlShapes.Count);

            foreach (var shape in vmlShapes)
            {
                var result = ProcessVmlShape(shape, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;

                if (!result.Success)
                    shapeErrors++;
            }

            var inlineShapes = body.Descendants<Wp.Inline>().ToList();
            Logger?.LogDebug("Найдено Inline Shapes: {Count}", inlineShapes.Count);

            foreach (var inline in inlineShapes)
            {
                var result = ProcessDrawingMLElement(inline, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;

                if (!result.Success)
                    shapeErrors++;
            }

            var anchorShapes = body.Descendants<Wp.Anchor>().ToList();
            Logger?.LogDebug("Найдено Anchor Shapes: {Count}", anchorShapes.Count);

            foreach (var anchor in anchorShapes)
            {
                var result = ProcessDrawingMLElement(anchor, config);
                totalMatches += result.MatchesFound;
                processed += result.MatchesProcessed;

                if (!result.Success)
                    shapeErrors++;
            }

            var finalResult = ProcessingResult.Successful(totalMatches, processed, Logger, "Обработка фигур завершена");

            if (shapeErrors > 0)
                finalResult.AddWarning($"Не удалось обработать {shapeErrors} фигур", Logger);

            return finalResult;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Критическая ошибка обработки фигур: {ex.Message}", Logger, ex);
        }
    }

    private ProcessingResult ProcessVmlShape(DocumentFormat.OpenXml.Vml.Shape shape, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;

        try
        {
            var paragraphs = shape.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                var textElements = paragraph.Descendants<Text>().ToList();

                if (!textElements.Any())
                    continue;

                string fullText = TextRunHelper.CollectText(textElements);

                if (string.IsNullOrEmpty(fullText))
                    continue;

                var matches = FindAllMatches(fullText, config).ToList();

                if (!matches.Any())
                    continue;

                found += matches.Count;

                foreach (var match in matches.OrderByDescending(m => m.StartIndex))
                {
                    var replacement = config.ReplacementStrategy.Replace(match);

                    var currentTextElements = paragraph.Descendants<Text>().ToList();
                    var currentElementMap = TextRunHelper.MapTextElements(currentTextElements);

                    var result = TextRunHelper.ReplaceTextInRange(
                        currentElementMap,
                        match.StartIndex,
                        match.Length,
                        replacement);

                    if (result.Success)
                        processed++;
                    else
                        Logger?.LogWarning("Не удалось заменить текст в VML Shape на позиции {Position}: {Error}",
                            match.StartIndex, result.ErrorMessage);
                }
            }

            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки VML Shape");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки VML Shape: {ex.Message}", Logger);
        }
    }

    private ProcessingResult ProcessDrawingMLElement(OpenXmlElement element, ProcessingConfiguration config)
    {
        var found = 0;
        var processed = 0;

        try
        {
            var textElements = element.Descendants<A.Text>().ToList();

            if (!textElements.Any())
                return ProcessingResult.Successful(0, 0);

            var fullTextBuilder = new System.Text.StringBuilder();
            foreach (var text in textElements)
            {
                if (!string.IsNullOrEmpty(text.Text))
                    fullTextBuilder.Append(text.Text);
            }

            var fullText = fullTextBuilder.ToString();

            if (string.IsNullOrEmpty(fullText))
                return ProcessingResult.Successful(0, 0);

            var matches = FindAllMatches(fullText, config).ToList();

            if (!matches.Any())
                return ProcessingResult.Successful(0, 0);

            found = matches.Count;

            var elementMap = CreateDrawingTextMap(textElements);

            foreach (var match in matches.OrderByDescending(m => m.StartIndex))
            {
                var replacement = config.ReplacementStrategy.Replace(match);

                if (ReplaceInDrawingText(elementMap, match.StartIndex, match.Length, replacement))
                    processed++;
                else
                    Logger?.LogWarning("Не удалось заменить текст в DrawingML на позиции {Position}",
                        match.StartIndex);
            }

            return ProcessingResult.Successful(found, processed);
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка обработки DrawingML");
            return ProcessingResult.PartialSuccess(found, processed,
                $"Ошибка обработки DrawingML: {ex.Message}", Logger);
        }
    }

    private List<DrawingTextInfo> CreateDrawingTextMap(List<A.Text> textElements)
    {
        var map = new List<DrawingTextInfo>();
        var position = 0;

        foreach (var text in textElements)
        {
            var content = text.Text ?? string.Empty;
            map.Add(new DrawingTextInfo
            {
                Element = text,
                StartIndex = position,
                Length = content.Length,
                Content = content
            });
            position += content.Length;
        }

        return map;
    }

    private bool ReplaceInDrawingText(List<DrawingTextInfo> map, int startIndex, int length, string replacement)
    {
        var endIndex = startIndex + length;
        var affected = map.Where(e => e.StartIndex < endIndex && (e.StartIndex + e.Length) > startIndex).ToList();

        if (!affected.Any())
            return false;

        try
        {
            if (affected.Count == 1)
            {
                var elem = affected[0];
                var relStart = startIndex - elem.StartIndex;

                if (relStart < 0 || relStart + length > elem.Content.Length)
                    return false;

                elem.Element.Text = elem.Content.Remove(relStart, length).Insert(relStart, replacement);
                elem.Content = elem.Element.Text;
                elem.Length = elem.Content.Length;
            }
            else
            {
                var first = affected[0];
                var last = affected[^1];
                var cutStart = startIndex - first.StartIndex;
                var cutEnd = (last.StartIndex + last.Length) - endIndex;

                if (cutStart < 0 || cutStart > first.Content.Length)
                    return false;
                if (cutEnd < 0 || cutEnd > last.Content.Length)
                    return false;

                var before = first.Content[..cutStart];
                var after = last.Content[^cutEnd..];

                first.Element.Text = before + replacement;
                first.Content = first.Element.Text;
                first.Length = first.Content.Length;

                for (var i = 1; i < affected.Count - 1; i++)
                {
                    affected[i].Element.Text = string.Empty;
                    affected[i].Content = string.Empty;
                    affected[i].Length = 0;
                }

                if (affected.Count > 1)
                {
                    last.Element.Text = after;
                    last.Content = after;
                    last.Length = after.Length;
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            Logger?.LogError(ex, "Ошибка замены текста в DrawingML");
            return false;
        }
    }

    private class DrawingTextInfo
    {
        public A.Text Element { get; set; } = null!;
        public int StartIndex { get; set; }
        public int Length { get; set; }
        public string Content { get; set; } = string.Empty;
    }
}