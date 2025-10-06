using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Strategies.Replacement;
using DocumentProcessingLibrary.Core.Strategies.Search;
using DocumentProcessingLibrary.Documents.Factories;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Facade;

/// <summary>
/// Главный фасад библиотеки для анонимизации документов
/// </summary>
public class DocumentAnonymizer : IDisposable
{
    private readonly DocumentProcessorFactory _factory;
    private bool _disposed;
    public DocumentAnonymizer(bool visible = false, bool useOpenXml = true)
    {
        _factory = new DocumentProcessorFactory(visible, useOpenXml);
    }
    /// <summary>
    /// Анонимизирует документ с настройками по умолчанию
    /// </summary>
    public ProcessingResult AnonymizeDocument(string inputFilePath, string outputDirectory)
    {
        var configuration = CreateDefaultConfiguration();
        return AnonymizeDocument(inputFilePath, outputDirectory, configuration);
    }
    /// <summary>
    /// Анонимизирует документ с пользовательской конфигурацией
    /// </summary>
    public ProcessingResult AnonymizeDocument(
        string inputFilePath,
        string outputDirectory,
        ProcessingConfiguration configuration)
    {
        ValidateInputs(inputFilePath, outputDirectory);
        var request = new DocumentProcessingRequest
        {
            InputFilePath = inputFilePath, OutputDirectory = outputDirectory, Configuration = configuration, ExportOptions = new ExportOptions
            {
                ExportToPdf = true, SaveModified = true, Quality = PdfQuality.Standard
            }, PreserveOriginal = true
        };
        return AnonymizeDocument(request);
    }
    /// <summary>
    /// Анонимизирует документ с полным контролем настроек
    /// </summary>
    public ProcessingResult AnonymizeDocument(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        ValidateInputs(request.InputFilePath, request.OutputDirectory);
        using (var processor = _factory.CreateProcessor(request.InputFilePath))
        {
            return processor.Process(request);
        }
    }
    /// <summary>
    /// Анонимизирует документ с удалением кодов организаций (двухпроходная обработка)
    /// </summary>
    public ProcessingResult AnonymizeDocumentWithCodeRemoval(
        string inputFilePath,
        string outputDirectory)
    {
        ValidateInputs(inputFilePath, outputDirectory);
        var codeRemovalStrategy = new OrganizationCodeRemovalStrategy();
        var firstPassConfig = new ProcessingConfiguration
        {
            SearchStrategies = new List<ITextSearchStrategy>
            {
                CommonSearchStrategies.DecimalDesignations, CommonSearchStrategies.PersonNames
            }, ReplacementStrategy = codeRemovalStrategy, Options = new ProcessingOptions
            {
                ProcessProperties = true, ProcessTextBoxes = true, ProcessNotes = true, ProcessHeaders = true, ProcessFooters = true, MinMatchLength = 8, CaseSensitive = false
            }
        };
        var secondPassConfig = new ProcessingConfiguration
        {
            SearchStrategies = new List<ITextSearchStrategy>(),
            ReplacementStrategy = new RemoveReplacementStrategy(),
            Options = new ProcessingOptions
            {
                ProcessProperties = false,
                ProcessTextBoxes = true,
                ProcessNotes = true,
                ProcessHeaders = true,
                ProcessFooters = true,
                MinMatchLength = 1,
                CaseSensitive = false
            }
        };
        var twoPassConfig = new TwoPassProcessingConfiguration
        {
            FirstPassConfiguration = firstPassConfig, SecondPassConfiguration = secondPassConfig, CodeExtractionStrategy = codeRemovalStrategy
        };
        var request = new DocumentProcessingRequest
        {
            InputFilePath = inputFilePath, OutputDirectory = outputDirectory, Configuration = firstPassConfig, ExportOptions = new ExportOptions
            {
                ExportToPdf = true, SaveModified = true, Quality = PdfQuality.Standard
            }, PreserveOriginal = true
        };
        using (var processor = _factory.CreateProcessor(inputFilePath))
        {
            if (processor is ITwoPassDocumentProcessor twoPassProcessor)
            {
                return twoPassProcessor.ProcessTwoPass(request, twoPassConfig);
            }
            else
            {
                return processor.Process(request);
            }
        }
    }
    /// <summary>
    /// Пакетная обработка файлов
    /// </summary>
    public BatchProcessingResult AnonymizeBatch(
        IEnumerable<string> filePaths,
        string outputDirectory,
        ProcessingConfiguration? configuration = null)
    {
        if (filePaths == null)
            throw new ArgumentNullException(nameof(filePaths));
        ValidateOutputDirectory(outputDirectory);
        configuration = configuration ?? CreateDefaultConfiguration();
        var results = new List<FileProcessingResult>();
        var fileList = filePaths.ToList();
        foreach (var filePath in fileList)
        {
            var fileResult = new FileProcessingResult
            {
                FilePath = filePath, FileName = Path.GetFileName(filePath)
            };
            try
            {
                if (!File.Exists(filePath))
                {
                    fileResult.Success = false;
                    fileResult.Error = "Файл не найден";
                    results.Add(fileResult);
                    continue;
                }
                if (!_factory.IsSupported(filePath))
                {
                    fileResult.Success = false;
                    fileResult.Error = "Неподдерживаемый формат файла";
                    results.Add(fileResult);
                    continue;
                }
                var processingResult = AnonymizeDocument(filePath, outputDirectory, configuration);
                fileResult.Success = processingResult.Success;
                fileResult.MatchesFound = processingResult.MatchesFound;
                fileResult.MatchesProcessed = processingResult.MatchesProcessed;
                fileResult.Error = processingResult.Success ? null : string.Join("; ", processingResult.Errors);
            }
            catch (Exception ex)
            {
                fileResult.Success = false;
                fileResult.Error = ex.Message;
            }
            results.Add(fileResult);
        }
        return new BatchProcessingResult
        {
            TotalFiles = fileList.Count, SuccessfulFiles = results.Count(r => r.Success), FailedFiles = results.Count(r => !r.Success), Results = results
        };
    }
    /// <summary>
    /// Создает конфигурацию по умолчанию
    /// </summary>
    public static ProcessingConfiguration CreateDefaultConfiguration()
    {
        return new ProcessingConfiguration
        {
            SearchStrategies = new List<ITextSearchStrategy>
            {
                CommonSearchStrategies.DecimalDesignations, CommonSearchStrategies.PersonNames
            }, ReplacementStrategy = new DecimalDesignationReplacementStrategy(), Options = new ProcessingOptions
            {
                ProcessProperties = true, ProcessTextBoxes = true, ProcessNotes = true, ProcessHeaders = true, ProcessFooters = true, MinMatchLength = 8, CaseSensitive = false
            }
        };
    }
    public static ProcessingConfiguration CreateCustomConfiguration(
        IEnumerable<ITextSearchStrategy> searchStrategies,
        ITextReplacementStrategy replacementStrategy,
        ProcessingOptions? options = null)
    {
        if (searchStrategies == null || !searchStrategies.Any())
            throw new ArgumentException("Необходимо указать хотя бы одну стратегию поиска", nameof(searchStrategies));
        if (replacementStrategy == null)
            throw new ArgumentNullException(nameof(replacementStrategy));
        return new ProcessingConfiguration
        {
            SearchStrategies = searchStrategies.ToList(), ReplacementStrategy = replacementStrategy, Options = options ?? new ProcessingOptions()
        };
    }
    public IEnumerable<string> GetSupportedFormats()
    {
        return _factory.GetSupportedExtensions();
    }
    private void ValidateInputs(string inputFilePath, string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(inputFilePath))
            throw new ArgumentException("Не указан путь к входному файлу", nameof(inputFilePath));
        if (!File.Exists(inputFilePath))
            throw new FileNotFoundException("Входной файл не найден", inputFilePath);
        ValidateOutputDirectory(outputDirectory);
        if (!_factory.IsSupported(inputFilePath))
            throw new NotSupportedException($"Формат файла не поддерживается: {Path.GetExtension(inputFilePath)}");
    }
    private void ValidateOutputDirectory(string outputDirectory)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException("Не указана выходная директория", nameof(outputDirectory));
        if (!Directory.Exists(outputDirectory))
        {
            try
            {
                Directory.CreateDirectory(outputDirectory);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Не удалось создать выходную директорию: {ex.Message}", ex);
            }
        }
    }
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;
        if (disposing)
        {
            _factory?.Dispose();
        }
        _disposed = true;
    }
}