using DocumentProcessingLibrary.Core.Interfaces;
using DocumentProcessingLibrary.Core.Strategies.Replacement;
using DocumentProcessingLibrary.Core.Strategies.Search;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Facade;
using DocumentProcessingLibrary.Processing.Models;

namespace Examples
{
    /// <summary>
    /// Полный пример ручной настройки библиотеки DocumentAnonymizer
    /// </summary>
    public class ManualConfigurationExample
    {
        /// <summary>
        /// ГЛАВНЫЙ ПРИМЕР: Полная ручная настройка с двухпроходной обработкой
        /// </summary>
        public static void FullManualConfiguration()
        {
            string inputFile = @"/Users/paveldavydov/RiderProjects/DocumentProcessingLibrary/ProcessingTest/ЛР1 — копия.docx";
            string outputDir = @"/Users/paveldavydov/RiderProjects/DocumentProcessingLibrary/ProcessingTest/Output";

            Console.WriteLine("=== ПОЛНАЯ РУЧНАЯ НАСТРОЙКА БИБЛИОТЕКИ ===\n");

            // ========================================================================
            // ШАГ 1: Создаем стратегию для ПЕРВОГО ПРОХОДА (обозначения + имена)
            // ========================================================================

            // 1.1 Создаем стратегию поиска обозначений
            var designationSearchStrategy = new RegexSearchStrategy(
                "CustomDesignations",
                new RegexPattern(
                    "DesignationFullFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+\.(?:[0-9]{2,2}\.){2,}[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?"
                ),
                new RegexPattern(
                    "DesignationShortFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+-[А-Я0-9-]+\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                ),
                new RegexPattern(
                    "DesignationMinimalFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9]+\.[0-9]{2,2}\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                )
            );

            // 1.2 Создаем стратегию поиска имен
            var nameSearchStrategy = new RegexSearchStrategy(
                "PersonNames",
                new RegexPattern(
                    "SurnameFirst",
                    @"[А-Я][а-я]+\s[А-Я]\.\s?[А-Я]\."
                ),
                new RegexPattern(
                    "InitialsFirst",
                    @"[А-Я]\.\s?[А-Я]\.\s?[А-Я][а-я]+"
                )
            );

            // 1.3 Создаем стратегию замены с извлечением кодов организаций
            var codeExtractionStrategy = new OrganizationCodeRemovalStrategy();

            // 1.4 Создаем стратегию для удаления имен
            var nameRemovalStrategy = new RemoveReplacementStrategy();

            Console.WriteLine("Шаг 1: Настройка первого прохода");
            Console.WriteLine("  ✓ Стратегия поиска обозначений создана");
            Console.WriteLine("  ✓ Стратегия поиска имен создана");
            Console.WriteLine("  ✓ Стратегия замены с извлечением кодов создана\n");

            // ========================================================================
            // ШАГ 2: Настраиваем конфигурацию ПЕРВОГО ПРОХОДА
            // ========================================================================

            var firstPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>
                {
                    designationSearchStrategy,  // Ищем обозначения
                    nameSearchStrategy          // Ищем имена
                },
                ReplacementStrategy = new CompositeReplacementStrategy(
                    "FirstPassReplacement",
                    match => match.MatchType.Contains("Designation"),  // Проверяем тип совпадения
                    codeExtractionStrategy,  // Для обозначений - извлекаем код и заменяем
                    nameRemovalStrategy      // Для имен - просто удаляем
                ),
                Options = new ProcessingOptions
                {
                    ProcessProperties = true,   // Обрабатывать свойства документа
                    ProcessTextBoxes = true,    // Обрабатывать текстовые блоки
                    ProcessNotes = true,        // Обрабатывать заметки (для SolidWorks)
                    ProcessHeaders = true,      // Обрабатывать колонтитулы
                    ProcessFooters = true,      // Обрабатывать подвалы
                    MinMatchLength = 5,         // Минимальная длина совпадения (5 символов)
                    CaseSensitive = false       // Регистронезависимый поиск
                }
            };

            Console.WriteLine("Шаг 2: Конфигурация первого прохода");
            Console.WriteLine("  ✓ Добавлены стратегии поиска: обозначения, имена");
            Console.WriteLine("  ✓ Настроена композитная стратегия замены");
            Console.WriteLine("  ✓ Включены все опции обработки\n");

            // ========================================================================
            // ШАГ 3: Настраиваем конфигурацию ВТОРОГО ПРОХОДА
            // ========================================================================

            var secondPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>(), // Заполнится после 1-го прохода
                ReplacementStrategy = new RemoveReplacementStrategy(), // Просто удаляем коды
                Options = new ProcessingOptions
                {
                    ProcessProperties = false,  // Свойства уже обработаны
                    ProcessTextBoxes = true,    // Коды могут быть в текстовых блоках
                    ProcessNotes = true,        // Коды могут быть в заметках
                    ProcessHeaders = true,      // Коды могут быть в колонтитулах
                    ProcessFooters = true,      // Коды могут быть в подвалах
                    MinMatchLength = 1,         // Коды могут быть короткими (например "АБ")
                    CaseSensitive = true        // Коды чувствительны к регистру
                }
            };

            Console.WriteLine("Шаг 3: Конфигурация второго прохода");
            Console.WriteLine("  ✓ Стратегия удаления кодов создана");
            Console.WriteLine("  ✓ Минимальная длина = 1 (для коротких кодов)");
            Console.WriteLine("  ✓ Регистрозависимый поиск включен\n");

            // ========================================================================
            // ШАГ 4: Создаем конфигурацию двухпроходной обработки
            // ========================================================================

            var twoPassConfig = new TwoPassProcessingConfiguration
            {
                FirstPassConfiguration = firstPassConfig,
                SecondPassConfiguration = secondPassConfig,
                CodeExtractionStrategy = codeExtractionStrategy
            };

            Console.WriteLine("Шаг 4: Конфигурация двухпроходной обработки создана\n");

            // ========================================================================
            // ШАГ 5: Создаем запрос на обработку документа
            // ========================================================================

            var processingRequest = new DocumentProcessingRequest
            {
                InputFilePath = inputFile,
                OutputDirectory = outputDir,
                Configuration = firstPassConfig, // Используется для основной обработки
                ExportOptions = new ExportOptions
                {
                    ExportToPdf = true,              // Конвертировать в PDF
                    SaveModified = true,             // Сохранить измененный документ
                    PdfFileName = null,              // Имя PDF (null = автоматическое)
                    Quality = PdfQuality.HighQuality // Высокое качество PDF
                },
                PreserveOriginal = true              // Сохранить оригинал
            };

            Console.WriteLine("Шаг 5: Запрос на обработку создан");
            Console.WriteLine($"  Входной файл: {inputFile}");
            Console.WriteLine($"  Выходная папка: {outputDir}");
            Console.WriteLine($"  Экспорт в PDF: Да (высокое качество)");
            Console.WriteLine($"  Сохранить оригинал: Да\n");

            // ========================================================================
            // ШАГ 6: ВЫПОЛНЯЕМ ОБРАБОТКУ
            // ========================================================================

            Console.WriteLine("Шаг 6: Начало обработки документа...\n");
            Console.WriteLine("--- ПЕРВЫЙ ПРОХОД ---");

            using (var anonymizer = new DocumentAnonymizer(
                visible: false,      // Word/SolidWorks невидимы
                useOpenXml: true))  // Используем Interop для PDF
            {
                using (var processor = anonymizer.GetType()
                    .GetField("_factory", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                    ?.GetValue(anonymizer) as DocumentProcessingLibrary.Documents.Factories.DocumentProcessorFactory)
                {
                    var docProcessor = processor?.CreateProcessor(inputFile);

                    if (docProcessor is ITwoPassDocumentProcessor twoPassProcessor)
                    {
                        var result = twoPassProcessor.ProcessTwoPass(processingRequest, twoPassConfig);

                        // ========================================================================
                        // ШАГ 7: АНАЛИЗИРУЕМ РЕЗУЛЬТАТЫ
                        // ========================================================================

                        Console.WriteLine("\n--- РЕЗУЛЬТАТЫ ОБРАБОТКИ ---\n");

                        if (result.Success)
                        {
                            Console.WriteLine("✓ Обработка завершена УСПЕШНО!\n");

                            Console.WriteLine($"📊 Статистика:");
                            Console.WriteLine($"   Всего найдено совпадений: {result.MatchesFound}");
                            Console.WriteLine($"   Успешно обработано: {result.MatchesProcessed}");

                            // Получаем извлеченные коды
                            var extractedCodes = codeExtractionStrategy.GetExtractedCodes();

                            Console.WriteLine($"\n🔑 Извлеченные коды организаций ({extractedCodes.Count}):");
                            foreach (var code in extractedCodes)
                            {
                                Console.WriteLine($"   - {code}");
                            }

                            if (result.Metadata.ContainsKey("CodesRemoved"))
                            {
                                Console.WriteLine($"\n🗑️  Удалено отдельно стоящих кодов: {result.Metadata["CodesRemoved"]}");
                            }

                            Console.WriteLine($"\n📁 Результаты сохранены в: {outputDir}");
                        }
                        else
                        {
                            Console.WriteLine("✗ Обработка завершена С ОШИБКАМИ!\n");

                            Console.WriteLine("❌ Ошибки:");
                            foreach (var error in result.Errors)
                            {
                                Console.WriteLine($"   {error}");
                            }
                        }

                        // Предупреждения (если есть)
                        if (result.Warnings.Any())
                        {
                            Console.WriteLine("\n⚠️  Предупреждения:");
                            foreach (var warning in result.Warnings)
                            {
                                Console.WriteLine($"   {warning}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("⚠️  Процессор не поддерживает двухпроходную обработку");
                        Console.WriteLine("   Используется обычная обработка...");

                        var result = docProcessor?.Process(processingRequest);

                        if (result != null && result.Success)
                        {
                            Console.WriteLine($"\n✓ Обработано: {result.MatchesProcessed} совпадений");
                        }
                    }
                }
            }

            Console.WriteLine("\n=== ОБРАБОТКА ЗАВЕРШЕНА ===");
        }

        /// <summary>
        /// УПРОЩЕННЫЙ ПРИМЕР: Через готовый метод фасада
        /// </summary>
        public static void SimplifiedExample()
        {
            Console.WriteLine("=== УПРОЩЕННЫЙ ВАРИАНТ (через фасад) ===\n");

            using (var anonymizer = new DocumentAnonymizer(
                visible: false,
                useOpenXml: false))
            {
                // Все настроено автоматически!
                var result = anonymizer.AnonymizeDocumentWithCodeRemoval(
                    inputFilePath: @"C:\Documents\Test.docx",
                    outputDirectory: @"C:\Output"
                );

                if (result.Success)
                {
                    Console.WriteLine($"✓ Успех! Обработано: {result.MatchesProcessed}");

                    if (result.Metadata.ContainsKey("CodesRemoved"))
                    {
                        Console.WriteLine($"  Удалено кодов: {result.Metadata["CodesRemoved"]}");
                    }
                }
            }
        }

        /// <summary>
        /// ПРИМЕР С КАСТОМНЫМИ ПАТТЕРНАМИ
        /// </summary>
        public static void CustomPatternsExample()
        {
            Console.WriteLine("=== КАСТОМНЫЕ ПАТТЕРНЫ ===\n");

            // Создаем свои паттерны для обозначений
            var customDesignations = new RegexSearchStrategy(
                "MyDesignations",
                new RegexPattern("Format1", @"[A-Z]{2,4}\.[0-9]{3,5}"),
                new RegexPattern("Format2", @"[A-Z]+-[0-9]+\.[0-9]+")
            );

            // Создаем свои паттерны для имен
            var customNames = new RegexSearchStrategy(
                "MyNames",
                new RegexPattern("FullName", @"[А-Я][а-я]+ [А-Я][а-я]+ [А-Я][а-я]+"),
                new RegexPattern("ShortName", @"[А-Я][а-я]+ [А-Я]\. [А-Я]\.")
            );

            var codeExtractor = new OrganizationCodeRemovalStrategy();

            var config = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>
                {
                    customDesignations,
                    customNames
                },
                ReplacementStrategy = codeExtractor,
                Options = new ProcessingOptions
                {
                    ProcessProperties = true,
                    ProcessTextBoxes = true,
                    ProcessNotes = false,
                    ProcessHeaders = true,
                    ProcessFooters = false,
                    MinMatchLength = 3,
                    CaseSensitive = false
                }
            };

            using (var anonymizer = new DocumentAnonymizer())
            {
                var result = anonymizer.AnonymizeDocument(
                    @"C:\Documents\Custom.docx",
                    @"C:\Output",
                    config
                );

                Console.WriteLine($"Результат: {(result.Success ? "Успех" : "Ошибка")}");
                Console.WriteLine($"Обработано: {result.MatchesProcessed}");
            }
        }
    }

    /// <summary>
    /// ТОЧКА ВХОДА
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {

                // 1. Полная ручная настройка (РЕКОМЕНДУЕТСЯ для понимания)
                ManualConfigurationExample.FullManualConfiguration();

                // 2. Упрощенный вариант
                // ManualConfigurationExample.SimplifiedExample();

                // 3. С кастомными паттернами
                // ManualConfigurationExample.CustomPatternsExample();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ ОШИБКА: {ex.Message}");
                Console.WriteLine($"\nStack trace:\n{ex.StackTrace}");
            }

            Console.WriteLine("\nНажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}

