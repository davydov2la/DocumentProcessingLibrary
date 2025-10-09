using Microsoft.Extensions.Logging;
using DocumentProcessingLibrary.Processing.Models;

namespace DocumentProcessingLibrary.Logging;

/// <summary>
/// Обёртка вокруг Microsoft ILogger, которая дублирует сообщения Warning/Error в ProcessingResult (если передан).
/// </summary>
public sealed class ResultAwareLogger : ILogger
{
    private readonly ILogger _inner;
    private readonly ProcessingResult? _result;

    public ResultAwareLogger(ILogger inner, ProcessingResult? result = null)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _result = result;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        => _inner.BeginScope(state);

    public bool IsEnabled(LogLevel logLevel) => _inner.IsEnabled(logLevel);

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        if (formatter == null) return;

        var message = formatter(state, exception);

        // Дублируем в ProcessingResult при необходимости — защищённо
        if (_result != null && !string.IsNullOrWhiteSpace(message))
        {
            try
            {
                switch (logLevel)
                {
                    case LogLevel.Warning:
                        // УЛУЧШЕНИЕ: Проверяем на дубликаты
                        if (!_result.Warnings.Contains(message))
                            _result.Warnings.Add(message);
                        break;

                    case LogLevel.Error:
                    case LogLevel.Critical:
                        // УЛУЧШЕНИЕ: Добавляем информацию об исключении
                        var errorMessage = exception != null
                            ? $"{message} | Exception: {exception.GetType().Name}: {exception.Message}"
                            : message;

                        if (!_result.Errors.Contains(errorMessage))
                            _result.Errors.Add(errorMessage);
                        break;
                }
            }
            catch
            {
                // Не допускаем, чтобы логика логгирования ломала основной поток
            }
        }

        _inner.Log(logLevel, eventId, state, exception, formatter);
    }
}