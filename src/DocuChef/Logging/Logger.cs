using System.Diagnostics;

namespace DocuChef.Logging;

/// <summary>
/// Provides logging functionality for DocuChef with configurable log levels
/// </summary>
internal static class Logger
{
    /// <summary>
    /// Log levels supported by the logger
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    private static LogLevel _minimumLevel = LogLevel.Warning;
    private static Action<string, LogLevel> _logAction = DefaultLogAction;

    /// <summary>
    /// Gets or sets the minimum log level
    /// </summary>
    public static LogLevel MinimumLevel
    {
        get => _minimumLevel;
        set => _minimumLevel = value;
    }

    /// <summary>
    /// Sets a custom log handler
    /// </summary>
    public static void SetLogHandler(Action<string, LogLevel> logHandler)
    {
        _logAction = logHandler ?? DefaultLogAction;
    }

    /// <summary>
    /// Logs a debug message
    /// </summary>
    public static void Debug(string message)
    {
        if (_minimumLevel <= LogLevel.Debug)
            _logAction(message, LogLevel.Debug);
    }

    /// <summary>
    /// Logs an info message
    /// </summary>
    public static void Info(string message)
    {
        if (_minimumLevel <= LogLevel.Info)
            _logAction(message, LogLevel.Info);
    }

    /// <summary>
    /// Logs a warning message
    /// </summary>
    public static void Warning(string message)
    {
        if (_minimumLevel <= LogLevel.Warning)
            _logAction(message, LogLevel.Warning);
    }

    /// <summary>
    /// Logs an error message
    /// </summary>
    public static void Error(string message, Exception exception = null)
    {
        if (_minimumLevel <= LogLevel.Error)
        {
            string fullMessage = message;
            if (exception != null)
                fullMessage += $" Exception: {exception.Message}";

            _logAction(fullMessage, LogLevel.Error);
        }
    }

    /// <summary>
    /// Default logging implementation
    /// </summary>
    private static void DefaultLogAction(string message, LogLevel level)
    {
        string prefix = $"[DocuChef:{level}] ";
        System.Diagnostics.Debug.WriteLine($"{prefix}{message}");

        // For Error level, also write to trace
        if (level == LogLevel.Error)
            Trace.TraceError($"{prefix}{message}");
    }
}