namespace DocuChef.Utils;

public delegate void LogCallback(LogLevel level, string message, Exception? exception = null);

public enum LogLevel
{
    Debug,
    Information,
    Warning,
    Error,
    Critical
}

internal static class LoggingHelper
{
    private static LogCallback? _logCallback;

    public static void SetLogCallback(LogCallback? callback)
    {
        _logCallback = callback;
    }

    public static void LogDebug(string message, Exception? exception = null)
    {
        Log(LogLevel.Debug, message, exception);
    }

    public static void LogInformation(string message, Exception? exception = null)
    {
        Log(LogLevel.Information, message, exception);
    }

    public static void LogWarning(string message, Exception? exception = null)
    {
        Log(LogLevel.Warning, message, exception);
    }

    public static void LogError(string message, Exception? exception = null)
    {
        Log(LogLevel.Error, message, exception);
    }

    private static void Log(LogLevel level, string message, Exception? exception = null)
    {
        _logCallback?.Invoke(level, message, exception);

        // Fall back to console if no callback
        if (_logCallback == null)
        {
            Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}");
            if (exception != null)
            {
                Console.WriteLine($"Exception: {exception.Message}");
            }
        }
    }
}