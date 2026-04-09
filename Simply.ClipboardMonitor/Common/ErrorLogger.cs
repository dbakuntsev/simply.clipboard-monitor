using System.IO;
using System.Reflection;
using System.Text;

namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Thread-safe rolling error logger. Appends exceptions to a plain-text file named
/// <c>error_YYYY-MM-DD.txt</c> under <c>%LOCALAPPDATA%\Simply.ClipboardMonitor\</c>.
///
/// At most 3 log files are kept; the oldest are deleted after each write.
/// All methods are best-effort and never throw.
/// </summary>
internal static class ErrorLogger
{
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Simply.ClipboardMonitor");

    private static readonly string AppVersion =
        typeof(ErrorLogger).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion
        ?? typeof(ErrorLogger).Assembly.GetName().Version?.ToString()
        ?? "unknown";

    private static readonly object _lock = new();

    /// <summary>
    /// Full path of the log file written to most recently during this process lifetime,
    /// or <see langword="null"/> if nothing has been logged yet.
    /// </summary>
    public static string? CurrentLogFilePath { get; private set; }

    /// <summary>
    /// Appends <paramref name="exception"/> to today's log file and trims old files.
    /// Never throws.
    /// </summary>
    public static void Log(Exception exception)
    {
        try
        {
            var now         = DateTime.Now;
            var logFilePath = Path.Combine(LogDirectory, $"error_{now:yyyy-MM-dd}.txt");
            var entry       = FormatEntry(now, exception);

            lock (_lock)
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(logFilePath, entry, Encoding.UTF8);
                CurrentLogFilePath = logFilePath;
                PruneOldLogs();
            }
        }
        catch
        {
            // Best-effort — never propagate errors from the logger.
        }
    }

    // ── Formatting ───────────────────────────────────────────────────────────

    private static string FormatEntry(DateTime timestamp, Exception exception)
    {
        var sb = new StringBuilder();
        sb.Append(timestamp.ToString("O")); // ISO 8601 roundtrip with local time zone offset
        sb.Append(" | ");
        sb.Append(AppVersion);
        sb.Append(" | ");
        AppendException(sb, exception);
        sb.AppendLine();
        return sb.ToString();
    }

    private static void AppendException(StringBuilder sb, Exception exception)
    {
        sb.Append(exception.GetType().FullName);
        sb.Append(": ");
        sb.AppendLine(exception.Message);

        if (exception.StackTrace is { } st)
            sb.AppendLine(st);

        if (exception is AggregateException agg && agg.InnerExceptions.Count > 0)
        {
            for (var i = 0; i < agg.InnerExceptions.Count; i++)
            {
                sb.AppendLine($"--- Inner Exception [{i + 1}/{agg.InnerExceptions.Count}] ---");
                AppendException(sb, agg.InnerExceptions[i]);
            }
        }
        else if (exception.InnerException is { } inner)
        {
            sb.AppendLine("--- Inner Exception ---");
            AppendException(sb, inner);
        }
    }

    // ── Retention ────────────────────────────────────────────────────────────

    private static void PruneOldLogs()
    {
        // Called inside _lock — no additional synchronisation required.
        var toDelete = Directory.GetFiles(LogDirectory, "error_*.txt")
            .OrderDescending()
            .Skip(3);

        foreach (var file in toDelete)
        {
            try { File.Delete(file); }
            catch { /* best effort */ }
        }
    }
}
