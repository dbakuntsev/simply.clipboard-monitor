namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Shared display formatting utilities used across views and services.
/// </summary>
internal static class DisplayHelper
{
    /// <summary>
    /// Formats a byte count as a human-readable size string.
    /// Returns "Not created yet" for zero (covers the case where a database file has not been created).
    /// </summary>
    internal static string FormatFileSize(long bytes) => bytes switch
    {
        0                      => "Not created yet",
        >= 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024 * 1024):F1} GB",
        >= 1024L * 1024        => $"{bytes / (1024.0 * 1024):F1} MB",
        >= 1024                => $"{bytes / 1024.0:F1} KB",
        _                      => $"{bytes} B",
    };
}
