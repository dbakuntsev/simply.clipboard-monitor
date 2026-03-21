namespace Simply.ClipboardMonitor.Models;

/// <summary>
/// A single row from the sessions table, used to populate the history list.
/// </summary>
public sealed record SessionEntry(
    long     SessionId,
    DateTime Timestamp,
    string   FormatsText,
    long     TotalSize,
    IReadOnlyList<(uint FormatId, string FormatName)> Formats);
