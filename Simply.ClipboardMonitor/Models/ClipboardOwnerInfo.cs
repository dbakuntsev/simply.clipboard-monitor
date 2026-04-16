namespace Simply.ClipboardMonitor.Models;

/// <summary>Resolved clipboard owner information for status-bar display.</summary>
/// <param name="DisplayText">
/// Short text shown in the status bar, e.g. <c>"chrome.exe (1234)"</c>,
/// <c>"unknown process (1234)"</c>, or <c>"(owner unknown)"</c>.
/// </param>
/// <param name="TooltipText">
/// Multi-line tooltip with PID, full path, and command line details,
/// or <see langword="null"/> when no details are available.
/// </param>
public sealed record ClipboardOwnerInfo(string DisplayText, string? TooltipText);
