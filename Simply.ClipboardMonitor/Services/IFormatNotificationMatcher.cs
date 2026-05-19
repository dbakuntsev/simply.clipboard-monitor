namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Matches clipboard format names against a user-defined watch list so the
/// application can surface a system-tray notification when a watched format
/// appears on the clipboard.
/// </summary>
public interface IFormatNotificationMatcher
{
    /// <summary>Whether the feature is currently enabled by the user.</summary>
    bool Enabled { get; }

    /// <summary>
    /// True when the feature is enabled and at least one valid pattern is configured.
    /// Callers can short-circuit work (e.g. enumerating clipboard names) when this is false.
    /// </summary>
    bool HasActivePatterns { get; }

    /// <summary>
    /// Replaces the current configuration.  Patterns are one-per-line in
    /// <paramref name="patternsText"/>; <c>*</c> is treated as a wildcard and
    /// matching is case-insensitive against the whole format name.
    /// </summary>
    void Configure(bool enabled, string? patternsText);

    /// <summary>
    /// Returns the distinct format names that match at least one configured pattern.
    /// Order matches the input enumeration; an empty list is returned when matching
    /// is disabled or no name matches.
    /// </summary>
    IReadOnlyList<string> Match(IEnumerable<string> formatNames);
}
