using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Read/write persistence contract for the clipboard change history database.
/// Maintenance operations (size limits, clear) live in the narrower
/// <see cref="IHistoryMaintenance"/> interface so consumers that only need
/// session data do not depend on administrative behaviour.
/// </summary>
public interface IHistoryRepository
{
    /// <summary>
    /// Inserts a new session and all its format snapshots, then trims the oldest sessions
    /// until both the entry count and approximate database size are within their limits.
    /// Returns the new session's row ID and a flag indicating whether any older sessions
    /// were deleted (i.e. the history list needs a full reload).
    /// </summary>
    (long SessionId, bool Trimmed) AddSession(
        IReadOnlyList<FormatSnapshot> snapshots,
        DateTime timestamp,
        int  maxEntries,
        long maxDatabaseBytes);

    /// <summary>
    /// Returns all sessions ordered newest-first (up to <paramref name="maxCount"/>).
    /// Returns an empty list if the database does not exist yet.
    /// </summary>
    List<SessionEntry> LoadSessions(int maxCount = 2000);

    /// <summary>
    /// Returns all format snapshots for a given session, with blobs decompressed.
    /// Returns an empty list if the database does not exist.
    /// </summary>
    List<FormatSnapshot> LoadSessionFormats(long sessionId);

    /// <summary>Builds a comma-separated summary of format names for display in the history list.</summary>
    string BuildFormatsText(IReadOnlyList<FormatSnapshot> snapshots);
}
