using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Persistence contract for the clipboard change history database.</summary>
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
        int maxEntries,
        long maxDatabaseBytes);

    /// <summary>Returns the history database file size in bytes, or 0 if it does not exist.</summary>
    long GetDatabaseFileSize();

    /// <summary>
    /// Applies entry-count and size limits to an existing database, vacuuming if needed.
    /// Returns true if any sessions were deleted. No-op when the database does not exist.
    /// </summary>
    bool EnforceLimits(int maxEntries, long maxDatabaseBytes);

    /// <summary>
    /// Deletes all sessions, items, and content blobs, then compacts the file with VACUUM.
    /// Safe to call when the database does not yet exist.
    /// </summary>
    void ClearHistory();

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
