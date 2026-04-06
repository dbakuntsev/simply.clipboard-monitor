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
    /// <para>
    /// <paramref name="textContents"/> is a parallel list to <paramref name="snapshots"/>:
    /// each element is the full decoded text for the corresponding format, or <see langword="null"/>
    /// if the format is not text-compatible. Stored in the DB for full-text history search.
    /// </para>
    /// <para><paramref name="pillsText"/> is the space-separated pill labels for the session
    /// (e.g. <c>"IMG TXT HTML"</c>), used for pill-label history search.</para>
    /// </summary>
    (long SessionId, bool Trimmed) AddSession(
        IReadOnlyList<FormatSnapshot> snapshots,
        IReadOnlyList<string?>        textContents,
        string                        pillsText,
        DateTime timestamp,
        int  maxEntries,
        long maxDatabaseBytes);

    /// <summary>
    /// Returns sessions ordered newest-first (up to <paramref name="maxCount"/>).
    /// When <paramref name="filter"/> is non-empty, only sessions whose timestamp,
    /// pill labels, format names, or text content contain the term (case-insensitive)
    /// are returned. Returns an empty list if the database does not exist yet.
    /// </summary>
    List<SessionEntry> LoadSessions(string? filter = null, int maxCount = 2000);

    /// <summary>
    /// Returns all format snapshots for a given session, with blobs decompressed.
    /// Returns an empty list if the database does not exist.
    /// </summary>
    List<FormatSnapshot> LoadSessionFormats(long sessionId);

    /// <summary>Builds a comma-separated summary of format names for display in the history list.</summary>
    string BuildFormatsText(IReadOnlyList<FormatSnapshot> snapshots);

    /// <summary>
    /// Returns the total number of sessions in the database, ignoring any filter.
    /// Returns 0 if the database does not exist.
    /// </summary>
    int GetSessionCount();

    /// <summary>
    /// Deletes the session with the given ID, its items, and any content blobs that are
    /// no longer referenced by any other session.
    /// No-op when the database does not exist.
    /// </summary>
    void DeleteSession(long sessionId);

    /// <summary>
    /// Returns <see langword="true"/> when the most recent history session has exactly
    /// the same ordered format set and the same per-format content hashes as
    /// <paramref name="snapshots"/>, meaning the current clipboard is a duplicate of
    /// what was last recorded.
    /// Returns <see langword="false"/> when the database does not exist, contains no
    /// sessions, or the content differs in any way.
    /// </summary>
    bool IsDuplicateOfLastSession(IReadOnlyList<FormatSnapshot> snapshots);
}
