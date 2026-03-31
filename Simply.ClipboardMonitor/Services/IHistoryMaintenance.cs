namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Database maintenance operations for the clipboard history store.
/// Separated from <see cref="IHistoryRepository"/> so that consumers such as
/// <c>SettingsDialog</c> that only need maintenance operations do not take a
/// dependency on the full session-persistence contract.
/// </summary>
public interface IHistoryMaintenance
{
    /// <summary>Returns the history database file size in bytes, or 0 if it does not exist.</summary>
    long GetDatabaseFileSize();

    /// <summary>
    /// Applies entry-count and size limits to an existing database, vacuuming if needed.
    /// Returns true if any sessions were deleted. No-op when the database does not exist.
    /// </summary>
    bool EnforceLimits(int maxEntries, long maxDatabaseBytes);

    /// <summary>
    /// Applies any pending schema migrations (e.g. adding new columns) to an existing
    /// database so that search features work immediately on startup.
    /// No-op when the database does not yet exist.
    /// Safe to call from a background thread.
    /// </summary>
    void MigrateSchema();

    /// <summary>
    /// Deletes all sessions, items, and content blobs, then compacts the file with VACUUM.
    /// Safe to call when the database does not yet exist.
    /// </summary>
    void ClearHistory();
}
