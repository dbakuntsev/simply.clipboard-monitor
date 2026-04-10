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

    /// <summary>
    /// Runs <c>PRAGMA integrity_check(1)</c> on the history database and returns the result.
    /// Returns <see cref="DatabaseIntegrityStatus.Absent"/> when the file does not exist.
    /// Returns <see cref="DatabaseIntegrityStatus.Corrupted"/> if the file cannot be opened.
    /// Bails after the first error, so large healthy databases are not penalised.
    /// Safe to call from a background thread.
    /// </summary>
    DatabaseIntegrityStatus CheckIntegrity();

    /// <summary>
    /// Attempts to recover a corrupt history database, trying three strategies in order:
    /// <list type="number">
    ///   <item><description><c>VACUUM INTO</c> a temp file (clean rebuild by SQLite).</description></item>
    ///   <item><description>Table-by-table bulk copy into a fresh database.</description></item>
    ///   <item><description>Row-by-row rescue for any table whose bulk copy fails.</description></item>
    /// </list>
    /// On success the recovered file replaces the original. The corrupt original is deleted.
    /// Safe to call from a background thread.
    /// </summary>
    RecoveryResult TryRecover();

    /// <summary>
    /// Deletes the history database file entirely.
    /// No-op when the file does not exist.
    /// Safe to call from a background thread.
    /// </summary>
    void DeleteDatabase();

    /// <summary>
    /// Creates a new empty history database at the standard path and applies the current schema.
    /// Call after <see cref="DeleteDatabase"/> to establish a fresh store.
    /// Safe to call from a background thread.
    /// </summary>
    void InitializeFreshDatabase();
}
