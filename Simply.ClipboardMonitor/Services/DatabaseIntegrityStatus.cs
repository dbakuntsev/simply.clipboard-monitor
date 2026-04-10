namespace Simply.ClipboardMonitor.Services;

/// <summary>Result of a <c>PRAGMA integrity_check</c> run on the history database.</summary>
public enum DatabaseIntegrityStatus
{
    /// <summary>The database file does not exist.</summary>
    Absent,

    /// <summary><c>PRAGMA integrity_check(1)</c> returned <c>"ok"</c>.</summary>
    Healthy,

    /// <summary>
    /// <c>PRAGMA integrity_check(1)</c> reported at least one error,
    /// or the database file could not be opened at all.
    /// </summary>
    Corrupted,
}
