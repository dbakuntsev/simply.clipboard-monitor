namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Result of a database recovery attempt performed by
/// <see cref="IHistoryMaintenance.TryRecover"/>.
/// </summary>
/// <param name="Success">
/// <see langword="true"/> when the database was recovered and replaced successfully;
/// <see langword="false"/> when all recovery strategies failed.
/// </param>
/// <param name="HadUnreadableRows">
/// <see langword="true"/> when at least one row could not be read from the corrupt source
/// and was therefore omitted from the recovered database.
/// Meaningful only when <see cref="Success"/> is <see langword="true"/>.
/// </param>
/// <param name="Strategy">
/// Human-readable name of the recovery strategy that succeeded (e.g. "VACUUM INTO" or
/// "manual copy"). Empty string when <see cref="Success"/> is <see langword="false"/>.
/// </param>
/// <param name="SessionsRecovered">
/// Number of session rows present in the recovered database.
/// Meaningful only when <see cref="Success"/> is <see langword="true"/>.
/// </param>
/// <param name="SessionsLost">
/// Estimated number of session rows that could not be salvaged.
/// 0 when the original count was unreadable or no rows were lost.
/// Meaningful only when <see cref="Success"/> is <see langword="true"/>.
/// </param>
public readonly record struct RecoveryResult(
    bool Success,
    bool HadUnreadableRows,
    string Strategy = "",
    int SessionsRecovered = 0,
    int SessionsLost = 0);
