namespace Simply.ClipboardMonitor.Models;

/// <summary>
/// One clipboard format entry captured at a point in time, with its raw bytes.
/// </summary>
public sealed record FormatSnapshot(
    int     Ordinal,
    uint    FormatId,
    string  FormatName,
    string  HandleType,
    /// <summary>Raw (uncompressed) bytes; null for handle types with no readable data.</summary>
    byte[]? Data,
    long    OriginalSize);
