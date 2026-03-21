namespace Simply.ClipboardMonitor.Models;

/// <summary>
/// Represents one clipboard format entry as stored in / loaded from a .clipdb file.
/// </summary>
public sealed record SavedClipboardFormat(
    int    Ordinal,
    uint   FormatId,
    string FormatName,
    /// <summary>
    /// How the Windows clipboard handle was originally obtained.
    /// One of: "hglobal" | "hbitmap" | "henhmetafile" | "none"
    /// </summary>
    string HandleType,
    /// <summary>Raw bytes; null for formats with no readable data (e.g. CF_PALETTE).</summary>
    byte[]? Data);
