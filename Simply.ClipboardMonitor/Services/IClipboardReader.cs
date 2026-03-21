using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Reads the current state of the Windows clipboard without any UI coupling.
/// </summary>
public interface IClipboardReader
{
    /// <summary>Enumerates all formats currently available on the clipboard.</summary>
    IReadOnlyList<ClipboardFormatItem> EnumerateFormats();

    /// <summary>
    /// Opens the clipboard, reads raw bytes for every format in <paramref name="formats"/>,
    /// then closes the clipboard.  Returns an empty list if the clipboard cannot be opened.
    /// Each <see cref="FormatSnapshot"/> carries the handle type, the raw bytes (null for
    /// "none"-type handles), and the original byte size.
    /// </summary>
    IReadOnlyList<FormatSnapshot> CaptureAllFormats(IReadOnlyList<ClipboardFormatItem> formats);

    /// <summary>
    /// Reads the raw bytes for a single format.
    /// Returns false (with a non-null <paramref name="failureMessage"/>) when the data
    /// cannot be read; <paramref name="data"/> may still be null for "none" handle types.
    /// </summary>
    bool TryReadFormatBytes(uint formatId, string handleType,
        out byte[]? data, out string failureMessage);

    /// <summary>Resolves the display name for a clipboard format ID.</summary>
    string GetFormatDisplayName(uint formatId);

    /// <summary>
    /// Returns the handle type for a format ID:
    /// "hglobal" | "hbitmap" | "henhmetafile" | "none"
    /// </summary>
    string GetHandleType(uint formatId);

    /// <summary>Returns the clipboard sequence number (for delayed-render detection).</summary>
    uint GetSequenceNumber();

    /// <summary>
    /// Attempts to open the clipboard, retrying up to 10 times with 20 ms delays.
    /// Returns true on success.
    /// </summary>
    bool TryOpenClipboard(IntPtr hwnd);

    /// <summary>Closes the clipboard previously opened by <see cref="TryOpenClipboard"/>.</summary>
    void CloseClipboard();
}
