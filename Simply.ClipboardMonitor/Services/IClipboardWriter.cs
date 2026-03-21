using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Restores clipboard formats from saved data back onto the Windows clipboard.
/// </summary>
public interface IClipboardWriter
{
    /// <summary>
    /// Empties the clipboard and writes every format in <paramref name="formats"/> to it.
    /// Custom format names are re-registered so IDs survive across sessions.
    /// </summary>
    void RestoreFormats(IReadOnlyList<SavedClipboardFormat> formats);
}
