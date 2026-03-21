using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Restores clipboard formats from saved data back onto the Windows clipboard.
/// Implements <see cref="IClipboardWriter"/>; all Win32 write operations are
/// encapsulated here so that callers depend only on the interface.
/// Handle-type dispatch is delegated to injected <see cref="IHandleWriteStrategy"/> instances.
/// </summary>
internal sealed class ClipboardWriterService : IClipboardWriter
{
    private readonly IReadOnlyDictionary<string, IHandleWriteStrategy> _writeStrategies;

    public ClipboardWriterService(IEnumerable<IHandleWriteStrategy> handleWriteStrategies)
    {
        _writeStrategies = handleWriteStrategies.ToDictionary(s => s.HandleType);
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The clipboard must already be open (via <c>NativeMethods.OpenClipboard</c>) and empty
    /// (via <c>NativeMethods.EmptyClipboard</c>) before calling this method.
    /// Custom format names (IDs ≥ 0xC000) are re-registered so their IDs are valid in the
    /// current Windows session, which may differ from the session in which they were saved.
    /// </remarks>
    public void RestoreFormats(IReadOnlyList<SavedClipboardFormat> formats)
    {
        foreach (var fmt in formats)
            RestoreFormat(fmt);
    }

    // ── Private restore dispatch ────────────────────────────────────────────

    private void RestoreFormat(SavedClipboardFormat fmt)
    {
        // Custom format IDs (≥ 0xC000) are assigned dynamically per Windows session;
        // re-register by name to get the current-session ID.
        uint actualId = fmt.FormatId >= 0xC000
            ? NativeMethods.RegisterClipboardFormat(fmt.FormatName)
            : fmt.FormatId;

        if (actualId == 0)
            return;

        // "none" (e.g. CF_PALETTE): no bytes were captured; skip silently.
        if (_writeStrategies.TryGetValue(fmt.HandleType, out var strategy))
            strategy.Restore(actualId, fmt.Data);
    }
}
