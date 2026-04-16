using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Resolves information about the process that currently owns the clipboard.</summary>
public interface IClipboardOwnerService
{
    /// <summary>
    /// Resolves the current clipboard owner eagerly (path and command line are read at call time).
    /// Returns <see langword="null"/> when the clipboard is empty.
    /// </summary>
    ClipboardOwnerInfo? Resolve();
}
