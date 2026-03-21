namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Reads raw bytes from a clipboard handle of a specific type.
/// The clipboard must already be open when <see cref="TryRead"/> is called.
/// </summary>
internal interface IHandleReadStrategy
{
    /// <summary>The handle type this strategy handles (e.g. "hglobal", "hbitmap").</summary>
    string HandleType { get; }

    /// <summary>
    /// Reads the data for <paramref name="formatId"/>.
    /// Returns true and sets <paramref name="data"/> on success;
    /// returns false with a non-empty <paramref name="failureMessage"/> on failure.
    /// </summary>
    bool TryRead(uint formatId, out byte[]? data, out string failureMessage);
}
