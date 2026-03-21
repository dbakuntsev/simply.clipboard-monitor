namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Restores a single clipboard format using a specific handle type.
/// The clipboard must already be open and empty when <see cref="Restore"/> is called.
/// </summary>
internal interface IHandleWriteStrategy
{
    /// <summary>The handle type this strategy handles (e.g. "hglobal", "hbitmap").</summary>
    string HandleType { get; }

    /// <summary>
    /// Restores <paramref name="formatId"/> to the open clipboard using <paramref name="data"/>.
    /// A null or empty <paramref name="data"/> is silently ignored.
    /// </summary>
    void Restore(uint formatId, byte[]? data);
}
