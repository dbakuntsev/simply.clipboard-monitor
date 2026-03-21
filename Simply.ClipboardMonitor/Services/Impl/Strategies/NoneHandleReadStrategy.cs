namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Handles the "none" handle type (e.g. CF_PALETTE): no data can be read as raw bytes.
/// </summary>
internal sealed class NoneHandleReadStrategy : IHandleReadStrategy
{
    public string HandleType => "none";

    public bool TryRead(uint formatId, out byte[]? data, out string failureMessage)
    {
        data           = null;
        failureMessage = "Selected format uses a non-memory handle type (not HGLOBAL).";
        return false;
    }
}
