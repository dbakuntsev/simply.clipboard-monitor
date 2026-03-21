using Simply.ClipboardMonitor.Common;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Restores a clipboard format whose data is stored as an HENHMETAFILE handle.
/// </summary>
internal sealed class HEnhMetaFileHandleWriteStrategy : IHandleWriteStrategy
{
    public string HandleType => "henhmetafile";

    public void Restore(uint formatId, byte[]? data)
    {
        if (data is not { Length: > 0 })
            return;

        var hemf = NativeMethods.SetEnhMetaFileBits((uint)data.Length, data);
        if (hemf == IntPtr.Zero)
            return;

        if (NativeMethods.SetClipboardData(formatId, hemf) == IntPtr.Zero)
            NativeMethods.DeleteEnhMetaFile(hemf);
    }
}
