using Simply.ClipboardMonitor.Common;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Reads clipboard data stored as an HENHMETAFILE handle (CF_ENHMETAFILE, CF_DSPENHMETAFILE).
/// Retrieves the raw EMF bytes via GetEnhMetaFileBits.
/// </summary>
internal sealed class HEnhMetaFileHandleReadStrategy : IHandleReadStrategy
{
    public string HandleType => HandleTypes.HEnhMetaFile;

    public bool TryRead(uint formatId, out byte[]? data, out string failureMessage)
    {
        data = null;
        var handle = NativeMethods.GetClipboardData(formatId);
        if (handle == IntPtr.Zero)
        {
            failureMessage = "No data handle is available for this format.";
            return false;
        }

        return TryReadEnhMetaFileAsBytes(handle, out data, out failureMessage);
    }

    private static bool TryReadEnhMetaFileAsBytes(IntPtr hEmf, out byte[]? bytes, out string failureMessage)
    {
        bytes = null;

        var size = NativeMethods.GetEnhMetaFileBits(hEmf, 0, null);
        if (size == 0)
        {
            failureMessage = "Failed to determine EMF data size.";
            return false;
        }

        bytes = new byte[size];
        if (NativeMethods.GetEnhMetaFileBits(hEmf, size, bytes) != size)
        {
            bytes          = null;
            failureMessage = "Failed to read EMF data.";
            return false;
        }

        failureMessage = string.Empty;
        return true;
    }
}
