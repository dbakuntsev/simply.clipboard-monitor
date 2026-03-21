using Simply.ClipboardMonitor.Common;
using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Reads clipboard data stored as an HGLOBAL memory block (the most common handle type).
/// </summary>
internal sealed class HGlobalHandleReadStrategy : IHandleReadStrategy
{
    private const int ERROR_NOT_LOCKED = 158;

    public string HandleType => "hglobal";

    public bool TryRead(uint formatId, out byte[]? data, out string failureMessage)
    {
        data = null;

        var handle = NativeMethods.GetClipboardData(formatId);
        if (handle == IntPtr.Zero)
        {
            failureMessage = "No data handle is available for this format.";
            return false;
        }

        var globalSize = NativeMethods.GlobalSize(handle);
        if (globalSize == UIntPtr.Zero)
        {
            failureMessage = "Clipboard format size is unavailable.";
            return false;
        }

        var size64 = (ulong)globalSize;
        if (size64 > int.MaxValue)
        {
            failureMessage = "Clipboard data is too large to render in this viewer.";
            return false;
        }

        var dataPtr = NativeMethods.GlobalLock(handle);
        if (dataPtr == IntPtr.Zero)
        {
            failureMessage = "Failed to lock clipboard data.";
            return false;
        }

        try
        {
            data = new byte[(int)size64];
            Marshal.Copy(dataPtr, data, 0, data.Length);
            failureMessage = string.Empty;
            return true;
        }
        finally
        {
            if (!NativeMethods.GlobalUnlock(handle))
            {
                var error = Marshal.GetLastWin32Error();
                if (error != 0 && error != ERROR_NOT_LOCKED)
                {
                    // Best-effort unlock on clipboard-provided memory; ignore residual errors.
                }
            }
        }
    }
}
