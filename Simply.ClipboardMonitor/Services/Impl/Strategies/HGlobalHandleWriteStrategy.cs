using Simply.ClipboardMonitor.Common;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;
using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Restores a clipboard format whose data is stored as an HGLOBAL memory block.
/// </summary>
internal sealed class HGlobalHandleWriteStrategy : IHandleWriteStrategy
{
    public string HandleType => HandleTypes.HGlobal;

    public void Restore(uint formatId, byte[]? data)
    {
        if (data is not { Length: > 0 })
            return;

        const uint GMEM_MOVEABLE = 0x0002;
        var hGlobal = NativeMethods.GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)(uint)data.Length);
        if (hGlobal == IntPtr.Zero)
            return;

        var ptr = NativeMethods.GlobalLock(hGlobal);
        if (ptr == IntPtr.Zero)
        {
            NativeMethods.GlobalFree(hGlobal);
            return;
        }

        try   { Marshal.Copy(data, 0, ptr, data.Length); }
        finally { NativeMethods.GlobalUnlock(hGlobal); }

        if (NativeMethods.SetClipboardData(formatId, hGlobal) == IntPtr.Zero)
            NativeMethods.GlobalFree(hGlobal);
    }
}
