using Simply.ClipboardMonitor.Common;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;
using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Restores a clipboard format whose data is stored as an HBITMAP handle.
/// The stored bytes are expected to be a BITMAPINFOHEADER followed by 32 bpp BI_RGB pixel data,
/// as produced by <see cref="HBitmapHandleReadStrategy"/>.
/// </summary>
internal sealed class HBitmapHandleWriteStrategy : IHandleWriteStrategy
{
    public string HandleType => HandleTypes.HBitmap;

    public void Restore(uint formatId, byte[]? data)
    {
        var headerSize = Marshal.SizeOf<BITMAPINFOHEADER>();
        if (data == null || data.Length <= headerSize)
            return;

        // The stored block is BITMAPINFOHEADER + pixel data (produced by GetDIBits, 32 bpp BI_RGB).
        var header    = MemoryMarshal.Read<BITMAPINFOHEADER>(data.AsSpan(0, headerSize));
        var pixelData = data.AsSpan(headerSize).ToArray();

        var hdc = NativeMethods.GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero)
            return;

        try
        {
            const uint CBM_INIT       = 4;
            const uint DIB_RGB_COLORS = 0;
            var hBitmap = NativeMethods.CreateDIBitmap(hdc, ref header, CBM_INIT,
                                                        pixelData, ref header, DIB_RGB_COLORS);
            if (hBitmap == IntPtr.Zero)
                return;

            if (NativeMethods.SetClipboardData(formatId, hBitmap) == IntPtr.Zero)
                NativeMethods.DeleteObject(hBitmap);
        }
        finally
        {
            NativeMethods.ReleaseDC(IntPtr.Zero, hdc);
        }
    }
}
