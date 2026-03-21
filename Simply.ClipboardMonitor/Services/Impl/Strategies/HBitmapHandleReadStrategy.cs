using Simply.ClipboardMonitor.Common;
using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Reads clipboard data stored as an HBITMAP handle (CF_BITMAP, CF_DSPBITMAP).
/// Converts the bitmap to a CF_DIB-compatible byte block via GetDIBits.
/// </summary>
internal sealed class HBitmapHandleReadStrategy : IHandleReadStrategy
{
    public string HandleType => "hbitmap";

    public bool TryRead(uint formatId, out byte[]? data, out string failureMessage)
    {
        data = null;
        var handle = NativeMethods.GetClipboardData(formatId);
        if (handle == IntPtr.Zero)
        {
            failureMessage = "No data handle is available for this format.";
            return false;
        }

        return TryReadHBitmapAsBytes(handle, out data, out failureMessage);
    }

    private static bool TryReadHBitmapAsBytes(IntPtr hBitmap, out byte[]? bytes, out string failureMessage)
    {
        bytes = null;

        if (NativeMethods.GetObject(hBitmap, Marshal.SizeOf<BITMAP>(), out var bm) == 0)
        {
            failureMessage = "Failed to read bitmap metadata.";
            return false;
        }

        // Request 32 bpp BI_RGB output so there is no colour table to worry about.
        var header = new BITMAPINFOHEADER
        {
            biSize        = (uint)Marshal.SizeOf<BITMAPINFOHEADER>(),
            biWidth       = bm.bmWidth,
            biHeight      = bm.bmHeight,  // positive → bottom-up DIB
            biPlanes      = 1,
            biBitCount    = 32,
            biCompression = 0,            // BI_RGB
        };

        var hdc = NativeMethods.GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero)
        {
            failureMessage = "Failed to acquire a device context.";
            return false;
        }

        try
        {
            var height = (uint)Math.Abs(bm.bmHeight);

            // First call: let GDI fill in biSizeImage.
            NativeMethods.GetDIBits(hdc, hBitmap, 0, height, null, ref header, 0 /* DIB_RGB_COLORS */);

            if (header.biSizeImage == 0)
            {
                var stride = (bm.bmWidth * 32 + 31) / 32 * 4;
                header.biSizeImage = (uint)(stride * height);
            }

            var pixelData = new byte[header.biSizeImage];
            if (NativeMethods.GetDIBits(hdc, hBitmap, 0, height, pixelData, ref header, 0) == 0)
            {
                failureMessage = "Failed to retrieve bitmap pixel data.";
                return false;
            }

            // Assemble a CF_DIB-compatible block: BITMAPINFOHEADER immediately followed by pixels.
            var headerSize = Marshal.SizeOf<BITMAPINFOHEADER>();
            bytes = new byte[headerSize + pixelData.Length];

            var pin = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try   { Marshal.StructureToPtr(header, pin.AddrOfPinnedObject(), fDeleteOld: false); }
            finally { pin.Free(); }

            Buffer.BlockCopy(pixelData, 0, bytes, headerSize, pixelData.Length);
            failureMessage = string.Empty;
            return true;
        }
        finally
        {
            NativeMethods.ReleaseDC(IntPtr.Zero, hdc);
        }
    }
}
