using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Restores clipboard formats from saved data back onto the Windows clipboard.
/// Implements <see cref="IClipboardWriter"/>; all Win32 write operations are
/// encapsulated here so that callers depend only on the interface.
/// </summary>
internal sealed class ClipboardWriterService : IClipboardWriter
{
    /// <inheritdoc/>
    /// <remarks>
    /// The clipboard must already be open (via <c>NativeMethods.OpenClipboard</c>) and empty
    /// (via <c>NativeMethods.EmptyClipboard</c>) before calling this method.
    /// Custom format names (IDs ≥ 0xC000) are re-registered so their IDs are valid in the
    /// current Windows session, which may differ from the session in which they were saved.
    /// </remarks>
    public void RestoreFormats(IReadOnlyList<SavedClipboardFormat> formats)
    {
        foreach (var fmt in formats)
            RestoreFormat(fmt);
    }

    // ── Private restore dispatch ────────────────────────────────────────────

    private static void RestoreFormat(SavedClipboardFormat fmt)
    {
        // Custom format IDs (≥ 0xC000) are assigned dynamically per Windows session;
        // re-register by name to get the current-session ID.
        uint actualId = fmt.FormatId >= 0xC000
            ? NativeMethods.RegisterClipboardFormat(fmt.FormatName)
            : fmt.FormatId;

        if (actualId == 0)
            return;

        switch (fmt.HandleType)
        {
            case "hglobal":
                RestoreAsHGlobal(actualId, fmt.Data);
                break;
            case "hbitmap":
                RestoreAsHBitmap(actualId, fmt.Data);
                break;
            case "henhmetafile":
                RestoreAsHEnhMetaFile(actualId, fmt.Data);
                break;
            // "none" (e.g. CF_PALETTE): no bytes were captured; skip silently.
        }
    }

    // ── HGLOBAL restore ─────────────────────────────────────────────────────

    private static void RestoreAsHGlobal(uint formatId, byte[]? data)
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

    // ── HENHMETAFILE restore ────────────────────────────────────────────────

    private static void RestoreAsHEnhMetaFile(uint formatId, byte[]? data)
    {
        if (data is not { Length: > 0 })
            return;

        var hemf = NativeMethods.SetEnhMetaFileBits((uint)data.Length, data);
        if (hemf == IntPtr.Zero)
            return;

        if (NativeMethods.SetClipboardData(formatId, hemf) == IntPtr.Zero)
            NativeMethods.DeleteEnhMetaFile(hemf);
    }

    // ── HBITMAP restore ─────────────────────────────────────────────────────

    private static void RestoreAsHBitmap(uint formatId, byte[]? data)
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
