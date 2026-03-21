using System.Runtime.InteropServices;
using System.Text;

namespace Simply.ClipboardMonitor.Common;

[StructLayout(LayoutKind.Sequential)]
internal struct BITMAP
{
    public int    bmType;
    public int    bmWidth;
    public int    bmHeight;
    public int    bmWidthBytes;
    public short  bmPlanes;
    public short  bmBitsPixel;
    public IntPtr bmBits;
}

[StructLayout(LayoutKind.Sequential)]
internal struct BITMAPINFOHEADER
{
    public uint  biSize;
    public int   biWidth;
    public int   biHeight;
    public short biPlanes;
    public short biBitCount;
    public uint  biCompression;
    public uint  biSizeImage;
    public int   biXPelsPerMeter;
    public int   biYPelsPerMeter;
    public uint  biClrUsed;
    public uint  biClrImportant;
}

internal static partial class NativeMethods
{
    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool AddClipboardFormatListener(IntPtr hwnd);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool RemoveClipboardFormatListener(IntPtr hwnd);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool OpenClipboard(IntPtr hWndNewOwner);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool CloseClipboard();

    [LibraryImport("user32.dll")]
    internal static partial int CountClipboardFormats();

    [LibraryImport("user32.dll", SetLastError = true)]
    internal static partial uint EnumClipboardFormats(uint format);

    [LibraryImport("user32.dll", SetLastError = true)]
    internal static partial IntPtr GetClipboardData(uint uFormat);

    [LibraryImport("user32.dll", EntryPoint = "GetClipboardFormatNameW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int GetClipboardFormatName(uint format, Span<char> lpszFormatName, int cchMaxCount);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    internal static partial IntPtr GlobalLock(IntPtr hMem);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GlobalUnlock(IntPtr hMem);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    internal static partial UIntPtr GlobalSize(IntPtr hMem);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    internal static partial uint GetOEMCP();

    [LibraryImport("gdi32.dll", EntryPoint = "GetObjectW", SetLastError = true)]
    internal static partial int GetObject(IntPtr h, int c, out BITMAP pv);

    [LibraryImport("gdi32.dll", SetLastError = true)]
    internal static partial int GetDIBits(IntPtr hdc, IntPtr hbm, uint start, uint cLines,
        byte[]? lpvBits, ref BITMAPINFOHEADER lpbmi, uint usage);

    [LibraryImport("user32.dll")]
    internal static partial IntPtr GetDC(IntPtr hwnd);

    [LibraryImport("user32.dll")]
    internal static partial int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [LibraryImport("gdi32.dll", SetLastError = true)]
    internal static partial uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[]? lpData);

    // ── Clipboard write (used for save/load) ────────────────────────────────

    [LibraryImport("kernel32.dll", SetLastError = true)]
    internal static partial IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    internal static partial IntPtr GlobalFree(IntPtr hMem);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool EmptyClipboard();

    [LibraryImport("user32.dll", SetLastError = true)]
    internal static partial IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

    [LibraryImport("user32.dll", EntryPoint = "RegisterClipboardFormatW", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    internal static partial uint RegisterClipboardFormat(string lpszFormat);

    // Creates an HENHMETAFILE from a raw EMF byte stream.
    [LibraryImport("gdi32.dll", SetLastError = true)]
    internal static partial IntPtr SetEnhMetaFileBits(uint nSize, byte[] pb);

    [LibraryImport("gdi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DeleteEnhMetaFile(IntPtr hemf);

    // Creates an HBITMAP device-dependent bitmap from a DIB.
    // lpbmih: pointer to BITMAPINFOHEADER; pjBits: pixel data; lpbmi: BITMAPINFO (= header for 32 bpp BI_RGB).
    [LibraryImport("gdi32.dll", SetLastError = true)]
    internal static partial IntPtr CreateDIBitmap(IntPtr hdc, ref BITMAPINFOHEADER lpbmih,
        uint fdwInit, byte[] pjBits, ref BITMAPINFOHEADER lpbmi, uint iUsage);

    [LibraryImport("gdi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DeleteObject(IntPtr ho);

    [LibraryImport("user32.dll")]
    internal static partial uint GetClipboardSequenceNumber();
}
