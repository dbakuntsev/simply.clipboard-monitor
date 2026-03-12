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

internal static class NativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool AddClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern bool CloseClipboard();

    [DllImport("user32.dll")]
    internal static extern int CountClipboardFormats();

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern uint EnumClipboardFormats(uint format);

    [DllImport("user32.dll", SetLastError = true)]
    internal static extern IntPtr GetClipboardData(uint uFormat);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int GetClipboardFormatName(uint format, StringBuilder lpszFormatName, int cchMaxCount);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern bool GlobalUnlock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern UIntPtr GlobalSize(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    internal static extern uint GetOEMCP();

    [DllImport("gdi32.dll", SetLastError = true)]
    internal static extern int GetObject(IntPtr h, int c, out BITMAP pv);

    [DllImport("gdi32.dll", SetLastError = true)]
    internal static extern int GetDIBits(IntPtr hdc, IntPtr hbm, uint start, uint cLines,
        byte[]? lpvBits, ref BITMAPINFOHEADER lpbmi, uint usage);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetDC(IntPtr hwnd);

    [DllImport("user32.dll")]
    internal static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("gdi32.dll", SetLastError = true)]
    internal static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[]? lpData);
}
