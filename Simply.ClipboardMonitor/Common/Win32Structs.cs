using System.Runtime.InteropServices;

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
