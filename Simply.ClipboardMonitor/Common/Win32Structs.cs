using System.Runtime.InteropServices;

namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Returned by NtQueryInformationProcess (class 0 = ProcessBasicInformation).
/// Six pointer-sized fields; works for both x86 and x64 because IntPtr adapts to the
/// running process's pointer size.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal struct PROCESS_BASIC_INFORMATION
{
    public IntPtr Reserved1;
    public IntPtr PebBaseAddress;
    public IntPtr Reserved2_0;
    public IntPtr Reserved2_1;
    public IntPtr UniqueProcessId;
    public IntPtr Reserved3;
}

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
