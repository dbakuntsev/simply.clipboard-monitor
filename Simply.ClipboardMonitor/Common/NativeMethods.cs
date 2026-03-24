using System.Runtime.InteropServices;
using System.Text;

namespace Simply.ClipboardMonitor.Common;

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

    // ── System tray (Shell_NotifyIcon) ───────────────────────────────────────

    internal const uint NIM_ADD    = 0x00000000;
    internal const uint NIM_MODIFY = 0x00000001;
    internal const uint NIM_DELETE = 0x00000002;

    internal const uint NIF_MESSAGE = 0x00000001;
    internal const uint NIF_ICON    = 0x00000002;
    internal const uint NIF_TIP     = 0x00000004;
    internal const uint NIF_INFO    = 0x00000010;

    internal const uint NIIF_INFO   = 0x00000001;

    internal const uint MF_STRING    = 0x00000000;
    internal const uint MF_SEPARATOR = 0x00000800;

    // Fixed-buffer fields require an unsafe struct; Shell_NotifyIcon is declared
    // with LibraryImport so the struct must be blittable (no ByValTStr).
    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NOTIFYICONDATA
    {
        public uint   cbSize;
        public IntPtr hWnd;
        public uint   uID;
        public uint   uFlags;
        public uint   uCallbackMessage;
        public IntPtr hIcon;
        public fixed char szTip[128];
        public uint   dwState;
        public uint   dwStateMask;
        public fixed char szInfo[256];
        public uint   uTimeoutOrVersion;
        public fixed char szInfoTitle[64];
        public uint   dwInfoFlags;
        public Guid   guidItem;
        public IntPtr hBalloonIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int x;
        public int y;
    }

    [LibraryImport("shell32.dll", EntryPoint = "Shell_NotifyIconW", StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static unsafe partial bool Shell_NotifyIcon(uint dwMessage, ref NOTIFYICONDATA lpData);

    [LibraryImport("user32.dll")]
    internal static partial IntPtr CreatePopupMenu();

    [LibraryImport("user32.dll", EntryPoint = "AppendMenuW", StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool AppendMenu(IntPtr hMenu, uint uFlags, UIntPtr uIDNewItem, string? lpNewItem);

    [LibraryImport("user32.dll")]
    internal static partial uint TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DestroyMenu(IntPtr hMenu);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GetCursorPos(out POINT lpPoint);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool SetForegroundWindow(IntPtr hWnd);

    [LibraryImport("user32.dll", EntryPoint = "PostMessageW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [LibraryImport("kernel32.dll", EntryPoint = "GetModuleHandleW", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial IntPtr GetModuleHandle(string? lpModuleName);
}
