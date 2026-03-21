using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.Runtime.InteropServices;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Reads the current state of the Windows clipboard.
/// All Win32 clipboard interaction (enumeration, data reading, size queries) is
/// encapsulated here so that callers depend only on <see cref="IClipboardReader"/>.
/// Handle-type dispatch is delegated to injected <see cref="IHandleReadStrategy"/> instances.
/// </summary>
internal sealed class ClipboardReaderService : IClipboardReader
{
    private readonly IReadOnlyDictionary<string, IHandleReadStrategy> _readStrategies;

    public ClipboardReaderService(IEnumerable<IHandleReadStrategy> handleReadStrategies)
    {
        _readStrategies = handleReadStrategies.ToDictionary(s => s.HandleType);
    }

    // ── IClipboardReader ────────────────────────────────────────────────────

    /// <inheritdoc/>
    public IReadOnlyList<ClipboardFormatItem> EnumerateFormats()
    {
        var results = new List<ClipboardFormatItem>();

        if (!TryOpenClipboard(IntPtr.Zero))
            return results;

        try
        {
            uint format  = 0;
            int  ordinal = 1;
            while (true)
            {
                format = NativeMethods.EnumClipboardFormats(format);
                if (format == 0) break;

                var name           = GetFormatDisplayName(format);
                var hasSize        = TryGetClipboardDataSize(format, out var sizeBytes);
                var contentSizeVal = hasSize ? (long)sizeBytes : -1L;
                var contentSize    = hasSize ? sizeBytes.ToString("N0") : "n/a";

                results.Add(new ClipboardFormatItem(ordinal, format, name, contentSize, contentSizeVal));
                ordinal++;
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        return results;
    }

    /// <inheritdoc/>
    public IReadOnlyList<FormatSnapshot> CaptureAllFormats(IReadOnlyList<ClipboardFormatItem> formats)
    {
        if (formats.Count == 0)
            return [];

        if (!TryOpenClipboard(IntPtr.Zero))
            return [];

        var snapshots = new List<FormatSnapshot>(formats.Count);
        try
        {
            foreach (var item in formats)
            {
                var handleType   = GetHandleType(item.FormatId);
                byte[]? data     = null;
                if (handleType != "none")
                    TryReadFormatBytes(item.FormatId, handleType, out data, out _);

                var originalSize = data?.LongLength
                    ?? (item.ContentSizeValue >= 0 ? item.ContentSizeValue : 0L);

                snapshots.Add(new FormatSnapshot(item.Ordinal, item.FormatId, item.Name,
                                                 handleType, data, originalSize));
            }
        }
        finally
        {
            CloseClipboard();
        }

        return snapshots;
    }

    /// <inheritdoc/>
    public bool TryReadFormatBytes(uint formatId, string handleType,
        out byte[]? data, out string failureMessage)
    {
        data = null;
        var strategy = _readStrategies.TryGetValue(handleType, out var s)
            ? s : _readStrategies["hglobal"];
        return strategy.TryRead(formatId, out data, out failureMessage);
    }

    /// <inheritdoc/>
    public string GetFormatDisplayName(uint format)
    {
        if (WellKnownFormats.TryGetValue(format, out var wellKnownName))
            return wellKnownName;

        Span<char> buffer = stackalloc char[256];
        var chars  = NativeMethods.GetClipboardFormatName(format, buffer, buffer.Length);
        if (chars > 0)
            return new string(buffer[..chars]);

        return "Unknown";
    }

    /// <inheritdoc/>
    public string GetHandleType(uint formatId) =>
        HBitmapFormats.Contains(formatId)         ? "hbitmap"      :
        HEnhMetaFileFormats.Contains(formatId)    ? "henhmetafile" :
        NonGlobalMemoryFormats.Contains(formatId) ? "none"         :
        "hglobal";

    /// <inheritdoc/>
    public uint GetSequenceNumber() => NativeMethods.GetClipboardSequenceNumber();

    /// <inheritdoc/>
    public bool TryOpenClipboard(IntPtr hwnd)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            if (NativeMethods.OpenClipboard(hwnd))
                return true;

            Thread.Sleep(20);
        }

        return false;
    }

    /// <inheritdoc/>
    public void CloseClipboard() => NativeMethods.CloseClipboard();

    // ── Private size-query helpers ──────────────────────────────────────────

    private static bool TryGetClipboardDataSize(uint format, out ulong sizeBytes)
    {
        sizeBytes = 0;

        if (NonGlobalMemoryFormats.Contains(format))
            return false;

        var handle = NativeMethods.GetClipboardData(format);
        if (handle == IntPtr.Zero)
            return false;

        if (HBitmapFormats.Contains(format))
        {
            if (NativeMethods.GetObject(handle, Marshal.SizeOf<BITMAP>(), out var bm) == 0)
                return false;
            var stride = (bm.bmWidth * bm.bmBitsPixel + 31) / 32 * 4;
            sizeBytes  = (ulong)(Math.Abs(bm.bmHeight) * stride * bm.bmPlanes);
            return sizeBytes > 0;
        }

        if (HEnhMetaFileFormats.Contains(format))
        {
            var emfSize = NativeMethods.GetEnhMetaFileBits(handle, 0, null);
            if (emfSize == 0)
                return false;
            sizeBytes = emfSize;
            return true;
        }

        var size = NativeMethods.GlobalSize(handle);
        if (size == UIntPtr.Zero)
            return false;

        sizeBytes = (ulong)size;
        return true;
    }
}
