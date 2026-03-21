using System.Buffers.Binary;
using System.IO;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Creates WPF <see cref="BitmapSource"/> previews from raw clipboard format bytes.
/// Supports DIB/DIBV5 (CF_DIB, CF_DIBV5, HBITMAP-derived) and encoded image formats
/// (PNG, JPEG, GIF, BMP, TIFF, etc.) via WPF's built-in BitmapDecoder.
/// </summary>
internal sealed class ImagePreviewService : IImagePreviewService
{
    // Format IDs that yield HBITMAP data converted to DIB bytes by ClipboardReaderService.
    // CF_BITMAP (2), CF_DSPBITMAP (0x82) — same set as ClipboardReaderService.HBitmapFormats.
    private static readonly HashSet<uint> HBitmapFormats = [2, 0x0082];

    // ── IImagePreviewService ────────────────────────────────────────────────

    /// <inheritdoc/>
    public bool IsImageCompatible(uint formatId, string formatName)
    {
        // CF_DIB (8), CF_DIBV5 (17), CF_BITMAP (2), CF_DSPBITMAP (0x82)
        if (formatId is 8 or 17 || HBitmapFormats.Contains(formatId))
            return true;

        var normalized = formatName.ToLowerInvariant();
        return normalized.Contains("png",    StringComparison.Ordinal) ||
               normalized.Contains("jpeg",   StringComparison.Ordinal) ||
               normalized.Contains("jpg",    StringComparison.Ordinal) ||
               normalized.Contains("gif",    StringComparison.Ordinal) ||
               normalized.Contains("dib",    StringComparison.Ordinal) ||
               normalized.Contains("bitmap", StringComparison.Ordinal) ||
               normalized.Contains("image",  StringComparison.Ordinal);
    }

    /// <inheritdoc/>
    public bool TryCreatePreview(uint formatId, string formatName, byte[] data,
        out BitmapSource? preview, out string failureMessage)
    {
        preview = null;

        if (!IsImageCompatible(formatId, formatName))
        {
            failureMessage = "Image preview unavailable for this format.";
            return false;
        }

        if (data.Length == 0)
        {
            failureMessage = "Image data is empty.";
            return false;
        }

        try
        {
            if (formatId is 8 or 17 || HBitmapFormats.Contains(formatId))
            {
                // CF_DIB, CF_DIBV5, and HBITMAP-derived formats all yield a
                // BITMAPINFOHEADER + pixels block.
                if (!TryCreateBitmapFromDib(data, out preview))
                {
                    failureMessage = "Failed to decode DIB image data.";
                    return false;
                }
            }
            else
            {
                preview = CreateBitmapFromEncodedImage(data);
            }

            failureMessage = string.Empty;
            return preview != null;
        }
        catch
        {
            preview        = null;
            failureMessage = "Failed to decode image preview for this format.";
            return false;
        }
    }

    // ── Private helpers ─────────────────────────────────────────────────────

    /// <summary>
    /// Decodes a raw encoded image stream (PNG, JPEG, GIF, BMP, TIFF, etc.)
    /// using WPF's <see cref="BitmapDecoder"/>.
    /// </summary>
    private static BitmapSource CreateBitmapFromEncodedImage(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        var decoder = BitmapDecoder.Create(
            stream,
            BitmapCreateOptions.PreservePixelFormat,
            BitmapCacheOption.OnLoad);
        var frame = decoder.Frames[0];
        frame.Freeze();
        return frame;
    }

    /// <summary>
    /// Converts a DIB byte block (BITMAPINFOHEADER + optional colour table + pixels)
    /// to a WPF <see cref="BitmapSource"/> by prepending a BITMAPFILEHEADER and
    /// decoding the resulting BMP stream.
    /// </summary>
    private static bool TryCreateBitmapFromDib(byte[] dibBytes, out BitmapSource? bitmap)
    {
        bitmap = null;
        if (dibBytes.Length < 40)
            return false;

        var headerSize = BitConverter.ToUInt32(dibBytes, 0);
        if (headerSize < 40 || headerSize > dibBytes.Length)
            return false;

        var bitCount    = BitConverter.ToUInt16(dibBytes, 14);
        var compression = BitConverter.ToUInt32(dibBytes, 16);
        var colorsUsed  = BitConverter.ToUInt32(dibBytes, 32);

        uint masksSize = 0;
        if ((compression == 3 || compression == 6) && headerSize == 40)
        {
            masksSize = compression == 6 ? 16u : 12u;
        }

        uint colorTableEntries = colorsUsed;
        if (colorTableEntries == 0 && bitCount <= 8)
            colorTableEntries = 1u << bitCount;

        var colorTableSize = colorTableEntries * 4;
        var pixelOffset    = 14u + headerSize + masksSize + colorTableSize;
        if (pixelOffset > dibBytes.Length + 14u)
            return false;

        // Prepend the 14-byte BITMAPFILEHEADER.
        var fileBytes = new byte[dibBytes.Length + 14];
        fileBytes[0] = (byte)'B';
        fileBytes[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(fileBytes.AsSpan(2),  (uint)fileBytes.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(fileBytes.AsSpan(10), pixelOffset);
        Buffer.BlockCopy(dibBytes, 0, fileBytes, 14, dibBytes.Length);

        bitmap = CreateBitmapFromEncodedImage(fileBytes);
        return true;
    }
}
