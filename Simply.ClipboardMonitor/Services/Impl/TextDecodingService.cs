using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.Text;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Decodes raw clipboard bytes as human-readable text.
/// Encoding auto-detection follows this priority order:
/// <list type="number">
///   <item>Format-specific rules (CF_UNICODETEXT, CF_TEXT, CF_OEMTEXT)</item>
///   <item>UTF-8 BOM</item>
///   <item>UTF-16 LE/BE BOM or null-byte heuristic</item>
///   <item>Strict UTF-8</item>
///   <item>System ANSI code page</item>
/// </list>
/// </summary>
internal sealed class TextDecodingService : ITextDecodingService
{
    private IReadOnlyList<EncodingItem>? _encodingCache;

    // ── ITextDecodingService ────────────────────────────────────────────────

    /// <inheritdoc/>
    public bool IsTextCompatible(uint formatId, string formatName)
    {
        var normalized = formatName.ToLowerInvariant();
        return formatId is 1 or 7 or 13 ||
               normalized.Contains("text",  StringComparison.Ordinal) ||
               normalized.Contains("html",  StringComparison.Ordinal) ||
               normalized.Contains("rtf",   StringComparison.Ordinal) ||
               normalized.Contains("xml",   StringComparison.Ordinal) ||
               normalized.Contains("json",  StringComparison.Ordinal) ||
               normalized.Contains("csv",   StringComparison.Ordinal);
    }

    /// <inheritdoc/>
    public TextDecodeResult Decode(uint formatId, string formatName, byte[] data)
    {
        if (!IsTextCompatible(formatId, formatName))
        {
            return new TextDecodeResult(
                null, null, false,
                "Text preview unavailable for this format.");
        }

        try
        {
            string   text;
            Encoding encoding;

            if (formatId == 13) // CF_UNICODETEXT
            {
                text     = Encoding.Unicode.GetString(data).TrimEnd('\0');
                encoding = Encoding.Unicode;
            }
            else if (formatId == 1) // CF_TEXT
            {
                text     = Encoding.Default.GetString(TrimAtNull(data, 1));
                encoding = Encoding.Default;
            }
            else if (formatId == 7) // CF_OEMTEXT
            {
                var oemEncoding = Encoding.GetEncoding((int)NativeMethods.GetOEMCP());
                text            = oemEncoding.GetString(TrimAtNull(data, 1));
                encoding        = oemEncoding;
            }
            else
            {
                text     = DecodeWithFallback(data, out encoding);
            }

            return new TextDecodeResult(text, encoding, true);
        }
        catch (Exception ex)
        {
            return new TextDecodeResult(
                null, null, false,
                $"Failed to decode this format as text: {ex.Message}");
        }
    }

    /// <inheritdoc/>
    public TextDecodeResult DecodeWith(byte[] data, Encoding encoding, int unitSize = 1)
    {
        try
        {
            var decoded = encoding.GetString(data).TrimEnd('\0');
            return new TextDecodeResult(decoded, encoding, true);
        }
        catch (Exception ex)
        {
            return new TextDecodeResult(
                null, encoding, false,
                $"Cannot decode as {encoding.EncodingName}: {ex.Message}");
        }
    }

    /// <inheritdoc/>
    public IReadOnlyList<EncodingItem> GetAvailableEncodings()
    {
        if (_encodingCache != null)
            return _encodingCache;

        _encodingCache = Encoding.GetEncodings()
            .Select(info =>
            {
                try   { return new EncodingItem(info.GetEncoding(), $"{info.DisplayName} ({info.Name})"); }
                catch { return null; }
            })
            .OfType<EncodingItem>()
            .Append(new EncodingItem(
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                "Unicode (UTF-8, strict) (utf-8-strict)"))
            .OrderBy(e => e.DisplayName)
            .ToList();

        return _encodingCache;
    }

    /// <inheritdoc/>
    public string GetDecodedTextStats(string text)
    {
        int chars    = text.Length;
        int nonWs    = text.Count(c => !char.IsWhiteSpace(c));
        int newlines = 0;

        for (int i = 0; i < text.Length; i++)
        {
            // Count standalone \n and \r as well as \r\n sequences as a single newline.
            if (text[i] == '\r')
            {
                newlines++;
                if (i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
            }
            else if (text[i] == '\n')
            {
                newlines++;
            }
        }

        int lines = chars == 0 ? 0 : newlines + 1;
        return $"{chars:N0} character{(chars == 1 ? "" : "s")} " +
               $"({nonWs:N0} non-whitespace) · " +
               $"{lines:N0} line{(lines == 1 ? "" : "s")}";
    }

    // ── Private helpers ─────────────────────────────────────────────────────

    private enum Utf16Variant { None, LittleEndian, BigEndian }

    /// <summary>
    /// Detects UTF-16 encoding via BOM or heuristic null-byte pattern.
    /// For LE, ASCII-range characters place 0x00 at every odd byte position.
    /// For BE, they place 0x00 at every even byte position.
    /// </summary>
    private static Utf16Variant DetectUtf16(byte[] bytes)
    {
        if (bytes.Length < 4 || bytes.Length % 2 != 0)
            return Utf16Variant.None;

        // Explicit BOM takes priority.
        if (bytes[0] == 0xFF && bytes[1] == 0xFE) return Utf16Variant.LittleEndian;
        if (bytes[0] == 0xFE && bytes[1] == 0xFF) return Utf16Variant.BigEndian;

        // Heuristic: sample up to the first 512 bytes.
        var sampleLen = Math.Min(bytes.Length, 512);
        var pairs     = sampleLen / 2;
        int nullAtOdd = 0, nullAtEven = 0;

        for (var i = 0; i < sampleLen - 1; i += 2)
        {
            if (bytes[i]     == 0) nullAtEven++;
            if (bytes[i + 1] == 0) nullAtOdd++;
        }

        // Require ≥80 % of high bytes to be null AND <5 % of low bytes to be null.
        const double highThreshold = 0.80;
        const double lowMaxRatio   = 0.05;

        if ((double)nullAtOdd  / pairs >= highThreshold && (double)nullAtEven / pairs < lowMaxRatio)
            return Utf16Variant.LittleEndian;
        if ((double)nullAtEven / pairs >= highThreshold && (double)nullAtOdd  / pairs < lowMaxRatio)
            return Utf16Variant.BigEndian;

        return Utf16Variant.None;
    }

    private static string DecodeWithFallback(byte[] bytes, out Encoding usedEncoding)
    {
        // UTF-8 BOM
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            usedEncoding = Encoding.UTF8;
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3).TrimEnd('\0');
        }

        // UTF-16 LE / BE — BOM or heuristic
        var utf16 = DetectUtf16(bytes);
        if (utf16 == Utf16Variant.LittleEndian)
        {
            usedEncoding = Encoding.Unicode;
            var start    = (bytes[0] == 0xFF && bytes[1] == 0xFE) ? 2 : 0;
            var trimmed  = TrimAtNull(bytes[start..], 2);
            return Encoding.Unicode.GetString(trimmed).TrimEnd('\0');
        }
        if (utf16 == Utf16Variant.BigEndian)
        {
            usedEncoding = Encoding.BigEndianUnicode;
            var start    = (bytes[0] == 0xFE && bytes[1] == 0xFF) ? 2 : 0;
            var trimmed  = TrimAtNull(bytes[start..], 2);
            return Encoding.BigEndianUnicode.GetString(trimmed).TrimEnd('\0');
        }

        // Strict UTF-8, then system ANSI
        try
        {
            usedEncoding = Encoding.UTF8;
            return new UTF8Encoding(false, true).GetString(bytes).TrimEnd('\0');
        }
        catch
        {
            usedEncoding = Encoding.Default;
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }
    }

    /// <summary>Unit-size-aware null terminator trimmer.</summary>
    private static byte[] TrimAtNull(byte[] bytes, int unitSize)
    {
        if (unitSize <= 1)
        {
            var index = Array.IndexOf(bytes, (byte)0);
            return index < 0 ? bytes : bytes[..index];
        }

        for (var i = 0; i <= bytes.Length - unitSize; i += unitSize)
        {
            var allZero = true;
            for (var j = 0; j < unitSize; j++)
            {
                if (bytes[i + j] != 0)
                {
                    allZero = false;
                    break;
                }
            }

            if (allZero)
                return bytes[..i];
        }

        return bytes;
    }
}
