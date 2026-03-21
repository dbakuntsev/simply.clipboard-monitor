using Simply.ClipboardMonitor.Models;
using System.Text;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Decodes raw clipboard bytes as text, with auto-detection and manual override support.
/// </summary>
public interface ITextDecodingService
{
    /// <summary>True when the format is eligible for a text preview.</summary>
    bool IsTextCompatible(uint formatId, string formatName);

    /// <summary>
    /// Auto-detects encoding and decodes <paramref name="data"/>.
    /// Uses format-specific rules for the well-known text format IDs.
    /// </summary>
    TextDecodeResult Decode(uint formatId, string formatName, byte[] data);

    /// <summary>Re-decodes <paramref name="data"/> with the given explicit encoding.</summary>
    TextDecodeResult DecodeWith(byte[] data, Encoding encoding, int unitSize = 1);

    /// <summary>Returns all encodings available for manual selection.</summary>
    IReadOnlyList<EncodingItem> GetAvailableEncodings();

    /// <summary>Formats character/line statistics for the status bar.</summary>
    string GetDecodedTextStats(string text);
}
