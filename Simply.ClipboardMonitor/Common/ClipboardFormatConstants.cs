namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Central registry of well-known Windows clipboard format IDs and their classification sets.
/// Keeping these constants in one place ensures all services agree on the same values and that
/// supporting a new standard format requires editing only this file.
/// </summary>
internal static class ClipboardFormatConstants
{
    // ── Named format IDs ────────────────────────────────────────────────────

    internal const uint CF_TEXT           = 1;
    internal const uint CF_BITMAP         = 2;
    internal const uint CF_METAFILEPICT   = 3;
    internal const uint CF_SYLK           = 4;
    internal const uint CF_DIF            = 5;
    internal const uint CF_TIFF           = 6;
    internal const uint CF_OEMTEXT        = 7;
    internal const uint CF_DIB            = 8;
    internal const uint CF_PALETTE        = 9;
    internal const uint CF_PENDATA        = 10;
    internal const uint CF_RIFF           = 11;
    internal const uint CF_WAVE           = 12;
    internal const uint CF_UNICODETEXT    = 13;
    internal const uint CF_ENHMETAFILE    = 14;
    internal const uint CF_HDROP          = 15;
    internal const uint CF_LOCALE         = 16;
    internal const uint CF_DIBV5          = 17;
    internal const uint CF_OWNERDISPLAY   = 0x0080;
    internal const uint CF_DSPTEXT        = 0x0081;
    internal const uint CF_DSPBITMAP      = 0x0082;
    internal const uint CF_DSPMETAFILEPICT = 0x0083;
    internal const uint CF_DSPENHMETAFILE = 0x008E;

    // ── Display name map ────────────────────────────────────────────────────

    /// <summary>Standard Windows clipboard format names keyed by format ID.</summary>
    internal static readonly IReadOnlyDictionary<uint, string> WellKnownFormats =
        new Dictionary<uint, string>
        {
            [CF_TEXT]           = "CF_TEXT",
            [CF_BITMAP]         = "CF_BITMAP",
            [CF_METAFILEPICT]   = "CF_METAFILEPICT",
            [CF_SYLK]           = "CF_SYLK",
            [CF_DIF]            = "CF_DIF",
            [CF_TIFF]           = "CF_TIFF",
            [CF_OEMTEXT]        = "CF_OEMTEXT",
            [CF_DIB]            = "CF_DIB",
            [CF_PALETTE]        = "CF_PALETTE",
            [CF_PENDATA]        = "CF_PENDATA",
            [CF_RIFF]           = "CF_RIFF",
            [CF_WAVE]           = "CF_WAVE",
            [CF_UNICODETEXT]    = "CF_UNICODETEXT",
            [CF_ENHMETAFILE]    = "CF_ENHMETAFILE",
            [CF_HDROP]          = "CF_HDROP",
            [CF_LOCALE]         = "CF_LOCALE",
            [CF_DIBV5]          = "CF_DIBV5",
            [CF_OWNERDISPLAY]   = "CF_OWNERDISPLAY",
            [CF_DSPTEXT]        = "CF_DSPTEXT",
            [CF_DSPBITMAP]      = "CF_DSPBITMAP",
            [CF_DSPMETAFILEPICT] = "CF_DSPMETAFILEPICT",
            [CF_DSPENHMETAFILE] = "CF_DSPENHMETAFILE",
        };

    // ── Handle-type name constants ──────────────────────────────────────────

    /// <summary>String tokens used to identify a clipboard format's underlying handle type.</summary>
    internal static class HandleTypes
    {
        internal const string HGlobal      = "hglobal";
        internal const string HBitmap      = "hbitmap";
        internal const string HEnhMetaFile = "henhmetafile";
        internal const string None         = "none";
    }

    // ── Handle-type classification sets ────────────────────────────────────

    /// <summary>CF_BITMAP / CF_DSPBITMAP: stored as HBITMAP, converted to DIB via GetDIBits.</summary>
    internal static readonly IReadOnlySet<uint> HBitmapFormats =
        new HashSet<uint> { CF_BITMAP, CF_DSPBITMAP };

    /// <summary>CF_ENHMETAFILE / CF_DSPENHMETAFILE: stored as HENHMETAFILE, raw bytes via GetEnhMetaFileBits.</summary>
    internal static readonly IReadOnlySet<uint> HEnhMetaFileFormats =
        new HashSet<uint> { CF_ENHMETAFILE, CF_DSPENHMETAFILE };

    /// <summary>Formats whose handles cannot be usefully read as raw bytes (e.g. HPALETTE).</summary>
    internal static readonly IReadOnlySet<uint> NonGlobalMemoryFormats =
        new HashSet<uint> { CF_PALETTE };

    // ── Format classification helpers ────────────────────────────────────────

    /// <summary>
    /// Returns true if the format is likely to contain image data, based on its ID and/or name.
    /// Covers HBITMAP and DIB handle types, plus common encoded image format names.
    /// </summary>
    internal static bool IsImageFormat(uint formatId, string formatName)
    {
        if (HBitmapFormats.Contains(formatId) || formatId == CF_DIB || formatId == CF_DIBV5)
            return true;
        var n = formatName.ToLowerInvariant();
        return n.Contains("png")    || n.Contains("jpeg")   || n.Contains("jpg")  ||
               n.Contains("gif")    || n.Contains("dib")    || n.Contains("bitmap") ||
               n.Contains("image");
    }
}
