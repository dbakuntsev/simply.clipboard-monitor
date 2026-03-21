using Simply.ClipboardMonitor.Models;
using System.Text;
using System.Windows.Media;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Classifies clipboard formats into display categories for the history list,
/// producing coloured pills and tooltip text.
/// </summary>
internal sealed class FormatClassifierService : IFormatClassifier
{
    // ── Pill colours (frozen so they can be safely shared across threads) ───

    private static readonly SolidColorBrush PillBrushImage = MakeBrush(0x5B, 0x9B, 0xD5); // blue   — images
    private static readonly SolidColorBrush PillBrushText  = MakeBrush(0x70, 0xAD, 0x47); // green  — text
    private static readonly SolidColorBrush PillBrushHtml  = MakeBrush(0xED, 0x7D, 0x31); // orange — HTML
    private static readonly SolidColorBrush PillBrushRtf   = MakeBrush(0x9E, 0x52, 0x9F); // purple — RTF
    private static readonly SolidColorBrush PillBrushFile  = MakeBrush(0x8B, 0x65, 0x33); // brown  — files
    private static readonly SolidColorBrush PillBrushOther = MakeBrush(0x80, 0x80, 0x80); // gray   — other

    private static SolidColorBrush MakeBrush(byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }

    // ── IFormatClassifier ───────────────────────────────────────────────────

    /// <inheritdoc/>
    public IReadOnlyList<FormatPill> ComputePills(
        IReadOnlyList<(uint FormatId, string FormatName)> formats)
    {
        bool hasI = false, hasT = false, hasH = false,
             hasR = false, hasF = false, hasO = false;

        foreach (var (id, name) in formats)
        {
            if      (IsImageFormat(id, name)) hasI = true;
            else if (IsHtmlFormat(name))      hasH = true;
            else if (IsRtfFormat(name))       hasR = true;
            else if (IsTextFormat(id, name))  hasT = true;
            else if (IsFileFormat(id))        hasF = true;
            else                              hasO = true;
        }

        var pills = new List<FormatPill>(6);
        if (hasI) pills.Add(new FormatPill("IMG",  PillBrushImage));
        if (hasT) pills.Add(new FormatPill("TXT",  PillBrushText));
        if (hasH) pills.Add(new FormatPill("HTML", PillBrushHtml));
        if (hasR) pills.Add(new FormatPill("RTF",  PillBrushRtf));
        if (hasF) pills.Add(new FormatPill("FILE", PillBrushFile));
        // "OTHER" is a fallback shown only when no well-known format category was recognised.
        if (hasO && pills.Count == 0) pills.Add(new FormatPill("OTHER", PillBrushOther));
        return pills;
    }

    /// <inheritdoc/>
    public string ComputeTooltip(
        IReadOnlyList<(uint FormatId, string FormatName)> formats)
    {
        var sb = new StringBuilder();
        sb.Append($"{formats.Count} format{(formats.Count == 1 ? "" : "s")}");
        foreach (var (id, name) in formats)
            sb.Append($"\n{name} ({id})");
        return sb.ToString();
    }

    // ── Private classification helpers ──────────────────────────────────────

    private static bool IsImageFormat(uint id, string name)
    {
        // CF_BITMAP (2), CF_DIB (8), CF_DIBV5 (17), CF_DSPBITMAP (0x82)
        if (id is 2 or 8 or 17 or 0x0082) return true;
        var n = name.ToLowerInvariant();
        return n.Contains("png")    || n.Contains("jpeg") || n.Contains("jpg") ||
               n.Contains("dib")    || n.Contains("bitmap") || n.Contains("image");
    }

    private static bool IsHtmlFormat(string name) =>
        name.Contains("html", StringComparison.OrdinalIgnoreCase);

    private static bool IsRtfFormat(string name) =>
        name.Contains("rtf",       StringComparison.OrdinalIgnoreCase) ||
        name.Contains("rich text", StringComparison.OrdinalIgnoreCase);

    private static bool IsTextFormat(uint id, string name)
    {
        // CF_TEXT (1), CF_OEMTEXT (7), CF_UNICODETEXT (13)
        if (id is 1 or 7 or 13) return true;
        // Name-based: "text"-like, but not HTML or RTF (those have their own pills).
        return name.Contains("text", StringComparison.OrdinalIgnoreCase) &&
               !IsHtmlFormat(name) && !IsRtfFormat(name);
    }

    private static bool IsFileFormat(uint id) => id == 15; // CF_HDROP
}
