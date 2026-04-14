using Simply.ClipboardMonitor.Models;
using System.Text;
using System.Windows.Media;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

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
            switch (ClassifyFormat(id, name))
            {
                case FormatCategory.Image: hasI = true; break;
                case FormatCategory.Html:  hasH = true; break;
                case FormatCategory.Rtf:   hasR = true; break;
                case FormatCategory.Text:  hasT = true; break;
                case FormatCategory.File:  hasF = true; break;
                default:                   hasO = true; break;
            }
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

    private enum FormatCategory { Image, Html, Rtf, Text, File, Other }

    private static FormatCategory ClassifyFormat(uint id, string name)
    {
        if (IsImageFormat(id, name))   return FormatCategory.Image;
        if (IsHtmlFormat(name))        return FormatCategory.Html;
        if (IsRtfFormat(name))         return FormatCategory.Rtf;
        if (id == CF_TEXT || id == CF_OEMTEXT || id == CF_UNICODETEXT ||
            name.Contains("text", StringComparison.OrdinalIgnoreCase))   return FormatCategory.Text;
        if (id == CF_HDROP)            return FormatCategory.File;
        return FormatCategory.Other;
    }

    /// <inheritdoc/>
    public string? GetFormatPillLabel(uint formatId, string formatName) =>
        ClassifyFormat(formatId, formatName) switch
        {
            FormatCategory.Image => "IMG",
            FormatCategory.Html  => "HTML",
            FormatCategory.Rtf   => "RTF",
            FormatCategory.Text  => "TXT",
            FormatCategory.File  => "FILE",
            _                    => null,
        };
}
