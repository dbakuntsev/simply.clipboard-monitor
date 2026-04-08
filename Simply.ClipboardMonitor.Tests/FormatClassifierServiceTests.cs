using Simply.ClipboardMonitor.Services.Impl;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

// FormatClassifierService creates WPF SolidColorBrush objects in its static
// initialiser.  SolidColorBrush is a DependencyObject and must be constructed
// on an STA thread, so every test here uses [StaFact].
public class FormatClassifierServiceTests
{
    // ── Format ID constants ─────────────────────────────────────────────────
    private const uint CF_TEXT        = 1;
    private const uint CF_BITMAP      = 2;
    private const uint CF_OEMTEXT     = 7;
    private const uint CF_DIB         = 8;
    private const uint CF_UNICODETEXT = 13;
    private const uint CF_HDROP       = 15;
    private const uint CF_DIBV5       = 17;

    // ── Pill label constants (used across multiple tests) ───────────────────
    private const string PillImg   = "IMG";
    private const string PillTxt   = "TXT";
    private const string PillHtml  = "HTML";
    private const string PillRtf   = "RTF";
    private const string PillFile  = "FILE";
    private const string PillOther = "OTHER";

    // ── Format name constants (used across multiple tests) ──────────────────
    private const string FormatNameCfDib         = "CF_DIB";
    private const string FormatNameCfText        = "CF_TEXT";
    private const string FormatNameCfUnicodeText = "CF_UNICODETEXT";
    private const string FormatNameCfHdrop       = "CF_HDROP";
    private const string FormatNameHtmlFormat    = "HTML Format";
    private const string FormatNameSomePrivate   = "SomePrivateFormat";

    // ── GetFormatPillLabel ──────────────────────────────────────────────────

    [StaFact] public void GetFormatPillLabel_CfDib_ReturnsImg()
        => Assert.Equal(PillImg, new FormatClassifierService().GetFormatPillLabel(CF_DIB, FormatNameCfDib));

    [StaFact] public void GetFormatPillLabel_CfDibV5_ReturnsImg()
        => Assert.Equal(PillImg, new FormatClassifierService().GetFormatPillLabel(CF_DIBV5, "CF_DIBV5"));

    [StaFact] public void GetFormatPillLabel_CfBitmap_ReturnsImg()
        => Assert.Equal(PillImg, new FormatClassifierService().GetFormatPillLabel(CF_BITMAP, "CF_BITMAP"));

    [StaFact] public void GetFormatPillLabel_NameContainingPng_ReturnsImg()
        => Assert.Equal(PillImg, new FormatClassifierService().GetFormatPillLabel(99999u, "image/png"));

    [StaFact] public void GetFormatPillLabel_NameContainingJpeg_ReturnsImg()
        => Assert.Equal(PillImg, new FormatClassifierService().GetFormatPillLabel(99999u, "image/jpeg"));

    [StaFact] public void GetFormatPillLabel_CfText_ReturnsTxt()
        => Assert.Equal(PillTxt, new FormatClassifierService().GetFormatPillLabel(CF_TEXT, FormatNameCfText));

    [StaFact] public void GetFormatPillLabel_CfUnicodeText_ReturnsTxt()
        => Assert.Equal(PillTxt, new FormatClassifierService().GetFormatPillLabel(CF_UNICODETEXT, FormatNameCfUnicodeText));

    [StaFact] public void GetFormatPillLabel_CfOemText_ReturnsTxt()
        => Assert.Equal(PillTxt, new FormatClassifierService().GetFormatPillLabel(CF_OEMTEXT, "CF_OEMTEXT"));

    [StaFact] public void GetFormatPillLabel_HtmlFormatName_ReturnsHtml()
        => Assert.Equal(PillHtml, new FormatClassifierService().GetFormatPillLabel(99999u, FormatNameHtmlFormat));

    [StaFact] public void GetFormatPillLabel_RtfName_ReturnsRtf()
        => Assert.Equal(PillRtf, new FormatClassifierService().GetFormatPillLabel(99999u, "Rich Text Format"));

    [StaFact] public void GetFormatPillLabel_RtfAbbreviation_ReturnsRtf()
        => Assert.Equal(PillRtf, new FormatClassifierService().GetFormatPillLabel(99999u, "text/rtf"));

    [StaFact] public void GetFormatPillLabel_CfHdrop_ReturnsFile()
        => Assert.Equal(PillFile, new FormatClassifierService().GetFormatPillLabel(CF_HDROP, FormatNameCfHdrop));

    [StaFact] public void GetFormatPillLabel_UnknownFormat_ReturnsNull()
        => Assert.Null(new FormatClassifierService().GetFormatPillLabel(99999u, FormatNameSomePrivate));

    // ── ComputePills ────────────────────────────────────────────────────────

    [StaFact]
    public void ComputePills_EmptyList_ReturnsEmptyList()
    {
        var pills = new FormatClassifierService().ComputePills([]);
        Assert.Empty(pills);
    }

    [StaFact]
    public void ComputePills_SingleUnknownFormat_ReturnsOther()
    {
        var pills = new FormatClassifierService().ComputePills([(99999u, FormatNameSomePrivate)]);
        Assert.Single(pills);
        Assert.Equal(PillOther, pills[0].Label);
    }

    [StaFact]
    public void ComputePills_SingleImageFormat_ReturnsImgOnly()
    {
        var pills = new FormatClassifierService().ComputePills([(CF_DIB, FormatNameCfDib)]);
        Assert.Single(pills);
        Assert.Equal(PillImg, pills[0].Label);
    }

    [StaFact]
    public void ComputePills_KnownAndUnknownFormats_OtherSuppressed()
    {
        // PillOther must not appear when at least one well-known category is present.
        var formats = new (uint, string)[] { (CF_TEXT, FormatNameCfText), (99999u, FormatNameSomePrivate) };
        var pills   = new FormatClassifierService().ComputePills(formats);

        Assert.DoesNotContain(pills, p => p.Label == PillOther);
        Assert.Contains(pills, p => p.Label == PillTxt);
    }

    [StaFact]
    public void ComputePills_MultipleCategories_ReturnsAllMatchingPills()
    {
        var formats = new (uint, string)[]
        {
            (CF_DIB,         FormatNameCfDib),
            (CF_UNICODETEXT, FormatNameCfUnicodeText),
            (99999u,         FormatNameHtmlFormat),
            (CF_HDROP,       FormatNameCfHdrop),
        };
        var labels = new FormatClassifierService().ComputePills(formats)
                         .Select(p => p.Label)
                         .ToHashSet();

        Assert.Contains(PillImg,  labels);
        Assert.Contains(PillTxt,  labels);
        Assert.Contains(PillHtml, labels);
        Assert.Contains(PillFile, labels);
        Assert.DoesNotContain(PillOther, labels);
    }

    [StaFact]
    public void ComputePills_HtmlTakesPriorityOverText_BothPillsPresent()
    {
        // A format name like FormatNameHtmlFormat matches HTML; FormatNameCfText matches TXT.
        // Both pills should appear.
        var formats = new (uint, string)[] { (CF_TEXT, FormatNameCfText), (99999u, FormatNameHtmlFormat) };
        var labels  = new FormatClassifierService().ComputePills(formats)
                          .Select(p => p.Label)
                          .ToHashSet();

        Assert.Contains(PillTxt,  labels);
        Assert.Contains(PillHtml, labels);
    }

    // ── ComputeTooltip ──────────────────────────────────────────────────────

    [StaFact]
    public void ComputeTooltip_SingleFormat_ShowsSingularCountAndName()
    {
        var tooltip = new FormatClassifierService()
            .ComputeTooltip([(CF_TEXT, FormatNameCfText)]);

        Assert.Contains("1 format",         tooltip);
        Assert.Contains(FormatNameCfText,   tooltip);
        Assert.Contains($"({CF_TEXT})",     tooltip);
    }

    [StaFact]
    public void ComputeTooltip_TwoFormats_ShowsPluralCountAndBothNames()
    {
        var tooltip = new FormatClassifierService()
            .ComputeTooltip([(CF_TEXT, FormatNameCfText), (CF_UNICODETEXT, FormatNameCfUnicodeText)]);

        Assert.Contains("2 formats",              tooltip);
        Assert.Contains(FormatNameCfText,         tooltip);
        Assert.Contains(FormatNameCfUnicodeText,  tooltip);
    }

    [StaFact]
    public void ComputeTooltip_EmptyList_ShowsZeroFormats()
    {
        var tooltip = new FormatClassifierService().ComputeTooltip([]);
        Assert.Contains("0 formats", tooltip);
    }
}
