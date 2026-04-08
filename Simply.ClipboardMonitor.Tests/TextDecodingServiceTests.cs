using Simply.ClipboardMonitor.Services.Impl;
using System.Text;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

public class TextDecodingServiceTests
{
    private readonly TextDecodingService _sut = new();

    // ── Format ID constants ─────────────────────────────────────────────────
    private const uint CF_TEXT        = 1;
    private const uint CF_OEMTEXT     = 7;
    private const uint CF_UNICODETEXT = 13;
    private const uint CF_DIB         = 8;

    // ── Format name constants (used across multiple tests) ──────────────────
    private const string FormatNameCfText        = "CF_TEXT";
    private const string FormatNameCfUnicodeText = "CF_UNICODETEXT";
    private const string FormatNameCfDib         = "CF_DIB";
    private const string FormatNameTextPlain     = "text/plain";

    // ── IsTextCompatible ────────────────────────────────────────────────────

    [Fact] public void IsTextCompatible_CfText_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(CF_TEXT, FormatNameCfText));

    [Fact] public void IsTextCompatible_CfOemText_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(CF_OEMTEXT, "CF_OEMTEXT"));

    [Fact] public void IsTextCompatible_CfUnicodeText_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(CF_UNICODETEXT, FormatNameCfUnicodeText));

    [Fact] public void IsTextCompatible_NameContainingHtml_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(99999u, "HTML Format"));

    [Fact] public void IsTextCompatible_NameContainingRtf_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(99999u, "Rich Text Format"));

    [Fact] public void IsTextCompatible_NameContainingXml_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(99999u, "text/xml"));

    [Fact] public void IsTextCompatible_NameContainingJson_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(99999u, "data.json"));

    [Fact] public void IsTextCompatible_NameContainingCsv_ReturnsTrue()
        => Assert.True(_sut.IsTextCompatible(99999u, "export.csv"));

    [Fact] public void IsTextCompatible_CfDib_ReturnsFalse()
        => Assert.False(_sut.IsTextCompatible(CF_DIB, FormatNameCfDib));

    [Fact] public void IsTextCompatible_ArbitraryBinaryFormatName_ReturnsFalse()
        => Assert.False(_sut.IsTextCompatible(99999u, "CF_BITMAP"));

    // ── Decode ──────────────────────────────────────────────────────────────

    [Fact]
    public void Decode_CfUnicodeText_DecodesUtf16LeAndStripsNullTerminator()
    {
        var data   = Encoding.Unicode.GetBytes("Hello\0");
        var result = _sut.Decode(CF_UNICODETEXT, FormatNameCfUnicodeText, data);

        Assert.True(result.Success);
        Assert.Equal("Hello", result.Text);
        Assert.Equal(Encoding.Unicode, result.DetectedEncoding);
    }

    [Fact]
    public void Decode_CfText_DecodesAndStripsNullTerminator()
    {
        var data   = Encoding.Default.GetBytes("World\0");
        var result = _sut.Decode(CF_TEXT, FormatNameCfText, data);

        Assert.True(result.Success);
        Assert.Equal("World", result.Text);
    }

    [Fact]
    public void Decode_NonTextFormat_ReturnsFailure()
    {
        var result = _sut.Decode(CF_DIB, FormatNameCfDib, [0x42, 0x4D]);

        Assert.False(result.Success);
        Assert.Null(result.Text);
    }

    [Fact]
    public void Decode_UnknownTextFormatWithUtf8Bytes_DecodesCorrectly()
    {
        var data   = Encoding.UTF8.GetBytes("Héllo");
        var result = _sut.Decode(99999u, FormatNameTextPlain, data);

        Assert.True(result.Success);
        Assert.Equal("Héllo", result.Text);
    }

    [Fact]
    public void Decode_Utf8Bom_StripsBomAndDecodesCorrectly()
    {
        byte[] bom  = [0xEF, 0xBB, 0xBF];
        var    text = Encoding.UTF8.GetBytes("BOM test");
        var    data = bom.Concat(text).ToArray();

        var result = _sut.Decode(99999u, FormatNameTextPlain, data);

        Assert.True(result.Success);
        Assert.Equal("BOM test", result.Text);
    }

    [Fact]
    public void Decode_Utf16LeBom_DecodesCorrectly()
    {
        byte[] bom  = [0xFF, 0xFE];
        var    text = Encoding.Unicode.GetBytes("LE BOM");
        var    data = bom.Concat(text).ToArray();

        var result = _sut.Decode(99999u, FormatNameTextPlain, data);

        Assert.True(result.Success);
        Assert.Equal("LE BOM", result.Text);
    }

    [Fact]
    public void Decode_Utf16BeBom_DecodesCorrectly()
    {
        byte[] bom  = [0xFE, 0xFF];
        var    text = Encoding.BigEndianUnicode.GetBytes("BE BOM");
        var    data = bom.Concat(text).ToArray();

        var result = _sut.Decode(99999u, FormatNameTextPlain, data);

        Assert.True(result.Success);
        Assert.Equal("BE BOM", result.Text);
    }

    // ── DecodeWith ──────────────────────────────────────────────────────────

    [Fact]
    public void DecodeWith_ValidEncoding_Succeeds()
    {
        var data   = Encoding.UTF8.GetBytes("explicit");
        var result = _sut.DecodeWith(data, Encoding.UTF8);

        Assert.True(result.Success);
        Assert.Equal("explicit", result.Text);
    }

    [Fact]
    public void DecodeWith_InvalidBytesForStrictEncoding_ReturnsFailure()
    {
        var strictUtf8   = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
        var invalidBytes = new byte[] { 0x80, 0x81, 0x82 }; // invalid UTF-8 continuation bytes

        var result = _sut.DecodeWith(invalidBytes, strictUtf8);

        Assert.False(result.Success);
        Assert.Null(result.Text);
    }

    // ── GetDecodedTextStats ──────────────────────────────────────────────────

    [Fact]
    public void GetDecodedTextStats_EmptyString_ShowsAllZeros()
        => Assert.Equal("0 characters (0 non-whitespace) · 0 lines",
                        _sut.GetDecodedTextStats(""));

    [Fact]
    public void GetDecodedTextStats_SingleWordNoWhitespace_OneLine()
        => Assert.Equal("5 characters (5 non-whitespace) · 1 line",
                        _sut.GetDecodedTextStats("Hello"));

    [Fact]
    public void GetDecodedTextStats_WordWithSpace_CountsNonWhitespace()
    {
        var stats = _sut.GetDecodedTextStats("Hi There");
        Assert.Contains("7 non-whitespace", stats);
        Assert.Contains("1 line", stats);
    }

    [Fact]
    public void GetDecodedTextStats_LfNewline_CountsTwoLines()
        => Assert.Contains("2 lines", _sut.GetDecodedTextStats("a\nb"));

    [Fact]
    public void GetDecodedTextStats_CrLf_CountsAsOneLineBreak()
        => Assert.Contains("2 lines", _sut.GetDecodedTextStats("a\r\nb"));

    [Fact]
    public void GetDecodedTextStats_StandaloneCr_CountsAsLineBreak()
        => Assert.Contains("2 lines", _sut.GetDecodedTextStats("a\rb"));

    [Fact]
    public void GetDecodedTextStats_ThreeLinesViaCrLf_CountsCorrectly()
        => Assert.Contains("3 lines", _sut.GetDecodedTextStats("a\r\nb\r\nc"));
}
