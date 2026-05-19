using Simply.ClipboardMonitor.Services.Impl;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

public class FormatNotificationMatcherTests
{
    [Fact]
    public void Disabled_NeverMatches()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: false, "x-custom-internal");

        Assert.False(matcher.Enabled);
        Assert.False(matcher.HasActivePatterns);
        Assert.Empty(matcher.Match(new[] { "x-custom-internal" }));
    }

    [Fact]
    public void EnabledWithEmptyPatterns_HasNoActivePatterns()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "");

        Assert.True(matcher.Enabled);
        Assert.False(matcher.HasActivePatterns);
        Assert.Empty(matcher.Match(new[] { "CF_TEXT" }));
    }

    [Fact]
    public void EnabledWithWhitespacePatterns_HasNoActivePatterns()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "   \n\t\n");

        Assert.True(matcher.Enabled);
        Assert.False(matcher.HasActivePatterns);
    }

    [Fact]
    public void ExactMatch_IsCaseInsensitive()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "x-custom-internal");

        Assert.Equal(new[] { "X-Custom-Internal" },
            matcher.Match(new[] { "X-Custom-Internal", "CF_TEXT" }));
    }

    [Fact]
    public void WildcardMatchesPrefix()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "x-custom-*");

        Assert.Equal(new[] { "x-custom-alpha", "x-custom-beta" },
            matcher.Match(new[] { "x-custom-alpha", "CF_TEXT", "x-custom-beta", "other" }));
    }

    [Fact]
    public void WildcardMatchesSuffix()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "*-marker");

        Assert.Equal(new[] { "alpha-marker", "beta-marker" },
            matcher.Match(new[] { "alpha-marker", "no-match", "beta-marker" }));
    }

    [Fact]
    public void WildcardMatchesInfix()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "x-*-internal");

        Assert.Equal(new[] { "x-custom-internal", "x-foo-internal" },
            matcher.Match(new[] { "x-custom-internal", "x-foo-internal", "x-other" }));
    }

    [Fact]
    public void StarOnly_MatchesEverything()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "*");

        Assert.Equal(new[] { "CF_TEXT", "CF_UNICODETEXT" },
            matcher.Match(new[] { "CF_TEXT", "CF_UNICODETEXT" }));
    }

    [Fact]
    public void MultiplePatterns_OnePerLine_AreAllApplied()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "CF_TEXT\nx-custom-*\n# not a comment, just literal\n*-marker");

        Assert.Equal(new[] { "CF_TEXT", "x-custom-alpha", "end-marker" },
            matcher.Match(new[] { "CF_TEXT", "x-custom-alpha", "irrelevant", "end-marker" }));
    }

    [Fact]
    public void DuplicateFormatNames_AreReportedOnce()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "CF_TEXT");

        Assert.Equal(new[] { "CF_TEXT" },
            matcher.Match(new[] { "CF_TEXT", "CF_TEXT", "cf_text" }));
    }

    [Fact]
    public void FormatNameMatchedByMultiplePatterns_ReportedOnce()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "x-custom-*\n*-internal");

        Assert.Equal(new[] { "x-custom-internal" },
            matcher.Match(new[] { "x-custom-internal" }));
    }

    [Fact]
    public void PatternsWithSurroundingWhitespace_AreTrimmed()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "  CF_TEXT  \n\t x-custom-*\t");

        Assert.Equal(new[] { "CF_TEXT", "x-custom-foo" },
            matcher.Match(new[] { "CF_TEXT", "x-custom-foo" }));
    }

    [Fact]
    public void Reconfigure_ReplacesPreviousState()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "CF_TEXT");
        matcher.Configure(enabled: true, "CF_BITMAP");

        Assert.Empty(matcher.Match(new[] { "CF_TEXT" }));
        Assert.Equal(new[] { "CF_BITMAP" }, matcher.Match(new[] { "CF_BITMAP" }));
    }

    [Fact]
    public void EmptyOrNullFormatName_IsIgnored()
    {
        var matcher = new FormatNotificationMatcher();
        matcher.Configure(enabled: true, "*");

        Assert.Equal(new[] { "CF_TEXT" }, matcher.Match(new[] { "", null!, "CF_TEXT" }));
    }
}
