using Simply.ClipboardMonitor.Common;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

public class DisplayHelperTests
{
    private const long OneKiB = 1024;
    private const long OneAndAHalfKiB = OneKiB * 3 / 2;

    [Fact]
    public void FormatFileSize_Zero_ReturnsNotCreatedYet()
        => Assert.Equal("Not created yet", DisplayHelper.FormatFileSize(0));

    [Fact]
    public void FormatFileSize_SmallBytes_ReturnsBytesWithBSuffix()
        => Assert.Equal("512 B", DisplayHelper.FormatFileSize(512));

    [Fact]
    public void FormatFileSize_OneByte_SingularForm()
        => Assert.Equal("1 B", DisplayHelper.FormatFileSize(1));

    [Fact]
    public void FormatFileSize_1023Bytes_UsesBSuffix()
        => Assert.Equal("1023 B", DisplayHelper.FormatFileSize(1023));

    [Theory]
    [InlineData(OneKiB,                   "1.0 KB")]
    [InlineData(OneKiB * OneKiB,          "1.0 MB")]
    [InlineData(OneKiB * OneKiB * OneKiB, "1.0 GB")]
    public void FormatFileSize_ExactPowerOfTwo_FormatsCorrectly(long bytes, string expected)
        => Assert.Equal(expected, DisplayHelper.FormatFileSize(bytes));

    // Verifies that rounding is correct (not truncating) across all three scaled ranges.
    [Theory]
    [InlineData(OneAndAHalfKiB,                   "1.5 KB")]  // 1.5 × 1 KiB
    [InlineData(OneAndAHalfKiB * OneKiB,          "1.5 MB")]  // 1.5 × 1 MiB
    [InlineData(OneAndAHalfKiB * OneKiB * OneKiB, "1.5 GB")]  // 1.5 × 1 GiB
    public void FormatFileSize_HalfwayPoint_ProducesOneDecimal(long bytes, string expected)
        => Assert.Equal(expected, DisplayHelper.FormatFileSize(bytes));

    // Verifies that rounding is correct (not truncating) across all three scaled ranges.
    [Theory]
    [InlineData((OneAndAHalfKiB - 1),                   "1.5 KB")]  // 1.5 × 1 KiB
    [InlineData((OneAndAHalfKiB - 1) * OneKiB,          "1.5 MB")]  // 1.5 × 1 MiB
    [InlineData((OneAndAHalfKiB - 1) * OneKiB * OneKiB, "1.5 GB")]  // 1.5 × 1 GiB
    public void FormatFileSize_HalfwayPointMinusOne_RoundsToOneDecimal(long bytes, string expected)
        => Assert.Equal(expected, DisplayHelper.FormatFileSize(bytes));

    // Verifies that rounding is correct (not truncating) across all three scaled ranges.
    [Theory]
    [InlineData((OneAndAHalfKiB + 1),                   "1.5 KB")]  // 1.5 × 1 KiB
    [InlineData((OneAndAHalfKiB + 1) * OneKiB,          "1.5 MB")]  // 1.5 × 1 MiB
    [InlineData((OneAndAHalfKiB + 1) * OneKiB * OneKiB, "1.5 GB")]  // 1.5 × 1 GiB
    public void FormatFileSize_HalfwayPointPlusOne_RoundsToOneDecimal(long bytes, string expected)
        => Assert.Equal(expected, DisplayHelper.FormatFileSize(bytes));

    [Fact]
    public void FormatFileSize_GbRange_UsesGbSuffix()
    {
        var result = DisplayHelper.FormatFileSize(2L * OneKiB * OneKiB * OneKiB);
        Assert.EndsWith("GB", result);
    }
}
