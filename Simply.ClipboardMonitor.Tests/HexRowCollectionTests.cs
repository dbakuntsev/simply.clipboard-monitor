using Simply.ClipboardMonitor.Common;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

public class HexRowCollectionTests
{
    // ── Count ───────────────────────────────────────────────────────────────

    [Fact] public void Count_EmptyArray_IsZero()
        => Assert.Empty(new HexRowCollection([]));

    [Fact] public void Count_ExactlyOneRow_IsOne()
        => Assert.Single(new HexRowCollection(new byte[16]));

    [Fact] public void Count_OneByteOverRow_IsTwo()
        => Assert.Equal(2, new HexRowCollection(new byte[17]).Count);

    [Fact] public void Count_TwoFullRows_IsTwo()
        => Assert.Equal(2, new HexRowCollection(new byte[32]).Count);

    // ── Offset ──────────────────────────────────────────────────────────────

    [Fact] public void FirstRow_Offset_IsZero()
        => Assert.Equal("00000000", new HexRowCollection([0x41])[0].Offset);

    [Fact] public void SecondRow_Offset_IsSixteenHex()
        => Assert.Equal("00000010", new HexRowCollection(new byte[17])[1].Offset);

    // ── Hex content ─────────────────────────────────────────────────────────

    [Fact]
    public void FullRow_Hex_HasSixteenSpaceSeparatedPairs()
    {
        var data  = Enumerable.Range(0, 16).Select(i => (byte)i).ToArray();
        var parts = new HexRowCollection(data)[0].Hex.Split(' ');
        Assert.Equal(16, parts.Length);
        Assert.Equal("00", parts[0]);
        Assert.Equal("0F", parts[15]);
    }

    [Fact]
    public void PartialRow_Hex_PadsShortRowWithSpaces()
    {
        // 17-byte input: row 1 has only 1 byte; the remaining 15 slots are "  " padding.
        var data = Enumerable.Range(0, 17).Select(i => (byte)i).ToArray();
        var row1 = new HexRowCollection(data)[1].Hex;
        Assert.StartsWith("10", row1);             // byte value 0x10
        Assert.EndsWith("  ", row1);               // last padded slot
    }

    [Fact]
    public void SingleByte_Hex_StartsWithByteValue()
        => Assert.StartsWith("AB", new HexRowCollection([0xAB])[0].Hex);

    // ── ASCII content ────────────────────────────────────────────────────────

    [Fact] public void Ascii_PrintableChar_ShowsCharacter()
        => Assert.StartsWith("A", new HexRowCollection([(byte)'A'])[0].Ascii);

    [Fact] public void Ascii_NonPrintableByte_ShowsDot()
        => Assert.StartsWith(".", new HexRowCollection([0x01])[0].Ascii);

    [Fact] public void Ascii_SpaceByte_ShowsSpace()
        => Assert.StartsWith(" ", new HexRowCollection([0x20])[0].Ascii);

    [Fact] public void Ascii_TildeByte_ShowsTilde()
        => Assert.StartsWith("~", new HexRowCollection([0x7E])[0].Ascii);  // 126 = printable

    [Fact] public void Ascii_DelByte_ShowsDot()
        => Assert.StartsWith(".", new HexRowCollection([0x7F])[0].Ascii);  // 127 = non-printable

    // ── Indexer bounds ───────────────────────────────────────────────────────

    [Fact]
    public void Indexer_NegativeIndex_Throws()
        => Assert.Throws<ArgumentOutOfRangeException>(() => new HexRowCollection([0x00])[-1]);

    [Fact]
    public void Indexer_IndexEqualToCount_Throws()
        => Assert.Throws<ArgumentOutOfRangeException>(() => new HexRowCollection([0x00])[1]);

    // ── Caching ──────────────────────────────────────────────────────────────

    [Fact]
    public void Row_AccessedTwice_ReturnsSameInstance()
    {
        var col = new HexRowCollection([0x41, 0x42]);
        Assert.Same(col[0], col[0]);
    }

    // ── Enumeration ──────────────────────────────────────────────────────────

    [Fact]
    public void Enumerate_YieldsSameInstancesAsIndexer()
    {
        var data = Enumerable.Range(0, 33).Select(i => (byte)i).ToArray();
        var col  = new HexRowCollection(data);
        var list = col.ToList();

        Assert.Equal(col.Count, list.Count);
        for (var i = 0; i < col.Count; i++)
            Assert.Same(col[i], list[i]);
    }
}
