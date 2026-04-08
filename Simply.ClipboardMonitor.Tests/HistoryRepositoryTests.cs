using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services.Impl;
using System.IO;
using System.Text;
using Xunit;

namespace Simply.ClipboardMonitor.Tests;

/// <summary>
/// Integration tests for <see cref="HistoryRepository"/> using a temporary SQLite
/// database file per test class instance.  Each test class instance gets its own
/// isolated directory, which is deleted in <see cref="Dispose"/>.
/// </summary>
public sealed class HistoryRepositoryTests : IDisposable
{
    private readonly string            _tempDir;
    private readonly HistoryRepository _sut;

    // ── String constants used across multiple tests ──────────────────────────
    private const string PillsTxt            = "TXT";
    private const string FormatNameUnicode   = "CF_UNICODETEXT";
    private const string HandleTypeHGlobal   = "hglobal";

    public HistoryRepositoryTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"SCM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _sut = HistoryRepository.CreateForTesting(Path.Combine(_tempDir, "history.db"));
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); }
        catch { /* best-effort cleanup */ }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    /// <summary>Builds a single-format snapshot list backed by UTF-16 text bytes.</summary>
    private static IReadOnlyList<FormatSnapshot> TextSnap(string text, string formatName = FormatNameUnicode)
    {
        var bytes = Encoding.Unicode.GetBytes(text);
        return [new FormatSnapshot(1, 13u, formatName, HandleTypeHGlobal, bytes, bytes.Length)];
    }

    private long AddOne(string text = "Hello", string formatName = FormatNameUnicode)
    {
        var (id, _) = _sut.AddSession(TextSnap(text, formatName), [text], PillsTxt, DateTime.UtcNow, 0, 0);
        return id;
    }

    // ── GetSessionCount ──────────────────────────────────────────────────────

    [Fact]
    public void GetSessionCount_DatabaseNotYetCreated_ReturnsZero()
        => Assert.Equal(0, _sut.GetSessionCount());

    [Fact]
    public void GetSessionCount_AfterAddingTwoSessions_ReturnsTwo()
    {
        AddOne("First");
        AddOne("Second");
        Assert.Equal(2, _sut.GetSessionCount());
    }

    // ── AddSession / LoadSessions ────────────────────────────────────────────

    [Fact]
    public void AddSession_FirstSession_ReturnsTrimmedFalse()
    {
        var (_, trimmed) = _sut.AddSession(TextSnap("Hello"), ["Hello"], PillsTxt, DateTime.UtcNow, 0, 0);
        Assert.False(trimmed);
    }

    [Fact]
    public void AddSession_ExceedsEntryLimit_ReturnsTrimmedTrue()
    {
        AddOne("A");
        AddOne("B");
        // Adding a third session while limit is 2 should trim and report Trimmed=true.
        var (_, trimmed) = _sut.AddSession(TextSnap("C"), ["C"], PillsTxt, DateTime.UtcNow, maxEntries: 2, maxDatabaseBytes: 0);
        Assert.True(trimmed);
        Assert.Equal(2, _sut.GetSessionCount());
    }

    [Fact]
    public void LoadSessions_AfterOneAdd_ReturnsSingleEntry()
    {
        AddOne();
        Assert.Single(_sut.LoadSessions());
    }

    [Fact]
    public void LoadSessions_DefaultOrder_NewestFirst()
    {
        var id1 = AddOne("First");
        var id2 = AddOne("Second");
        var sessions = _sut.LoadSessions();

        Assert.Equal(id2, sessions[0].SessionId);
        Assert.Equal(id1, sessions[1].SessionId);
    }

    [Fact]
    public void LoadSessions_DatabaseNotYetCreated_ReturnsEmpty()
        => Assert.Empty(_sut.LoadSessions());

    [Fact]
    public void LoadSessions_WithMatchingFilter_ReturnsOnlyMatchingSessions()
    {
        AddOne("needle in haystack");
        AddOne("something else entirely");

        var results = _sut.LoadSessions("needle");
        Assert.Single(results);
    }

    [Fact]
    public void LoadSessions_FilterNoMatch_ReturnsEmpty()
    {
        AddOne("hello");
        Assert.Empty(_sut.LoadSessions("zzznomatch"));
    }

    // ── LoadSessionFormats ───────────────────────────────────────────────────

    [Fact]
    public void LoadSessionFormats_ReturnsCorrectFormatForSession()
    {
        var id      = AddOne("test content");
        var formats = _sut.LoadSessionFormats(id);

        Assert.Single(formats);
        Assert.Equal(FormatNameUnicode, formats[0].FormatName);
    }

    [Fact]
    public void LoadSessionFormats_RoundTripsRawBytes()
    {
        var original = "Round-trip test"u8.ToArray();
        var snaps    = new[] { new FormatSnapshot(1, 13u, FormatNameUnicode, HandleTypeHGlobal, original, original.Length) };
        var (id, _)  = _sut.AddSession(snaps, [null], PillsTxt, DateTime.UtcNow, 0, 0);

        var loaded = _sut.LoadSessionFormats(id);

        Assert.Equal(original, loaded[0].Data);
    }

    [Fact]
    public void LoadSessionFormats_NonExistentSessionId_ReturnsEmpty()
    {
        AddOne();
        Assert.Empty(_sut.LoadSessionFormats(long.MaxValue));
    }

    [Fact]
    public void LoadSessionFormats_DatabaseNotYetCreated_ReturnsEmpty()
        => Assert.Empty(_sut.LoadSessionFormats(1));

    // ── DeleteSession ────────────────────────────────────────────────────────

    [Fact]
    public void DeleteSession_RemovesSessionFromCount()
    {
        var id = AddOne();
        _sut.DeleteSession(id);
        Assert.Equal(0, _sut.GetSessionCount());
    }

    [Fact]
    public void DeleteSession_FormatsNoLongerReturned()
    {
        var id = AddOne();
        _sut.DeleteSession(id);
        Assert.Empty(_sut.LoadSessionFormats(id));
    }

    [Fact]
    public void DeleteSession_OtherSessionsUnaffected()
    {
        var id1 = AddOne("keep");
        var id2 = AddOne("delete");

        _sut.DeleteSession(id2);

        Assert.Equal(1, _sut.GetSessionCount());
        Assert.Single(_sut.LoadSessionFormats(id1));
    }

    [Fact]
    public void DeleteSession_SharedContentBlob_NotDeletedWhileOtherSessionReferencesIt()
    {
        // Two sessions with identical data share one content blob.
        // Deleting one session must not remove the blob the other still needs.
        const string sharedText = "shared content";
        var snaps    = TextSnap(sharedText);
        var (id1, _) = _sut.AddSession(snaps, [sharedText], PillsTxt, DateTime.UtcNow, 0, 0);
        var (id2, _) = _sut.AddSession(snaps, [sharedText], PillsTxt, DateTime.UtcNow, 0, 0);

        _sut.DeleteSession(id1);

        var formats = _sut.LoadSessionFormats(id2);
        Assert.Single(formats);
        Assert.NotNull(formats[0].Data); // blob still present
    }

    // ── ClearHistory ────────────────────────────────────────────────────────

    [Fact]
    public void ClearHistory_RemovesAllSessions()
    {
        AddOne("A");
        AddOne("B");
        AddOne("C");

        _sut.ClearHistory();

        Assert.Equal(0, _sut.GetSessionCount());
        Assert.Empty(_sut.LoadSessions());
    }

    // ── IsDuplicateOfLastSession ─────────────────────────────────────────────

    [Fact]
    public void IsDuplicateOfLastSession_DatabaseNotYetCreated_ReturnsFalse()
        => Assert.False(_sut.IsDuplicateOfLastSession(TextSnap("anything")));

    [Fact]
    public void IsDuplicateOfLastSession_IdenticalContent_ReturnsTrue()
    {
        var snaps = TextSnap("duplicate me");
        _sut.AddSession(snaps, ["duplicate me"], PillsTxt, DateTime.UtcNow, 0, 0);

        Assert.True(_sut.IsDuplicateOfLastSession(snaps));
    }

    [Fact]
    public void IsDuplicateOfLastSession_DifferentContent_ReturnsFalse()
    {
        _sut.AddSession(TextSnap("original"), ["original"], PillsTxt, DateTime.UtcNow, 0, 0);
        Assert.False(_sut.IsDuplicateOfLastSession(TextSnap("different")));
    }

    [Fact]
    public void IsDuplicateOfLastSession_DifferentFormatCount_ReturnsFalse()
    {
        _sut.AddSession(TextSnap("hello"), ["hello"], PillsTxt, DateTime.UtcNow, 0, 0);

        // Two formats vs the stored one-format session.
        var twoFormats = new[]
        {
            new FormatSnapshot(1, 13u, FormatNameUnicode, HandleTypeHGlobal,
                Encoding.Unicode.GetBytes("hello"), Encoding.Unicode.GetByteCount("hello")),
            new FormatSnapshot(2, 1u, "CF_TEXT", HandleTypeHGlobal,
                Encoding.Default.GetBytes("hello"), Encoding.Default.GetByteCount("hello")),
        };
        Assert.False(_sut.IsDuplicateOfLastSession(twoFormats));
    }

    // ── EnforceLimits ────────────────────────────────────────────────────────

    [Fact]
    public void EnforceLimits_OverEntryCount_TrimsOldestSessions()
    {
        var idOldest = AddOne("A");
        var idMiddle = AddOne("B");
        var idNewest = AddOne("C");

        _sut.EnforceLimits(maxEntries: 2, maxDatabaseBytes: 0);

        var remaining = _sut.LoadSessions().Select(s => s.SessionId).ToHashSet();
        Assert.DoesNotContain(idOldest, remaining);
        Assert.Contains(idMiddle,       remaining);
        Assert.Contains(idNewest,       remaining);
    }

    [Fact]
    public void EnforceLimits_DatabaseNotYetCreated_ReturnsFalse()
        => Assert.False(_sut.EnforceLimits(maxEntries: 1, maxDatabaseBytes: 0));

    // ── BuildFormatsText ─────────────────────────────────────────────────────

    [Fact]
    public void BuildFormatsText_SingleFormat_ReturnsName()
    {
        var snaps = TextSnap("x", FormatNameUnicode);
        Assert.Equal(FormatNameUnicode, _sut.BuildFormatsText(snaps));
    }

    [Fact]
    public void BuildFormatsText_LongNameList_TruncatesAtEightyChars()
    {
        var snaps = Enumerable.Range(1, 10)
            .Select(i => new FormatSnapshot(i, (uint)(50000 + i),
                $"VeryLongFormatNameNumber{i:D3}", HandleTypeHGlobal, null, 0))
            .ToArray();

        var text = _sut.BuildFormatsText(snaps);
        Assert.True(text.Length <= 80);
        Assert.EndsWith("...", text);
    }
}
