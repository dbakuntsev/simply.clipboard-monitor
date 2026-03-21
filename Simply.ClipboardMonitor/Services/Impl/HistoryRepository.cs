using Microsoft.Data.Sqlite;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;
using System.IO;
using System.Security.Cryptography;
using ZstdSharp;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Persists clipboard change history to a local SQLite database (history.db).
///
/// Blob data is stored compressed (ZStandard, max level) and deduplicated by SHA-256
/// hash so that identical payloads across different sessions are stored only once.
/// </summary>
internal sealed class HistoryRepository : IHistoryRepository, IHistoryMaintenance
{
    private const int ZstdMaxLevel = 22;

    private string DbPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Simply.ClipboardMonitor",
            "history.db");

    // ── Public API ──────────────────────────────────────────────────────────

    /// <summary>
    /// Inserts a new session and all its format snapshots into the database,
    /// then trims the oldest sessions until the entry count and approximate
    /// database size are both within the specified limits.
    /// Intended to be called from a background thread.
    /// </summary>
    /// <returns>
    /// The new session's row ID and a flag indicating whether any older sessions
    /// were deleted by <see cref="TrimToLimits"/> (i.e. the history list needs refresh).
    /// </returns>
    public (long SessionId, bool Trimmed) AddSession(
        IReadOnlyList<FormatSnapshot> snapshots,
        DateTime timestamp,
        int  maxEntries,
        long maxDatabaseBytes)
    {
        try
        {
            using var conn = OpenConnection(readOnly: false);
            CreateSchema(conn);

            long sessionId;
            using (var tx = conn.BeginTransaction())
            {
                try
                {
                    var formatsText = BuildFormatsText(snapshots);
                    var totalSize   = snapshots.Sum(s => s.OriginalSize);
                    sessionId       = InsertSession(conn, timestamp, formatsText, totalSize);

                    foreach (var snap in snapshots)
                    {
                        var formatDbId = EnsureClipboardFormat(conn, snap.FormatId, snap.FormatName);
                        long? contentId = null;

                        if (snap.Data is { Length: > 0 })
                        {
                            var hash = ComputeHash(snap.Data);
                            contentId = EnsureContent(conn, hash, snap.Data, snap.OriginalSize);
                        }

                        InsertSessionItem(conn, sessionId, snap.Ordinal, formatDbId, contentId, snap.HandleType);
                    }

                    tx.Commit();
                }
                catch
                {
                    tx.Rollback();
                    throw;
                }
            }

            // Trim outside the write transaction; VACUUM cannot run inside one.
            var trimmed = TrimToLimits(conn, maxEntries, maxDatabaseBytes);
            if (trimmed)
            {
                using var vac = conn.CreateCommand();
                vac.CommandText = "VACUUM";
                vac.ExecuteNonQuery();
            }

            return (sessionId, trimmed);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>Returns the history database file size in bytes, or 0 if it does not exist.</summary>
    public long GetDatabaseFileSize()
    {
        var path = DbPath;
        return File.Exists(path) ? new FileInfo(path).Length : 0L;
    }

    /// <summary>
    /// Applies <paramref name="maxEntries"/> and <paramref name="maxDatabaseBytes"/> limits
    /// to an existing database, then VACUUMs if anything was removed.
    /// Returns <see langword="true"/> if any sessions were deleted.
    /// No-op (returns <see langword="false"/>) when the database does not exist.
    /// Intended to be called from a background thread.
    /// </summary>
    public bool EnforceLimits(int maxEntries, long maxDatabaseBytes)
    {
        if (!File.Exists(DbPath))
            return false;

        try
        {
            using var conn = OpenConnection(readOnly: false);
            var removed = TrimToLimits(conn, maxEntries, maxDatabaseBytes);
            if (removed)
            {
                using var vac = conn.CreateCommand();
                vac.CommandText = "VACUUM";
                vac.ExecuteNonQuery();
            }
            return removed;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>
    /// Deletes all sessions, items, and content blobs, then compacts the file with VACUUM.
    /// Safe to call when the database does not yet exist.
    /// </summary>
    public void ClearHistory()
    {
        if (!File.Exists(DbPath))
            return;

        try
        {
            using var conn = OpenConnection(readOnly: false);

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    DELETE FROM session_items;
                    DELETE FROM sessions;
                    DELETE FROM clipboard_contents;
                    """;
                cmd.ExecuteNonQuery();
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "VACUUM";
                cmd.ExecuteNonQuery();
            }
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>
    /// Returns all sessions ordered newest-first (up to <paramref name="maxCount"/>).
    /// Returns an empty list if the database does not exist yet.
    /// </summary>
    public List<SessionEntry> LoadSessions(int maxCount = 2000)
    {
        if (!File.Exists(DbPath))
            return [];

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            // GROUP_CONCAT produces "formatId<TAB>formatName" pairs joined by newlines,
            // one per format in the session (ordered by item ordinal via subquery).
            cmd.CommandText = """
                SELECT s.id, s.timestamp, s.formats_text, s.total_size,
                       GROUP_CONCAT(CAST(cf.format_id AS TEXT) || char(9) || cf.format_name, char(10))
                FROM   sessions s
                LEFT JOIN session_items    si ON si.session_id          = s.id
                LEFT JOIN clipboard_formats cf ON cf.id                 = si.clipboard_format_id
                GROUP  BY s.id
                ORDER  BY s.id DESC
                LIMIT  @limit
                """;
            cmd.Parameters.AddWithValue("@limit", maxCount);

            var result = new List<SessionEntry>();
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var sessionId   = reader.GetInt64(0);
                var timestamp   = DateTime.Parse(reader.GetString(1));
                var formatsText = reader.GetString(2);
                var totalSize   = reader.GetInt64(3);

                // Parse "id\tname\nid\tname\n..." into a typed list.
                var formats = new List<(uint FormatId, string FormatName)>();
                if (!reader.IsDBNull(4))
                {
                    foreach (var line in reader.GetString(4).Split('\n', StringSplitOptions.RemoveEmptyEntries))
                    {
                        var tab = line.IndexOf('\t');
                        if (tab > 0 && uint.TryParse(line[..tab], out var fid))
                            formats.Add((fid, line[(tab + 1)..]));
                    }
                }

                result.Add(new SessionEntry(sessionId, timestamp, formatsText, totalSize, formats));
            }

            return result;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>
    /// Returns all format snapshots for a given session, with blobs decompressed.
    /// Returns an empty list if the database does not exist.
    /// </summary>
    public List<FormatSnapshot> LoadSessionFormats(long sessionId)
    {
        if (!File.Exists(DbPath))
            return [];

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = """
                SELECT si.ordinal, cf.format_id, cf.format_name, si.handle_type,
                       cc.data, cc.original_size
                FROM   session_items si
                JOIN   clipboard_formats  cf ON si.clipboard_format_id  = cf.id
                LEFT JOIN clipboard_contents cc ON si.clipboard_content_id = cc.id
                WHERE  si.session_id = @sessionId
                ORDER  BY si.ordinal
                """;
            cmd.Parameters.AddWithValue("@sessionId", sessionId);

            var result = new List<FormatSnapshot>();
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var ordinal      = reader.GetInt32(0);
                var formatId     = (uint)reader.GetInt64(1);
                var formatName   = reader.GetString(2);
                var handleType   = reader.GetString(3);
                byte[]? data     = null;
                long originalSize = 0;

                if (!reader.IsDBNull(4))
                {
                    var compressed = reader.GetFieldValue<byte[]>(4);
                    using var decompressor = new Decompressor();
                    data = decompressor.Unwrap(compressed).ToArray();
                }

                if (!reader.IsDBNull(5))
                    originalSize = reader.GetInt64(5);

                result.Add(new FormatSnapshot(ordinal, formatId, formatName, handleType, data, originalSize));
            }

            return result;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    // ── Schema ──────────────────────────────────────────────────────────────

    private void CreateSchema(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            PRAGMA journal_mode = DELETE;
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS sessions (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp    TEXT    NOT NULL,
                formats_text TEXT    NOT NULL,
                total_size   INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS clipboard_formats (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                format_id   INTEGER NOT NULL,
                format_name TEXT    NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS clipboard_contents (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                data          BLOB,
                content_hash  TEXT    NOT NULL UNIQUE,
                original_size INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS session_items (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id           INTEGER NOT NULL REFERENCES sessions(id),
                ordinal              INTEGER NOT NULL,
                clipboard_format_id  INTEGER NOT NULL REFERENCES clipboard_formats(id),
                clipboard_content_id INTEGER          REFERENCES clipboard_contents(id),
                handle_type          TEXT    NOT NULL
            );
            """;
        cmd.ExecuteNonQuery();
    }

    // ── Write helpers ────────────────────────────────────────────────────────

    private long InsertSession(SqliteConnection conn, DateTime timestamp, string formatsText, long totalSize)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO sessions (timestamp, formats_text, total_size)
            VALUES (@timestamp, @formatsText, @totalSize);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@timestamp",   timestamp.ToString("O"));
        cmd.Parameters.AddWithValue("@formatsText", formatsText);
        cmd.Parameters.AddWithValue("@totalSize",   totalSize);
        return (long)cmd.ExecuteScalar()!;
    }

    private long EnsureClipboardFormat(SqliteConnection conn, uint formatId, string formatName)
    {
        using var check = conn.CreateCommand();
        check.CommandText = "SELECT id FROM clipboard_formats WHERE format_name = @name";
        check.Parameters.AddWithValue("@name", formatName);
        var existing = check.ExecuteScalar();
        if (existing != null)
            return (long)existing;

        using var insert = conn.CreateCommand();
        insert.CommandText = """
            INSERT INTO clipboard_formats (format_id, format_name) VALUES (@fid, @name);
            SELECT last_insert_rowid();
            """;
        insert.Parameters.AddWithValue("@fid",  (long)formatId);
        insert.Parameters.AddWithValue("@name", formatName);
        return (long)insert.ExecuteScalar()!;
    }

    private long EnsureContent(SqliteConnection conn, string hash, byte[] data, long originalSize)
    {
        using var check = conn.CreateCommand();
        check.CommandText = "SELECT id FROM clipboard_contents WHERE content_hash = @hash";
        check.Parameters.AddWithValue("@hash", hash);
        var existing = check.ExecuteScalar();
        if (existing != null)
            return (long)existing;

        using var compressor = new Compressor(ZstdMaxLevel);
        var compressed = compressor.Wrap(data).ToArray();

        using var insert = conn.CreateCommand();
        insert.CommandText = """
            INSERT INTO clipboard_contents (data, content_hash, original_size)
            VALUES (@data, @hash, @originalSize);
            SELECT last_insert_rowid();
            """;
        insert.Parameters.AddWithValue("@data",         compressed);
        insert.Parameters.AddWithValue("@hash",         hash);
        insert.Parameters.AddWithValue("@originalSize", originalSize);
        return (long)insert.ExecuteScalar()!;
    }

    private void InsertSessionItem(SqliteConnection conn, long sessionId, int ordinal,
        long formatDbId, long? contentId, string handleType)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO session_items
                (session_id, ordinal, clipboard_format_id, clipboard_content_id, handle_type)
            VALUES (@sessionId, @ordinal, @formatId, @contentId, @handleType)
            """;
        cmd.Parameters.AddWithValue("@sessionId",  sessionId);
        cmd.Parameters.AddWithValue("@ordinal",    ordinal);
        cmd.Parameters.AddWithValue("@formatId",   formatDbId);
        cmd.Parameters.AddWithValue("@contentId",  (object?)contentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@handleType", handleType);
        cmd.ExecuteNonQuery();
    }

    // ── Utilities ────────────────────────────────────────────────────────────

    private string ComputeHash(byte[] data)
    {
        Span<byte> hash = stackalloc byte[32];
        SHA256.HashData(data, hash);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public string BuildFormatsText(IReadOnlyList<FormatSnapshot> snapshots)
    {
        var joined = string.Join(", ", snapshots.Select(s => s.FormatName));
        return joined.Length > 80 ? joined[..77] + "..." : joined;
    }

    // ── Limit enforcement ────────────────────────────────────────────────────

    /// <summary>
    /// Deletes oldest sessions (and orphaned content blobs) until both the
    /// session count and the approximate database size are within limits.
    /// Returns true if anything was deleted (caller should VACUUM afterwards).
    /// </summary>
    private bool TrimToLimits(SqliteConnection conn, int maxEntries, long maxDatabaseBytes)
    {
        var deleted = false;

        // --- count limit ---
        if (maxEntries > 0)
        {
            var count = GetSessionCount(conn);
            if (count > maxEntries)
            {
                using var tx = conn.BeginTransaction();
                DeleteOldestSessions(conn, (int)(count - maxEntries));
                DeleteOrphanedContent(conn);
                tx.Commit();
                deleted = true;
            }
        }

        // --- size limit ---
        if (maxDatabaseBytes > 0)
        {
            while (GetStoredBlobBytes(conn) > maxDatabaseBytes)
            {
                if (GetSessionCount(conn) <= 1) break;  // always keep at least 1 session

                using var tx = conn.BeginTransaction();
                DeleteOldestSessions(conn, 1);
                DeleteOrphanedContent(conn);
                tx.Commit();
                deleted = true;
            }
        }

        return deleted;
    }

    private long GetSessionCount(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM sessions";
        return (long)cmd.ExecuteScalar()!;
    }

    /// <summary>
    /// Returns the total number of compressed bytes stored in <c>clipboard_contents.data</c>.
    /// Decreases immediately after a DELETE + COMMIT, making it reliable for size-limit
    /// enforcement without needing a VACUUM first.
    /// </summary>
    private long GetStoredBlobBytes(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(SUM(LENGTH(data)), 0) FROM clipboard_contents";
        return (long)cmd.ExecuteScalar()!;
    }

    private void DeleteOldestSessions(SqliteConnection conn, int count)
    {
        using var delItems = conn.CreateCommand();
        delItems.CommandText = """
            DELETE FROM session_items
            WHERE session_id IN (SELECT id FROM sessions ORDER BY id ASC LIMIT @n)
            """;
        delItems.Parameters.AddWithValue("@n", count);
        delItems.ExecuteNonQuery();

        using var delSessions = conn.CreateCommand();
        delSessions.CommandText = """
            DELETE FROM sessions
            WHERE id IN (SELECT id FROM sessions ORDER BY id ASC LIMIT @n)
            """;
        delSessions.Parameters.AddWithValue("@n", count);
        delSessions.ExecuteNonQuery();
    }

    private void DeleteOrphanedContent(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            DELETE FROM clipboard_contents
            WHERE id NOT IN (
                SELECT DISTINCT clipboard_content_id
                FROM   session_items
                WHERE  clipboard_content_id IS NOT NULL
            )
            """;
        cmd.ExecuteNonQuery();
    }

    private SqliteConnection OpenConnection(bool readOnly)
    {
        var path = DbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        var csb = new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode       = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
        };
        var conn = new SqliteConnection(csb.ToString());
        conn.Open();
        return conn;
    }
}
