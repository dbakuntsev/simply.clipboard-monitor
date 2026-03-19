using Microsoft.Data.Sqlite;
using System.IO;
using System.Security.Cryptography;
using ZstdSharp;

namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// One clipboard format entry captured at a point in time, with its raw bytes.
/// </summary>
internal sealed record FormatSnapshot(
    int     Ordinal,
    uint    FormatId,
    string  FormatName,
    string  HandleType,
    /// <summary>Raw (uncompressed) bytes; null for handle types with no readable data.</summary>
    byte[]? Data,
    long    OriginalSize);

/// <summary>
/// A single row from the sessions table, used to populate the history list.
/// </summary>
internal sealed record SessionEntry(
    long     SessionId,
    DateTime Timestamp,
    string   FormatsText,
    long     TotalSize);

/// <summary>
/// Persists clipboard change history to a local SQLite database (history.db).
///
/// Blob data is stored compressed (ZStandard, max level) and deduplicated by SHA-256
/// hash so that identical payloads across different sessions are stored only once.
/// </summary>
internal static class ClipboardHistory
{
    private const int ZstdMaxLevel = 22;

    private static string DbPath =>
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
    public static (long SessionId, bool Trimmed) AddSession(
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
    public static long GetDatabaseFileSize()
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
    public static bool EnforceLimits(int maxEntries, long maxDatabaseBytes)
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
    public static void ClearHistory()
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
    public static List<SessionEntry> LoadSessions(int maxCount = 2000)
    {
        if (!File.Exists(DbPath))
            return [];

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = """
                SELECT id, timestamp, formats_text, total_size
                FROM   sessions
                ORDER  BY id DESC
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
                result.Add(new SessionEntry(sessionId, timestamp, formatsText, totalSize));
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
    public static List<FormatSnapshot> LoadSessionFormats(long sessionId)
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

    private static void CreateSchema(SqliteConnection conn)
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

    private static long InsertSession(SqliteConnection conn, DateTime timestamp, string formatsText, long totalSize)
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

    private static long EnsureClipboardFormat(SqliteConnection conn, uint formatId, string formatName)
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

    private static long EnsureContent(SqliteConnection conn, string hash, byte[] data, long originalSize)
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

    private static void InsertSessionItem(SqliteConnection conn, long sessionId, int ordinal,
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

    private static string ComputeHash(byte[] data)
    {
        Span<byte> hash = stackalloc byte[32];
        SHA256.HashData(data, hash);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    internal static string BuildFormatsText(IReadOnlyList<FormatSnapshot> snapshots)
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
    private static bool TrimToLimits(SqliteConnection conn, int maxEntries, long maxDatabaseBytes)
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

    private static long GetSessionCount(SqliteConnection conn)
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
    private static long GetStoredBlobBytes(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(SUM(LENGTH(data)), 0) FROM clipboard_contents";
        return (long)cmd.ExecuteScalar()!;
    }

    private static void DeleteOldestSessions(SqliteConnection conn, int count)
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

    private static void DeleteOrphanedContent(SqliteConnection conn)
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

    private static SqliteConnection OpenConnection(bool readOnly)
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
