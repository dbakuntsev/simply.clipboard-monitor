using Microsoft.Data.Sqlite;
using Simply.ClipboardMonitor.Common;
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

    private readonly string? _dbPathOverride;

    private string DbPath => _dbPathOverride ??
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Simply.ClipboardMonitor",
            "history.db");

    // Parameterless constructor used by the DI container.
    public HistoryRepository() { }

    private HistoryRepository(string dbPathOverride) { _dbPathOverride = dbPathOverride; }

    /// <summary>Creates an instance backed by a specific database file path, for use in tests.</summary>
    internal static HistoryRepository CreateForTesting(string dbPath) => new(dbPath);

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
        IReadOnlyList<string?>        textContents,
        string                        pillsText,
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
                    sessionId       = InsertSession(conn, timestamp, formatsText, totalSize, pillsText);

                    for (var i = 0; i < snapshots.Count; i++)
                    {
                        var snap        = snapshots[i];
                        var textContent = i < textContents.Count ? textContents[i] : null;
                        var formatDbId  = EnsureClipboardFormat(conn, snap.FormatId, snap.FormatName);
                        long? contentId = null;

                        if (snap.Data is { Length: > 0 })
                        {
                            var hash = ComputeHash(snap.Data);
                            contentId = EnsureContent(conn, hash, snap.Data, snap.OriginalSize);
                        }

                        InsertSessionItem(conn, sessionId, snap.Ordinal, formatDbId, contentId, snap.HandleType, textContent);
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
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            return default;
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
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            return false;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>
    /// Opens the existing database (if any) and applies any pending schema migrations.
    /// No-op when the database does not yet exist.
    /// Intended to be called from a background thread on application startup.
    /// </summary>
    public void MigrateSchema()
    {
        if (!File.Exists(DbPath))
            return;

        try
        {
            using var conn = OpenConnection(readOnly: false);
            CreateSchema(conn);
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
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
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <inheritdoc/>
    public List<SessionEntry> LoadSessions(string? filter = null, int maxCount = 2000)
    {
        if (!File.Exists(DbPath))
            return [];

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();

            if (string.IsNullOrWhiteSpace(filter))
            {
                // Unfiltered: GROUP BY with GROUP_CONCAT for formats in one pass.
                cmd.CommandText = """
                    SELECT s.id, s.timestamp, s.formats_text, s.total_size,
                           GROUP_CONCAT(CAST(cf.format_id AS TEXT) || char(9) || cf.format_name, char(10))
                    FROM   sessions s
                    LEFT JOIN session_items     si ON si.session_id         = s.id
                    LEFT JOIN clipboard_formats cf ON cf.id                 = si.clipboard_format_id
                    GROUP  BY s.id
                    ORDER  BY s.id DESC
                    LIMIT  @limit
                    """;
            }
            else
            {
                // Filtered: DISTINCT rows matching the term; formats via correlated subquery.
                cmd.CommandText = """
                    SELECT DISTINCT s.id, s.timestamp, s.formats_text, s.total_size,
                        (SELECT GROUP_CONCAT(CAST(cf2.format_id AS TEXT) || char(9) || cf2.format_name, char(10))
                         FROM   session_items si2
                         JOIN   clipboard_formats cf2 ON cf2.id = si2.clipboard_format_id
                         WHERE  si2.session_id = s.id) AS formats_concat
                    FROM   sessions s
                    LEFT JOIN session_items     si ON si.session_id         = s.id
                    LEFT JOIN clipboard_formats cf ON cf.id                 = si.clipboard_format_id
                    WHERE  LOWER(s.timestamp)      LIKE @term
                       OR  LOWER(s.pills_text)     LIKE @term
                       OR  LOWER(cf.format_name)   LIKE @term
                       OR  LOWER(si.text_content)  LIKE @term
                    ORDER  BY s.id DESC
                    LIMIT  @limit
                    """;
                cmd.Parameters.AddWithValue("@term", $"%{filter.ToLowerInvariant()}%");
            }

            cmd.Parameters.AddWithValue("@limit", maxCount);

            var result = new List<SessionEntry>();
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var sessionId   = reader.GetInt64(0);
                var timestamp   = DateTime.Parse(reader.GetString(1));
                var formatsText = reader.GetString(2);
                var totalSize   = reader.GetInt64(3);

                // Parse "formatId<TAB>formatName\n..." into a typed list.
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
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            return [];
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
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            return [];
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
                total_size   INTEGER NOT NULL,
                pills_text   TEXT    NOT NULL DEFAULT ''
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
                handle_type          TEXT    NOT NULL,
                text_content         TEXT
            );
            """;
        cmd.ExecuteNonQuery();

        // Migrate existing databases that predate the search columns.
        AddColumnIfMissing(conn, "sessions",      "pills_text",    "TEXT NOT NULL DEFAULT ''");
        AddColumnIfMissing(conn, "session_items", "text_content",  "TEXT");
    }

    private static void AddColumnIfMissing(SqliteConnection conn, string table, string column, string definition)
    {
        using var info = conn.CreateCommand();
        info.CommandText = $"PRAGMA table_info({table})";
        using var reader = info.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase))
                return; // already exists
        }
        using var alter = conn.CreateCommand();
        alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {definition}";
        alter.ExecuteNonQuery();
    }

    // ── Write helpers ────────────────────────────────────────────────────────

    private long InsertSession(SqliteConnection conn, DateTime timestamp, string formatsText, long totalSize, string pillsText)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO sessions (timestamp, formats_text, total_size, pills_text)
            VALUES (@timestamp, @formatsText, @totalSize, @pillsText);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@timestamp",   timestamp.ToString("O"));
        cmd.Parameters.AddWithValue("@formatsText", formatsText);
        cmd.Parameters.AddWithValue("@totalSize",   totalSize);
        cmd.Parameters.AddWithValue("@pillsText",   pillsText);
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
        long formatDbId, long? contentId, string handleType, string? textContent)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO session_items
                (session_id, ordinal, clipboard_format_id, clipboard_content_id, handle_type, text_content)
            VALUES (@sessionId, @ordinal, @formatId, @contentId, @handleType, @textContent)
            """;
        cmd.Parameters.AddWithValue("@sessionId",   sessionId);
        cmd.Parameters.AddWithValue("@ordinal",     ordinal);
        cmd.Parameters.AddWithValue("@formatId",    formatDbId);
        cmd.Parameters.AddWithValue("@contentId",   (object?)contentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@handleType",  handleType);
        cmd.Parameters.AddWithValue("@textContent", (object?)textContent ?? DBNull.Value);
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

    /// <inheritdoc/>
    public int GetSessionCount()
    {
        if (!File.Exists(DbPath))
            return 0;
        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM sessions";
            return (int)(long)cmd.ExecuteScalar()!;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex);
            return 0;
        }
    }

    /// <inheritdoc/>
    public void DeleteSession(long sessionId)
    {
        if (!File.Exists(DbPath))
            return;

        try
        {
            using var conn = OpenConnection(readOnly: false);
            using (var tx = conn.BeginTransaction())
            {
                try
                {
                    using var delItems = conn.CreateCommand();
                    delItems.CommandText = "DELETE FROM session_items WHERE session_id = @id";
                    delItems.Parameters.AddWithValue("@id", sessionId);
                    delItems.ExecuteNonQuery();

                    using var delSession = conn.CreateCommand();
                    delSession.CommandText = "DELETE FROM sessions WHERE id = @id";
                    delSession.Parameters.AddWithValue("@id", sessionId);
                    delSession.ExecuteNonQuery();

                    DeleteOrphanedContent(conn);
                    tx.Commit();
                }
                catch
                {
                    tx.Rollback();
                    throw;
                }
            }
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <inheritdoc/>
    public bool IsDuplicateOfLastSession(IReadOnlyList<FormatSnapshot> snapshots)
    {
        if (!File.Exists(DbPath) || snapshots.Count == 0)
            return false;

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = """
                SELECT cf.format_name, cc.content_hash
                FROM   session_items si
                JOIN   clipboard_formats    cf ON cf.id = si.clipboard_format_id
                LEFT JOIN clipboard_contents cc ON cc.id = si.clipboard_content_id
                WHERE  si.session_id = (SELECT MAX(id) FROM sessions)
                ORDER  BY si.ordinal
                """;

            var lastItems = new List<(string Name, string? Hash)>();
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                    lastItems.Add((reader.GetString(0), reader.IsDBNull(1) ? null : reader.GetString(1)));
            }

            if (lastItems.Count != snapshots.Count)
                return false;

            for (var i = 0; i < snapshots.Count; i++)
            {
                var snap              = snapshots[i];
                var (lastName, lastHash) = lastItems[i];

                if (!string.Equals(snap.FormatName, lastName, StringComparison.Ordinal))
                    return false;

                var currentHash = snap.Data is { Length: > 0 } ? ComputeHash(snap.Data) : null;
                if (currentHash != lastHash)
                    return false;
            }

            return true;
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            return false;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
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

    private SqliteConnection OpenConnection(bool readOnly) =>
        OpenConnectionByPath(DbPath, readOnly);

    private static SqliteConnection OpenConnectionByPath(string path, bool readOnly)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var csb = new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode       = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
        };
        var conn = new SqliteConnection(csb.ToString());
        conn.Open();
        return conn;
    }

    // ── Integrity and recovery ───────────────────────────────────────────────

    /// <inheritdoc/>
    public DatabaseIntegrityStatus CheckIntegrity()
    {
        if (!File.Exists(DbPath))
            return DatabaseIntegrityStatus.Absent;

        try
        {
            using var conn = OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            // Bail after the first error so healthy large databases are not slow to check.
            cmd.CommandText = "PRAGMA integrity_check(1)";
            var first = cmd.ExecuteScalar()?.ToString();
            return string.Equals(first, "ok", StringComparison.OrdinalIgnoreCase)
                ? DatabaseIntegrityStatus.Healthy
                : DatabaseIntegrityStatus.Corrupted;
        }
        catch
        {
            return DatabaseIntegrityStatus.Corrupted;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <inheritdoc/>
    public RecoveryResult TryRecover()
    {
        // Strategy 1: VACUUM INTO a temp file — fastest when SQLite can read most pages.
        var vacuumResult = TryVacuumInto();
        if (vacuumResult.Success)
            return vacuumResult;

        // Strategies 2 & 3: table-by-table bulk copy, with row-by-row rescue per table.
        return RecoverManually();
    }

    /// <inheritdoc/>
    public void DeleteDatabase()
    {
        try
        {
            SqliteConnection.ClearAllPools();
            var path = DbPath;
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            throw;
        }
    }

    /// <inheritdoc/>
    public void InitializeFreshDatabase()
    {
        try
        {
            using var conn = OpenConnection(readOnly: false);
            CreateSchema(conn);
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            throw;
        }
    }

    // ── Recovery strategies ──────────────────────────────────────────────────

    /// <summary>
    /// Attempts SQLite's built-in <c>VACUUM INTO</c> to produce a clean copy of the database.
    /// Replaces the original on success.
    /// </summary>
    private RecoveryResult TryVacuumInto()
    {
        var tempPath = DbPath + ".recovering";
        try
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);

            using (var conn = OpenConnection(readOnly: false))
            {
                using var cmd = conn.CreateCommand();
                // Single-quote escaping for the path literal.
                cmd.CommandText = $"VACUUM INTO '{tempPath.Replace("'", "''")}'";
                cmd.ExecuteNonQuery();
            }

            SqliteConnection.ClearAllPools();
            File.Delete(DbPath);
            File.Move(tempPath, DbPath);
            var sessionsRecovered = TryCountSessions();
            return new RecoveryResult(Success: true, HadUnreadableRows: false,
                Strategy: "VACUUM INTO", SessionsRecovered: sessionsRecovered, SessionsLost: 0);
        }
        catch
        {
            SqliteConnection.ClearAllPools();
            TryDeleteFile(tempPath);
            return new RecoveryResult(Success: false, HadUnreadableRows: false,
                Strategy: "VACUUM INTO");
        }
    }

    /// <summary>
    /// Manually copies each table from the corrupt database into a fresh one,
    /// trying a full-table bulk read first and falling back to individual-row reads.
    /// Replaces the original on success.
    /// </summary>
    private RecoveryResult RecoverManually()
    {
        var tempPath = DbPath + ".recovering";
        bool hadUnreadableRows = false;

        try
        {
            // Open the source first — if it cannot be opened, fail before creating anything.
            using var srcConn = OpenConnectionByPath(DbPath, readOnly: true);

            // Read session count from the corrupt source for loss statistics (best-effort).
            int originalSessionCount = TryCountRows(srcConn, "sessions");

            if (File.Exists(tempPath)) File.Delete(tempPath);

            using (var destConn = OpenConnectionByPath(tempPath, readOnly: false))
            {
                CreateSchema(destConn);

                // Disable FK enforcement during the copy so rows can be inserted in any order
                // even when some referenced rows are missing from the corrupt source.
                using (var pragma = destConn.CreateCommand())
                {
                    pragma.CommandText = "PRAGMA foreign_keys = OFF";
                    pragma.ExecuteNonQuery();
                }

                // Copy in dependency order: referenced tables first.
                hadUnreadableRows |= CopyTableWithFallback(srcConn, destConn,
                    table:         "clipboard_formats",
                    selectSql:     "SELECT id, format_id, format_name FROM clipboard_formats ORDER BY id",
                    selectByIdSql: "SELECT id, format_id, format_name FROM clipboard_formats WHERE id = @id",
                    insertSql:     "INSERT OR IGNORE INTO clipboard_formats (id, format_id, format_name) VALUES (@id, @fid, @name)",
                    bindParams:    (r, cmd) =>
                    {
                        cmd.Parameters.AddWithValue("@id",   r.GetInt64(0));
                        cmd.Parameters.AddWithValue("@fid",  r.GetInt64(1));
                        cmd.Parameters.AddWithValue("@name", r.GetString(2));
                    });

                hadUnreadableRows |= CopyTableWithFallback(srcConn, destConn,
                    table:         "clipboard_contents",
                    selectSql:     "SELECT id, data, content_hash, original_size FROM clipboard_contents ORDER BY id",
                    selectByIdSql: "SELECT id, data, content_hash, original_size FROM clipboard_contents WHERE id = @id",
                    insertSql:     "INSERT OR IGNORE INTO clipboard_contents (id, data, content_hash, original_size) VALUES (@id, @data, @hash, @size)",
                    bindParams:    (r, cmd) =>
                    {
                        cmd.Parameters.AddWithValue("@id",   r.GetInt64(0));
                        cmd.Parameters.AddWithValue("@data", r.IsDBNull(1) ? DBNull.Value : (object)r.GetFieldValue<byte[]>(1));
                        cmd.Parameters.AddWithValue("@hash", r.GetString(2));
                        cmd.Parameters.AddWithValue("@size", r.GetInt64(3));
                    });

                hadUnreadableRows |= CopyTableWithFallback(srcConn, destConn,
                    table:         "sessions",
                    selectSql:     "SELECT id, timestamp, formats_text, total_size, pills_text FROM sessions ORDER BY id",
                    selectByIdSql: "SELECT id, timestamp, formats_text, total_size, pills_text FROM sessions WHERE id = @id",
                    insertSql:     "INSERT OR IGNORE INTO sessions (id, timestamp, formats_text, total_size, pills_text) VALUES (@id, @ts, @ft, @size, @pills)",
                    bindParams:    (r, cmd) =>
                    {
                        cmd.Parameters.AddWithValue("@id",    r.GetInt64(0));
                        cmd.Parameters.AddWithValue("@ts",    r.GetString(1));
                        cmd.Parameters.AddWithValue("@ft",    r.GetString(2));
                        cmd.Parameters.AddWithValue("@size",  r.GetInt64(3));
                        cmd.Parameters.AddWithValue("@pills", r.GetString(4));
                    });

                hadUnreadableRows |= CopyTableWithFallback(srcConn, destConn,
                    table:         "session_items",
                    selectSql:     "SELECT id, session_id, ordinal, clipboard_format_id, clipboard_content_id, handle_type, text_content FROM session_items ORDER BY id",
                    selectByIdSql: "SELECT id, session_id, ordinal, clipboard_format_id, clipboard_content_id, handle_type, text_content FROM session_items WHERE id = @id",
                    insertSql:     """
                        INSERT OR IGNORE INTO session_items
                            (id, session_id, ordinal, clipboard_format_id, clipboard_content_id, handle_type, text_content)
                        VALUES (@id, @sid, @ord, @fid, @cid, @ht, @tc)
                        """,
                    bindParams:    (r, cmd) =>
                    {
                        cmd.Parameters.AddWithValue("@id",  r.GetInt64(0));
                        cmd.Parameters.AddWithValue("@sid", r.GetInt64(1));
                        cmd.Parameters.AddWithValue("@ord", r.GetInt32(2));
                        cmd.Parameters.AddWithValue("@fid", r.GetInt64(3));
                        cmd.Parameters.AddWithValue("@cid", r.IsDBNull(4) ? DBNull.Value : (object)r.GetInt64(4));
                        cmd.Parameters.AddWithValue("@ht",  r.GetString(5));
                        cmd.Parameters.AddWithValue("@tc",  r.IsDBNull(6) ? DBNull.Value : (object)r.GetString(6));
                    });

                // Null out session_items rows that reference missing clipboard_contents,
                // then remove orphaned rows and blobs.
                NullOrphanedContentReferences(destConn);
                DeleteOrphanedSessionItems(destConn);
                DeleteOrphanedContent(destConn);
            }

            SqliteConnection.ClearAllPools();
            File.Delete(DbPath);
            File.Move(tempPath, DbPath);
            var sessionsRecovered = TryCountSessions();
            var sessionsLost      = Math.Max(0, originalSessionCount - sessionsRecovered);
            return new RecoveryResult(Success: true, HadUnreadableRows: hadUnreadableRows,
                Strategy: "manual copy", SessionsRecovered: sessionsRecovered, SessionsLost: sessionsLost);
        }
        catch
        {
            SqliteConnection.ClearAllPools();
            TryDeleteFile(tempPath);
            return new RecoveryResult(Success: false, HadUnreadableRows: false,
                Strategy: "manual copy");
        }
    }

    // ── Recovery helpers ─────────────────────────────────────────────────────

    /// <summary>
    /// Attempts a bulk copy of <paramref name="table"/> from <paramref name="src"/> to
    /// <paramref name="dest"/>. If the bulk copy fails at any point, falls back to reading
    /// one row at a time by primary key.
    /// Returns <see langword="true"/> if any rows were unreadable.
    /// </summary>
    private static bool CopyTableWithFallback(
        SqliteConnection src,
        SqliteConnection dest,
        string table,
        string selectSql,
        string selectByIdSql,
        string insertSql,
        Action<SqliteDataReader, SqliteCommand> bindParams)
    {
        if (TryCopyTableBulk(src, dest, selectSql, insertSql, bindParams))
            return false; // no data loss

        // Bulk failed; row-by-row may recover what the bulk scan could not.
        return CopyTableRowByRow(src, dest, table, selectByIdSql, insertSql, bindParams);
    }

    /// <summary>
    /// Bulk-copies all rows of a table using a single DataReader scan.
    /// Returns <see langword="true"/> if all rows were copied without error.
    /// </summary>
    private static bool TryCopyTableBulk(
        SqliteConnection src,
        SqliteConnection dest,
        string selectSql,
        string insertSql,
        Action<SqliteDataReader, SqliteCommand> bindParams)
    {
        using var tx = dest.BeginTransaction();
        try
        {
            using var srcCmd = src.CreateCommand();
            srcCmd.CommandText = selectSql;
            using var reader = srcCmd.ExecuteReader();

            while (reader.Read())
            {
                using var ins = dest.CreateCommand();
                ins.CommandText = insertSql;
                ins.Transaction = tx;
                bindParams(reader, ins);
                ins.ExecuteNonQuery();
            }

            tx.Commit();
            return true;
        }
        catch
        {
            try { tx.Rollback(); } catch { }
            return false;
        }
    }

    /// <summary>
    /// Reads each row of a table individually by primary key and inserts it into the
    /// destination. Rows that throw are skipped.
    /// Returns <see langword="true"/> if any rows were unreadable.
    /// </summary>
    private static bool CopyTableRowByRow(
        SqliteConnection src,
        SqliteConnection dest,
        string table,
        string selectByIdSql,
        string insertSql,
        Action<SqliteDataReader, SqliteCommand> bindParams)
    {
        var ids = GetTableIds(src, table);
        if (ids.Count == 0)
            return false;

        bool hadUnreadable = false;

        using var tx = dest.BeginTransaction();
        try
        {
            foreach (var id in ids)
            {
                try
                {
                    using var cmd = src.CreateCommand();
                    cmd.CommandText = selectByIdSql;
                    cmd.Parameters.AddWithValue("@id", id);
                    using var reader = cmd.ExecuteReader();

                    if (!reader.Read()) { hadUnreadable = true; continue; }

                    using var ins = dest.CreateCommand();
                    ins.CommandText = insertSql;
                    ins.Transaction = tx;
                    bindParams(reader, ins);
                    ins.ExecuteNonQuery();
                }
                catch { hadUnreadable = true; }
            }

            tx.Commit();
        }
        catch
        {
            try { tx.Rollback(); } catch { }
            hadUnreadable = true;
        }

        return hadUnreadable;
    }

    /// <summary>
    /// Returns all primary key values from <paramref name="table"/>,
    /// collecting as many as readable before any page error stops the scan.
    /// </summary>
    private static List<long> GetTableIds(SqliteConnection conn, string table)
    {
        var ids = new List<long>();
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT id FROM {table}";
            using var reader = cmd.ExecuteReader();
            while (true)
            {
                try
                {
                    if (!reader.Read()) break;
                    ids.Add(reader.GetInt64(0));
                }
                catch { break; }
            }
        }
        catch { }
        return ids;
    }

    /// <summary>
    /// Sets <c>clipboard_content_id</c> to NULL for any <c>session_items</c> row that
    /// references a content blob not present in the recovered database.
    /// </summary>
    private static void NullOrphanedContentReferences(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE session_items
            SET    clipboard_content_id = NULL
            WHERE  clipboard_content_id IS NOT NULL
              AND  clipboard_content_id NOT IN (SELECT id FROM clipboard_contents)
            """;
        cmd.ExecuteNonQuery();
    }

    /// <summary>
    /// Deletes <c>session_items</c> rows whose <c>session_id</c> or
    /// <c>clipboard_format_id</c> no longer exists.
    /// </summary>
    private static void DeleteOrphanedSessionItems(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            DELETE FROM session_items
            WHERE session_id          NOT IN (SELECT id FROM sessions)
               OR clipboard_format_id NOT IN (SELECT id FROM clipboard_formats)
            """;
        cmd.ExecuteNonQuery();
    }

    private static void TryDeleteFile(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

    /// <summary>
    /// Returns the number of rows in <paramref name="table"/> from an already-open connection,
    /// or 0 if the query fails.
    /// </summary>
    private static int TryCountRows(SqliteConnection conn, string table)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = $"SELECT COUNT(*) FROM {table}";
            return Convert.ToInt32(cmd.ExecuteScalar());
        }
        catch { return 0; }
    }

    /// <summary>
    /// Opens the current database and returns its session count, or 0 on any failure.
    /// </summary>
    private int TryCountSessions()
    {
        try
        {
            using var conn = OpenConnection(readOnly: true);
            return TryCountRows(conn, "sessions");
        }
        catch { return 0; }
        finally { SqliteConnection.ClearAllPools(); }
    }
}
