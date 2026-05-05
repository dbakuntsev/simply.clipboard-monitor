using Microsoft.Data.Sqlite;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services.Impl.HistoryDb;
using System.IO;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Persists clipboard change history to a local SQLite database (history.db).
///
/// Blob storage, schema management, and database recovery are delegated to
/// focused helper classes in the <c>HistoryDb</c> subfolder.
/// </summary>
internal sealed class HistoryRepository : IHistoryRepository, IHistoryMaintenance
{
    private readonly HistoryDbPath           _dbPath;
    private readonly DatabaseRecoveryService _recovery;

    // Parameterless constructor used by the DI container.
    public HistoryRepository() : this(new HistoryDbPath()) { }

    private HistoryRepository(HistoryDbPath dbPath)
    {
        _dbPath   = dbPath;
        _recovery = new DatabaseRecoveryService(dbPath);
    }

    /// <summary>Creates an instance backed by a specific database file path, for use in tests.</summary>
    internal static HistoryRepository CreateForTesting(string dbPath) =>
        new(new HistoryDbPath(dbPath));

    // ── IHistoryRepository ───────────────────────────────────────────────────

    /// <inheritdoc/>
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
            using var conn = _dbPath.OpenConnection(readOnly: false);
            SchemaManager.CreateSchema(conn);

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
                            var hash = BlobStore.ComputeHash(snap.Data);
                            contentId = BlobStore.EnsureContent(conn, hash, snap.Data, snap.OriginalSize);
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

    /// <inheritdoc/>
    public List<SessionEntry> LoadSessions(string? filter = null, int maxCount = 2000)
    {
        if (!File.Exists(_dbPath.Value))
            return [];

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();

            if (string.IsNullOrWhiteSpace(filter))
            {
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

    /// <inheritdoc/>
    public List<FormatSnapshot> LoadSessionFormats(long sessionId)
    {
        if (!File.Exists(_dbPath.Value))
            return [];

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
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
                    data = BlobStore.Decompress(reader.GetFieldValue<byte[]>(4));

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

    /// <inheritdoc/>
    public string BuildFormatsText(IReadOnlyList<FormatSnapshot> snapshots)
    {
        var joined = string.Join(", ", snapshots.Select(s => s.FormatName));
        return joined.Length > 80 ? joined[..77] + "..." : joined;
    }

    /// <inheritdoc/>
    public int GetSessionCount()
    {
        if (!File.Exists(_dbPath.Value))
            return 0;
        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
            using var cmd  = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM sessions";
            return (int)(long)cmd.ExecuteScalar()!;
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex);
            return 0;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <inheritdoc/>
    public void DeleteSession(long sessionId)
    {
        if (!File.Exists(_dbPath.Value))
            return;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: false);
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

                    BlobStore.DeleteOrphanedContent(conn);
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
        if (!File.Exists(_dbPath.Value) || snapshots.Count == 0)
            return false;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
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
                var snap                 = snapshots[i];
                var (lastName, lastHash) = lastItems[i];

                if (!string.Equals(snap.FormatName, lastName, StringComparison.Ordinal))
                    return false;

                var currentHash = snap.Data is { Length: > 0 } ? BlobStore.ComputeHash(snap.Data) : null;
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

    // ── IHistoryMaintenance ──────────────────────────────────────────────────

    /// <inheritdoc/>
    public long GetDatabaseFileSize()
    {
        var path = _dbPath.Value;
        return File.Exists(path) ? new FileInfo(path).Length : 0L;
    }

    /// <inheritdoc/>
    public bool EnforceLimits(int maxEntries, long maxDatabaseBytes)
    {
        if (!File.Exists(_dbPath.Value))
            return false;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: false);
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

    /// <inheritdoc/>
    public void MigrateSchema()
    {
        if (!File.Exists(_dbPath.Value))
            return;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: false);
            SchemaManager.CreateSchema(conn);
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
    public void ClearHistory()
    {
        if (!File.Exists(_dbPath.Value))
            return;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: false);

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
    public DatabaseIntegrityStatus CheckIntegrity() => _recovery.CheckIntegrity();

    /// <inheritdoc/>
    public RecoveryResult TryRecover() => _recovery.TryRecover();

    /// <inheritdoc/>
    public void DeleteDatabase()
    {
        try
        {
            SqliteConnection.ClearAllPools();
            var path = _dbPath.Value;
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
            using var conn = _dbPath.OpenConnection(readOnly: false);
            SchemaManager.CreateSchema(conn);
        }
        catch (Exception ex) when (ex is not FileNotFoundException and not DirectoryNotFoundException)
        {
            ErrorLogger.Log(ex);
            throw;
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    // ── Write helpers ────────────────────────────────────────────────────────

    private static long InsertSession(SqliteConnection conn, DateTime timestamp, string formatsText, long totalSize, string pillsText)
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

    private static void InsertSessionItem(SqliteConnection conn, long sessionId, int ordinal,
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

    // ── Limit enforcement ────────────────────────────────────────────────────

    private bool TrimToLimits(SqliteConnection conn, int maxEntries, long maxDatabaseBytes)
    {
        var deleted = false;

        if (maxEntries > 0)
        {
            var count = GetSessionCount(conn);
            if (count > maxEntries)
            {
                using var tx = conn.BeginTransaction();
                DeleteOldestSessions(conn, (int)(count - maxEntries));
                BlobStore.DeleteOrphanedContent(conn);
                tx.Commit();
                deleted = true;
            }
        }

        if (maxDatabaseBytes > 0)
        {
            while (GetStoredBlobBytes(conn) > maxDatabaseBytes)
            {
                if (GetSessionCount(conn) <= 1) break;

                using var tx = conn.BeginTransaction();
                DeleteOldestSessions(conn, 1);
                BlobStore.DeleteOrphanedContent(conn);
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
}
