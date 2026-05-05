using Microsoft.Data.Sqlite;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Services;
using System.IO;

namespace Simply.ClipboardMonitor.Services.Impl.HistoryDb;

internal sealed class DatabaseRecoveryService
{
    private readonly HistoryDbPath _dbPath;

    internal DatabaseRecoveryService(HistoryDbPath dbPath) => _dbPath = dbPath;

    // ── Public API ───────────────────────────────────────────────────────────

    internal DatabaseIntegrityStatus CheckIntegrity()
    {
        if (!File.Exists(_dbPath.Value))
            return DatabaseIntegrityStatus.Absent;

        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
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

    internal RecoveryResult TryRecover()
    {
        // Strategy 1: VACUUM INTO a temp file — fastest when SQLite can read most pages.
        var vacuumResult = TryVacuumInto();
        if (vacuumResult.Success)
            return vacuumResult;

        // Strategies 2 & 3: table-by-table bulk copy, with row-by-row rescue per table.
        return RecoverManually();
    }

    // ── Recovery strategies ──────────────────────────────────────────────────

    private RecoveryResult TryVacuumInto()
    {
        var tempPath = _dbPath.Value + ".recovering";
        try
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);

            using (var conn = _dbPath.OpenConnection(readOnly: false))
            {
                using var cmd = conn.CreateCommand();
                // Single-quote escaping for the path literal.
                cmd.CommandText = $"VACUUM INTO '{tempPath.Replace("'", "''")}'";
                cmd.ExecuteNonQuery();
            }

            SqliteConnection.ClearAllPools();
            File.Delete(_dbPath.Value);
            File.Move(tempPath, _dbPath.Value);
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

    private RecoveryResult RecoverManually()
    {
        var tempPath = _dbPath.Value + ".recovering";
        bool hadUnreadableRows = false;

        try
        {
            // Open the source first — if it cannot be opened, fail before creating anything.
            using var srcConn = HistoryDbPath.OpenConnectionByPath(_dbPath.Value, readOnly: true);

            // Read session count from the corrupt source for loss statistics (best-effort).
            int originalSessionCount = TryCountRows(srcConn, "sessions");

            if (File.Exists(tempPath)) File.Delete(tempPath);

            using (var destConn = HistoryDbPath.OpenConnectionByPath(tempPath, readOnly: false))
            {
                SchemaManager.CreateSchema(destConn);

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
                BlobStore.DeleteOrphanedContent(destConn);
            }

            SqliteConnection.ClearAllPools();
            File.Delete(_dbPath.Value);
            File.Move(tempPath, _dbPath.Value);
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

    // ── Table copy helpers ───────────────────────────────────────────────────

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
            return false;

        return CopyTableRowByRow(src, dest, table, selectByIdSql, insertSql, bindParams);
    }

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

    // ── Recovery cleanup helpers ─────────────────────────────────────────────

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

    // ── Misc helpers ─────────────────────────────────────────────────────────

    private static void TryDeleteFile(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }

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

    private int TryCountSessions()
    {
        try
        {
            using var conn = _dbPath.OpenConnection(readOnly: true);
            return TryCountRows(conn, "sessions");
        }
        catch { return 0; }
        finally { SqliteConnection.ClearAllPools(); }
    }
}
