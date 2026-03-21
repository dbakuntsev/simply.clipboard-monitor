using Microsoft.Data.Sqlite;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Saves and loads clipboard snapshots to/from a passwordless SQLite database (.clipdb).
///
/// Identical data blocks are stored only once (content-addressed by SHA-256 hash),
/// so copying one large format many times does not inflate the file.
/// </summary>
internal sealed class ClipboardFileRepository : IClipboardFileRepository
{
    // ── Public API ──────────────────────────────────────────────────────────

    /// <summary>
    /// Creates (or overwrites) the file at <paramref name="path"/> and writes
    /// every format in <paramref name="formats"/> into it.
    /// </summary>
    public void Save(string path, IReadOnlyList<SavedClipboardFormat> formats)
    {
        if (File.Exists(path))
            File.Delete(path);

        try
        {
            using var conn = OpenConnection(path, readOnly: false);
            CreateSchema(conn);

            using var tx = conn.BeginTransaction();
            try
            {
                foreach (var fmt in formats)
                {
                    string? hash = null;
                    if (fmt.Data is { Length: > 0 })
                    {
                        hash = ComputeHash(fmt.Data);
                        EnsureBlob(conn, hash, fmt.Data);
                    }

                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = """
                    INSERT INTO clipboard_formats (ordinal, format_id, format_name, handle_type, data_hash)
                    VALUES (@ordinal, @formatId, @formatName, @handleType, @dataHash)
                    """;
                    cmd.Parameters.AddWithValue("@ordinal",    fmt.Ordinal);
                    cmd.Parameters.AddWithValue("@formatId",   (long)fmt.FormatId);
                    cmd.Parameters.AddWithValue("@formatName", fmt.FormatName);
                    cmd.Parameters.AddWithValue("@handleType", fmt.HandleType);
                    cmd.Parameters.AddWithValue("@dataHash",   (object?)hash ?? DBNull.Value);
                    cmd.ExecuteNonQuery();
                }

                tx.Commit();
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    /// <summary>
    /// Opens the file at <paramref name="path"/> and returns all stored clipboard formats.
    /// </summary>
    public List<SavedClipboardFormat> Load(string path)
    {
        try
        {
            using var conn = OpenConnection(path, readOnly: true);

            var result = new List<SavedClipboardFormat>();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = """
            SELECT cf.ordinal, cf.format_id, cf.format_name, cf.handle_type, db.data
            FROM   clipboard_formats cf
            LEFT JOIN data_blobs db ON cf.data_hash = db.hash
            ORDER  BY cf.ordinal
            """;

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var ordinal    = reader.GetInt32(0);
                var formatId   = (uint)reader.GetInt64(1);
                var formatName = reader.GetString(2);
                var handleType = reader.GetString(3);
                byte[]? data   = null;

                if (!reader.IsDBNull(4))
                    data = reader.GetFieldValue<byte[]>(4);

                result.Add(new SavedClipboardFormat(ordinal, formatId, formatName, handleType, data));
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

            CREATE TABLE IF NOT EXISTS data_blobs (
                hash  TEXT NOT NULL PRIMARY KEY,
                data  BLOB NOT NULL
            );

            CREATE TABLE IF NOT EXISTS clipboard_formats (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                ordinal     INTEGER NOT NULL,
                format_id   INTEGER NOT NULL,
                format_name TEXT    NOT NULL,
                handle_type TEXT    NOT NULL,
                data_hash   TEXT    REFERENCES data_blobs(hash)
            );
            """;
        cmd.ExecuteNonQuery();
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    private void EnsureBlob(SqliteConnection conn, string hash, byte[] data)
    {
        using var check = conn.CreateCommand();
        check.CommandText = "SELECT COUNT(1) FROM data_blobs WHERE hash = @hash";
        check.Parameters.AddWithValue("@hash", hash);
        if ((long)check.ExecuteScalar()! > 0)
            return;

        using var insert = conn.CreateCommand();
        insert.CommandText = "INSERT INTO data_blobs (hash, data) VALUES (@hash, @data)";
        insert.Parameters.AddWithValue("@hash", hash);
        insert.Parameters.AddWithValue("@data", data);
        insert.ExecuteNonQuery();
    }

    private string ComputeHash(byte[] data)
    {
        Span<byte> hash = stackalloc byte[32];
        SHA256.HashData(data, hash);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private SqliteConnection OpenConnection(string path, bool readOnly)
    {
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
