using Microsoft.Data.Sqlite;
using System.Security.Cryptography;
using ZstdSharp;

namespace Simply.ClipboardMonitor.Services.Impl.HistoryDb;

internal static class BlobStore
{
    private const int ZstdMaxLevel = 22;

    internal static string ComputeHash(byte[] data)
    {
        Span<byte> hash = stackalloc byte[32];
        SHA256.HashData(data, hash);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    internal static byte[] Compress(byte[] data)
    {
        using var compressor = new Compressor(ZstdMaxLevel);
        return compressor.Wrap(data).ToArray();
    }

    internal static byte[] Decompress(byte[] compressed)
    {
        using var decompressor = new Decompressor();
        return decompressor.Unwrap(compressed).ToArray();
    }

    /// <summary>
    /// Inserts a compressed + hashed content blob if it does not already exist.
    /// Returns the row ID of the existing or newly inserted row.
    /// </summary>
    internal static long EnsureContent(SqliteConnection conn, string hash, byte[] data, long originalSize)
    {
        using var check = conn.CreateCommand();
        check.CommandText = "SELECT id FROM clipboard_contents WHERE content_hash = @hash";
        check.Parameters.AddWithValue("@hash", hash);
        var existing = check.ExecuteScalar();
        if (existing != null)
            return (long)existing;

        var compressed = Compress(data);

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

    /// <summary>
    /// Removes content blobs that are no longer referenced by any session item.
    /// </summary>
    internal static void DeleteOrphanedContent(SqliteConnection conn)
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
}
