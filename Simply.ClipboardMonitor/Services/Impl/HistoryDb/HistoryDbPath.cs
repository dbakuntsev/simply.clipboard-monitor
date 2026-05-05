using Microsoft.Data.Sqlite;
using System.IO;

namespace Simply.ClipboardMonitor.Services.Impl.HistoryDb;

internal sealed class HistoryDbPath
{
    private readonly string? _override;

    internal HistoryDbPath(string? pathOverride = null) => _override = pathOverride;

    internal string Value => _override ??
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Simply.ClipboardMonitor",
            "history.db");

    internal SqliteConnection OpenConnection(bool readOnly) =>
        OpenConnectionByPath(Value, readOnly);

    internal static SqliteConnection OpenConnectionByPath(string path, bool readOnly)
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
}
