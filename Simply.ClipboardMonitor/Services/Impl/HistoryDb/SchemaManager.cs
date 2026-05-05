using Microsoft.Data.Sqlite;

namespace Simply.ClipboardMonitor.Services.Impl.HistoryDb;

internal static class SchemaManager
{
    internal static void CreateSchema(SqliteConnection conn)
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
        AddColumnIfMissing(conn, "sessions",      "pills_text",   "TEXT NOT NULL DEFAULT ''");
        AddColumnIfMissing(conn, "session_items", "text_content", "TEXT");
    }

    private static void AddColumnIfMissing(SqliteConnection conn, string table, string column, string definition)
    {
        using var info = conn.CreateCommand();
        info.CommandText = $"PRAGMA table_info({table})";
        using var reader = info.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase))
                return;
        }
        using var alter = conn.CreateCommand();
        alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {definition}";
        alter.ExecuteNonQuery();
    }
}
