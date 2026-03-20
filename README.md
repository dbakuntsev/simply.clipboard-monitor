# C# .NET WPF Windows clipboard viewer/monitor with binary, text and image preview

`Simply.ClipboardMonitor` is a C# .NET WPF Windows desktop app for inspecting the current clipboard contents in real time.

It shows:
- The list of data formats currently available in the clipboard.
- Raw bytes as a hex dump (when the format is byte-addressable).
- Decoded text preview (for text-like formats), with encoding detection and a manual encoding selector.
- Image preview (for image-like formats).
- A scrollable history of past clipboard changes, with per-session format previews.

![Screenshot of Basic operation, HEX Preview](.github/screenshots/demo1.png)
![Screenshot of History Tracking, Image Preview](.github/screenshots/demo2.png)

## Download Binaries

Pre-built portable binaries are available from the [GitHub releases page](../../releases/latest).

If you wish to build from source, follow [Build and Run](#build-and-run) instructions.

## What This Project Is For

The app is designed as a clipboard debugging and inspection utility for software developers, QA engineers, and power users who need to understand exactly what an application places on the Windows clipboard.

## Intended Uses

- Debug clipboard integrations in your own apps.
- Validate which formats are published at copy time (e.g. text, HTML, custom formats, DIB).
- Inspect raw clipboard payloads via hex rows and offsets.
- Quickly sanity-check text encoding behavior (`CF_TEXT`, `CF_OEMTEXT`, `CF_UNICODETEXT`).
- Preview clipboard images and verify basic rendering.
- Browse and replay clipboard history — inspect any past clipboard state in full detail.
- Save a clipboard snapshot to a `.clipdb` file for later analysis or sharing.
- Export an individual clipboard format as text, image, or raw binary file.

## Potential Uses

- Reverse engineering clipboard behavior of third-party apps.
- Creating reproducible bug reports for clipboard-related issues.
- Comparing clipboard outputs across different apps for compatibility testing.
- Learning how Win32 clipboard formats map to real payloads.

## How It Works

At runtime, the main window registers as a clipboard listener using `AddClipboardFormatListener`.
Whenever Windows sends `WM_CLIPBOARDUPDATE` and clipboard change monitoring is active, the app:

1. Opens the clipboard (with retry).
2. Enumerates available format IDs with `EnumClipboardFormats`.
3. Resolves display names (well-known map + `GetClipboardFormatName`).
4. Reads the data size for each format — strategy depends on handle type:
   - **HGLOBAL formats** (`CF_DIB`, `CF_UNICODETEXT`, custom formats, etc.): `GlobalSize`.
   - **HBITMAP formats** (`CF_BITMAP`, `CF_DSPBITMAP`): `GetObject` → stride × height.
   - **HENHMETAFILE formats** (`CF_ENHMETAFILE`, `CF_DSPENHMETAFILE`): `GetEnhMetaFileBits` with a zero-length buffer.
5. Updates the format list UI.

When you select a format:

1. The app reads bytes via `GetClipboardData`, then extracts them using a strategy matched to the handle type:
   - **HGLOBAL**: `GlobalLock` / `GlobalUnlock` copy.
   - **HBITMAP**: `GetDIBits` converts the GDI bitmap to an uncompressed 32 bpp DIB byte block.
   - **HENHMETAFILE**: `GetEnhMetaFileBits` copies the raw EMF stream.
2. It renders a lazy hex table (16 bytes per row).
3. It attempts text decoding for text-like formats.
4. It attempts image decoding for image-like formats.

### Saving and Loading (.clipdb files)

Use **File → Save As…** to snapshot the current clipboard to a `.clipdb` file. Use **File → Save** to overwrite the most recently loaded or saved file. Use **File → Load…** to restore a snapshot into the clipboard.

Internally a `.clipdb` file is a SQLite database with two tables:

- `data_blobs` — stores each distinct binary payload once, keyed by its SHA-256 hash. Identical payloads across formats are stored only once.
- `clipboard_formats` — one row per clipboard format, recording the format ID, name, handle type (`hglobal` / `hbitmap` / `henhmetafile` / `none`), and a reference to the data blob.

On load, each format is restored using the handle-type-appropriate Win32 call:
- **HGLOBAL** formats: `GlobalAlloc` + `SetClipboardData`.
- **HBITMAP** formats: `CreateDIBitmap` (rebuilds a GDI bitmap from the stored 32 bpp DIB block) + `SetClipboardData`.
- **HENHMETAFILE** formats: `SetEnhMetaFileBits` + `SetClipboardData`.
- Custom format names (IDs ≥ 0xC000) are re-registered with `RegisterClipboardFormat` before writing, so format IDs remain correct even if they changed between Windows sessions.

## Clipboard Monitoring

Clipboard change monitoring is **enabled by default**. When active, the format list refreshes automatically every time the clipboard contents change.

### Turning monitoring ON and OFF

Use **Clipboard → Monitor Changes** in the menu bar to toggle monitoring on or off.

The menu item displays a checkmark and the status bar at the bottom shows "Monitoring..." while monitoring is active.

When monitoring is turned off, a full snapshot of all current clipboard format data is taken automatically so that format previews remain functional even when the clipboard contents have since changed.

### Persisted preference

The monitoring state is saved automatically when changed and restored on the next launch. The preference is stored in `%LOCALAPPDATA%\Simply.ClipboardMonitor\preferences.json`.

## Clipboard History

Use **Clipboard → Track History** to enable history tracking. A history panel appears below the format list, showing a timestamped record of every clipboard change captured while **Track History** was active.

### History list

Each row in the history list shows:
- **Date** - timestamp of the clipboard change.
- **Formats** — colored badges identifying the data categories present in that clipboard snapshot. The tooltip shows a list of every format name and its ID in the snapshot.

  | Pill | Category | Formats matched |
  |------|----------|-----------------|
  | **IMG** | Image | `CF_DIB`, `CF_DIBV5`, `CF_BITMAP`, PNG, JPEG, GIF, etc. |
  | **TXT** | Text | `CF_TEXT`, `CF_UNICODETEXT`, `CF_OEMTEXT`, and text-like format names |
  | **HTML** | HTML | `HTML Format` and HTML-like format names |
  | **RTF** | Rich Text | `Rich Text Format` and RTF-like format names |
  | **FILE** | File drop | `CF_HDROP` |
  | **OTHER** | Other | Shown only when none of the above apply |
- **Size** — sum of the original (uncompressed) byte lengths of all formats in the snapshot.

Clicking a row retrieves that session's format snapshots for preview, exactly as they were at the time of capture.

### History database

History is persisted to `%LOCALAPPDATA%\Simply.ClipboardMonitor\history.db` (an SQLite database). Blob data is stored compressed (ZStandard at maximum level) and content-addressed by SHA-256 hash, so identical payloads shared across different formats and different sessions are stored only once.

The database schema has three tables:
- `clipboard_formats` — the set of all format names and IDs ever seen.
- `clipboard_contents` — deduplicated compressed blobs, keyed by SHA-256 hash.
- `sessions` — one row per clipboard change, with timestamp, formats summary text, and total uncompressed size.
- `session_items` — joins sessions to their per-format blobs.

### History limits

The history database grows over time. Use **File → Settings** to configure:
- **Max entries** — oldest sessions are deleted when the count exceeds this value.
- **Max database size (MB)** — oldest sessions are deleted until the total stored blob size falls below this limit. At least one session is always retained.

Limits are enforced automatically each time a new session is written, as well as when the user changes the limits in the Settings dialog.

The Settings dialog also shows the current database file size and provides a **Clear History** button.

When history tracking is active, the status bar shows "Tracking history (X.X MB storage size)...".

## Preview Behavior

### Hex
- Available only for byte-addressable clipboard data.
- Rows show offset, hex bytes, and ASCII view.
- Uses lazy row materialization to keep large payload display responsive.

### Text
- Enabled for classic text IDs and text-like format names (`text`, `html`, `rtf`, `xml`, `json`, `csv`).
- Decoding strategy:
  - `CF_UNICODETEXT`: UTF-16 LE.
  - `CF_TEXT`: system ANSI/default code page.
  - `CF_OEMTEXT`: system OEM code page (e.g. CP437).
  - Others: UTF-8 BOM → UTF-16 BOM/heuristic → strict UTF-8 → system default.
- Status line shows character count, non-whitespace character count, and line count. Any of `\r`, `\n`, or `\r\n` counts as a line separator.
- An **Encoding** drop-down lists all encodings supported by Windows. The auto-detected encoding is pre-selected. Selecting a different encoding re-decodes the raw bytes on the spot; decoding failures are shown inline in red.
- The manually selected encoding is used when exporting as `.txt` (see below).

### Image
- Attempts image preview for:
  - **DIB formats** (`CF_DIB`, `CF_DIBV5`): the raw DIB block is prefixed with a `BITMAPFILEHEADER` and decoded by WPF's `BitmapDecoder`.
  - **HBITMAP formats** (`CF_BITMAP`, `CF_DSPBITMAP`): converted to a 32 bpp DIB via `GetDIBits`, then decoded the same way as DIB formats above.
  - **Encoded image formats** (names containing `png`, `jpeg`, `gif`, etc.): decoded directly from the byte stream.
- Includes fit-to-viewport baseline scale and user zoom multiplier (Ctrl + mouse wheel or zoom controls).
- Middle mouse button pans the image inside the scroll viewer.

### Exporting a format

Use **File → Export Selected Format…** (`Ctrl+E`) to save the currently selected clipboard format to a file. The command is disabled when no format is selected or the format has no captured data (Size = n/a).

A **Save As** dialog opens with the file name pre-set to `clipboard-{format_name}-{timestamp}`. The available file types depend on what previews are active for the selected format:

| Extension | Availability | Default when… |
|-----------|-------------|---------------|
| `.txt` | Text preview is available | Text preview is available |
| `.png` | Image preview is available | Image preview is available and format is not a JPEG |
| `.jpg` | Image preview is available | Format is a JPEG image (e.g. `image/jpeg`) |
| `.bin` | Always | No other format applies |

- **`.txt`** — decodes the raw bytes using the auto-detected encoding. If the Text tab is active and the encoding was manually changed, the manually selected encoding is used instead.
- **`.png`** — re-encodes the current image preview as a PNG file.
- **`.jpg`** — for natively JPEG clipboard formats writes the raw bytes unchanged; for all other image sources re-encodes at 80% quality.
- **`.bin`** — writes the raw clipboard bytes as-is.

## Known Limitations

- **`CF_PALETTE` (HPALETTE)**: palette objects cannot be read as a raw byte stream. The format appears in the list but has no hex, image, or save/restore support.
- **`CF_ENHMETAFILE` / `CF_DSPENHMETAFILE`** (HENHMETAFILE): the raw EMF byte stream is available in the hex viewer and is saved/restored correctly, but no image rendering is provided — WPF has no native EMF decoder and rendering via GDI is outside the current scope.
- **`CF_METAFILEPICT` / `CF_DSPMETAFILEPICT`** (HGLOBAL wrapping a `METAFILEPICT` struct): the hex viewer shows the raw struct bytes; the embedded `HMETAFILE` handle value inside is not dereferenced and the metafile data is not separately captured.
- There are no automated tests in this repository currently.

## Tech Stack

- C#
- WPF
- .NET 8 (`net8.0-windows`)
- Win32 APIs via P/Invoke (`user32.dll`, `kernel32.dll`, `gdi32.dll`)
- SQLite via `Microsoft.Data.Sqlite` (clipboard database persistence and history)
- ZStandard via `ZstdSharp.Port` (history blob compression)

## Project Structure

- `Simply.ClipboardMonitor.sln` — solution
- `Simply.ClipboardMonitor/Simply.ClipboardMonitor.csproj` — app project
- `Simply.ClipboardMonitor/App.xaml` / `App.xaml.cs` — WPF application entry point
- `Simply.ClipboardMonitor/Views/MainWindow.xaml` — main window UI layout
- `Simply.ClipboardMonitor/Views/MainWindow.xaml.cs` — main window logic (clipboard listener, parsing, previews, encoding detection, export, zoom, pan, sort, history, preferences)
- `Simply.ClipboardMonitor/Views/AboutDialog.xaml` / `AboutDialog.xaml.cs` — About dialog
- `Simply.ClipboardMonitor/Views/SettingsDialog.xaml` / `SettingsDialog.xaml.cs` — Settings dialog (history limits, database size display, clear history)
- `Simply.ClipboardMonitor/Common/ClipboardFormatItem.cs` — format list row model
- `Simply.ClipboardMonitor/Common/HexRowCollection.cs` — virtualised hex-dump row collection
- `Simply.ClipboardMonitor/Common/ClipboardDatabase.cs` — `.clipdb` save/load logic (SQLite)
- `Simply.ClipboardMonitor/Common/ClipboardHistory.cs` — history database logic (SQLite, ZStandard compression, SHA-256 deduplication, session trimming)
- `Simply.ClipboardMonitor/Common/NativeMethods.cs` — Win32 P/Invoke declarations
- `Simply.ClipboardMonitor/Common/ShellHelper.cs` — helper for opening URLs in the default browser
- `Simply.ClipboardMonitor/Models/UserPreferences.cs` — preferences and column preference model

## Build and Run

1. Ensure you have the .NET 8 SDK installed on your machine (`dotnet --list-sdks`).
2. Clone the repository.
3. Navigate to the root directory of the repository (where `Simply.ClipboardMonitor.sln` file is located).
4. For debug builds:
	* Execute `dotnet run`.
5. For release builds:
	* Execute `dotnet run --project Simply.ClipboardMonitor\Simply.ClipboardMonitor.csproj -c Release`.

## License

MIT License; see `LICENSE.txt`.
