# C# .NET WPF Windows clipboard viewer/monitor with binary, text and image preview

`Simply.ClipboardMonitor` is a C# .NET WPF Windows desktop app for inspecting the current clipboard contents in real time.

It shows:
- The list of data formats currently available in the clipboard.
- Raw bytes as a hex dump (when the format is byte-addressable).
- Decoded text preview (for text-like formats), with encoding detection and a manual encoding selector.
- Locale preview (for `CF_LOCALE`): LCID hex value, BCP-47 language tag, and display name.
- HTML preview (rendered in an embedded WebView2 control).
- RTF preview (rendered by WPF's built-in rich text renderer, with compound GDI font-name normalization for correct bold/italic rendering).
- Image preview (for image-like formats).
- A scrollable history of past clipboard changes, with per-session format previews and drag-and-drop support for transferring history entries directly into other applications.
- The process that currently owns the clipboard, shown in the status bar with full path and command line in a tooltip.

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

### Pausing monitoring

When monitoring is ON, use **Clipboard → Pause** to temporarily suspend clipboard capture for a fixed duration without fully disabling the feature:

| Menu item | Duration added |
|-----------|---------------|
| Pause → 1 minute | 1 minute |
| Pause → 5 minutes | 5 minutes |
| Pause → 10 minutes | 10 minutes |
| Pause → 30 minutes | 30 minutes |
| Pause → 1 hour | 60 minutes |

Selecting a duration while a pause is already active extends the remaining pause by that amount (the submenu items are prefixed with **+** in that case).

Use **Clipboard → Resume** to cancel the pause early and resume monitoring immediately.

The status bar shows "Paused until HH:MM..." while a pause is active. Turning monitoring OFF also cancels any active pause.

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

Clicking a row retrieves that session's format snapshots for preview, exactly as they were at the time of capture. If a clipboard format was already selected in the format list and the new session contains a format with the same name, that format is automatically re-selected so the preview refreshes immediately without extra clicks. Likewise, the active content preview tab (Hex / Text / Image) is preserved across row changes, allowing rapid inspection of the same data view across multiple history entries.

The number of entries currently visible is shown in the **Clipboard History** group header (right-aligned). When a filter is active it reads *N items (M total)*, where *M* is the total unfiltered session count and *N* is the number of matched entries.

### Drag-and-drop from history

A history entry can be dragged from the history list and dropped directly into another application — no need to first load it back onto the clipboard.

The WPF `DataObject` placed on the drag session includes all recognised formats from the selected entry:

| Format | Source |
|--------|--------|
| Unicode / plain text | `CF_UNICODETEXT`, `CF_TEXT`, or `CF_OEMTEXT` |
| HTML | `HTML Format`, `text/html`, or other HTML-like format names |
| RTF | `Rich Text Format` or other RTF-like format names |
| File drop | `CF_HDROP` (DROPFILES) |
| DIB image | `CF_DIB` / `CF_DIBV5` raw bytes (for native apps such as Word or Paint) |
| Bitmap | `CF_BITMAP` / `CF_DSPBITMAP` — decoded to a frozen `BitmapSource` (for WPF/modern apps) |
| PNG | `PNG` or `image/png` custom formats |

Only formats that have captured data are included. If no supported format has data, the drag is suppressed.

### History entry actions

Right-clicking a row in the history list opens a context menu:

| Action | Description |
|--------|-------------|
| **Load into Clipboard** | Replaces the current clipboard with the full contents of the selected history entry. Available only when the entry differs from what is already on the clipboard; loading triggers a new history entry just like any other clipboard change. |
| **Save As…** | Opens a **Save As** dialog and saves the selected history entry to a `.clipdb` file without touching the live clipboard. |
| **Delete** | Permanently removes the selected entry. The format list reverts to the current live clipboard. |
| **Clear All** | Deletes all history entries (equivalent to **File → Settings → Clear History**). The format list reverts to the current live clipboard. |

### Startup and toggle-on capture

When **Track History** is enabled — either because it was on at last exit and is restored on startup, or because the user enables it manually — the current clipboard contents are automatically captured as a new history entry if they differ from the most recently recorded session. If the clipboard matches the last entry, that entry is selected instead. This ensures the history list always starts with the current clipboard state.

### Filtering

A filter box sits at the top of the history panel. Typing text in narrows the list to sessions whose timestamp, format pill labels, format names, or decoded text content contains the search term.

- The search is case-insensitive.
- The group header count updates to show *N items (M total)* while a filter is active, making it easy to see how many sessions matched.
- When a session is selected while a filter is active, format rows that match the filter term are highlighted with an orange left border.
  - Format rows are highlighted when the filter term appears in the **format name** or in the format's **pill label** (e.g. typing `img` highlights every image-compatible format, not only those whose registered name contains those letters).
- Clearing the filter (typing or clicking **×**) reloads the full list and automatically selects the latest session.
- New clipboard events captured while a filter is active re-run the filtered query; the format panel remains unchanged if the new session does not match the current filter.

### History database

History is persisted to `%LOCALAPPDATA%\Simply.ClipboardMonitor\history.db` (an SQLite database). Blob data is stored compressed (ZStandard at maximum level) and content-addressed by SHA-256 hash, so identical payloads shared across different formats and different sessions are stored only once.

The database schema has four tables:
- `clipboard_formats` — the set of all format names and IDs ever seen.
- `clipboard_contents` — deduplicated compressed blobs, keyed by SHA-256 hash.
- `sessions` — one row per clipboard change, with timestamp, formats summary text, total uncompressed size, and a space-separated list of pill labels (`pills_text`) used for fast pill-label filtering.
- `session_items` — joins sessions to their per-format blobs; includes a `text_content` column storing the decoded text for text-like formats, used for full-text filtering.

The schema is migrated automatically on first launch after an upgrade — new columns are added with `ALTER TABLE … ADD COLUMN` when they are missing, so existing databases are upgraded in place without data loss.

### History limits

The history database grows over time. Use **File → Settings** to configure:
- **Max entries** — oldest sessions are deleted when the count exceeds this value.
- **Max database size (MB)** — oldest sessions are deleted until the total stored blob size falls below this limit. At least one session is always retained.

Limits are enforced automatically each time a new session is written, as well as when the user changes the limits in the Settings dialog.

The Settings dialog also shows the current database file size and provides a **Clear History** button.

When history tracking is active, the status bar shows "Tracking history (X.X MB storage size)...".

### Database integrity

Each time history tracking is initialized — on startup when **Track History** is on, or when the user enables it — the app runs `PRAGMA integrity_check` on `history.db`. If corruption is detected, a dialog is shown with three choices:

| Choice | Effect |
|--------|--------|
| **Recover** | Attempts to salvage as much data as possible. Tries `VACUUM INTO` first (SQLite's built-in clean-copy mechanism); if that fails, falls back to a table-by-table bulk copy and then row-by-row rescue for any table whose bulk copy fails. On success the recovered file replaces the original. If some rows were unreadable, a warning is shown. If all strategies fail, a second dialog offers Delete or Disable. |
| **Delete and Start Fresh** | Deletes the corrupt file and creates a new empty database. |
| **Disable History Tracking** | Turns off history tracking. Any clipboard changes that arrived while the dialog was open are discarded. |

The integrity check is run again whenever **Track History** is toggled on after being off, unless the database was already verified during the current session.

## System Tray

The application can be kept running in the system tray instead of closing when the main window is dismissed.

### Enabling Minimize to System Tray

Open **File → Settings** and check **Minimize to System Tray**. When the setting is ON:

- The system tray icon is visible at all times while the application is running.
- Closing the main window hides it rather than exiting. The first time the window is hidden this way, a balloon notification appears to confirm the application is still running.
- Left-clicking the tray icon toggles the window between visible and hidden.
- Right-clicking the tray icon opens a context menu that mirrors the main menu bar:
  - **Show/Hide Window** — toggles window visibility.
  - **File** submenu — Load, Save, Save As, Export Selected Format, Settings. Items that open a dialog bring the main window to the foreground first.
  - **Clipboard** submenu — Clear, Refresh, Monitor Changes (checkable), Track History (checkable).
  - **Help** submenu — Submit Feedback, About.
  - **Pause** submenu — suspend monitoring for 1 minute, 5 minutes, 10 minutes, 30 minutes, or 1 hour. While a pause is active the current expiry time is shown as a disabled label at the top of the tray menu, and duration items are prefixed with **+** to extend the pause.
  - **Resume** — available only while a pause is active; cancels the pause immediately.
  - **Exit** — closes the application immediately without hiding to the tray.
- **File → Exit** always exits the application, regardless of the setting.

### Global Hotkey

When **Minimize to System Tray** is ON, a configurable global hotkey can be used to show, hide, or bring the main window to the foreground from anywhere on the desktop — even when another application has focus.

Open **File → Settings** to configure the hotkey:

- Check **Enable global hotkey** to activate it.
- Click the key-capture field and press the desired key combination to record it. The combination must include at least one of Alt, Ctrl, or Win; bare keys and Shift-only combinations are rejected.
- Press **Escape** while the field has focus to cancel the capture and keep the existing binding.
- The default binding is **Alt+Win+V**.

If the chosen combination is already registered by another application, a conflict warning is shown in the Settings dialog. The hotkey is automatically unregistered while the Settings dialog is open to avoid interfering with key capture.

Pressing the hotkey when the window is hidden shows and activates it; pressing it when the window is already in the foreground hides it to the tray.

## Auto-Start

Open **File → Settings** to configure how the application starts.

### Start at login

When **Start at login** is ON, the application registers itself in the Windows current-user auto-start registry key (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) so it launches automatically after login. Disabling the setting removes the entry immediately.

The status bar shows an orange **AUTO-START** pill on the right edge whenever this setting is ON, providing a persistent visual reminder.

### Start minimized

When **Start minimized** is ON, the application window does not appear on screen at startup:

- If **Minimize to System Tray** is also ON — the window is hidden and the tray icon is shown (same as manually closing the window with the tray option enabled).
- If **Minimize to System Tray** is OFF — the window starts minimized in the taskbar.

### Persisted preference

The **Minimize to System Tray** setting is saved in `%LOCALAPPDATA%\Simply.ClipboardMonitor\preferences.json` along with a flag that tracks whether the first-minimize balloon notification has been shown (so it appears only once ever). **Start at login** state is read directly from the registry on each launch so the UI always reflects the actual system state.

## Single Instance

Only one instance of the application can run per Windows login session. If a second instance is launched while the application is already running:

- The second instance signals the first instance to show and activate its main window.
- The second instance exits immediately.

The enforcement is session-scoped (using a named mutex without the `Global\` prefix), so separate interactive logon sessions on the same machine each get their own independent instance.

## Clipboard Owner

Whenever the clipboard contents change, the status bar shows which process last wrote to the clipboard, on the right side of the status bar:

```
Owner: chrome.exe (12345)
```

Hovering the text reveals a tooltip with three fields:

- **PID** — the numeric process identifier.
- **Path** — the full executable image path, resolved via `QueryFullProcessImageName`.
- **Command line** — each argument listed on its own indented line, resolved by reading the process's PEB through `NtQueryInformationProcess` (class 0 for native 64-bit processes; class 26 + WOW64 PEB walk for 32-bit processes running under WOW64). Long arguments wrap at a fixed tooltip width.

## Error Logs

When the application encounters an unhandled exception, it shows a crash dialog before closing. The dialog displays the full path to the error log file as a clickable link that opens the file directly.

Error logs are written to `%LOCALAPPDATA%\Simply.ClipboardMonitor\` and named `error_YYYY-MM-DD.txt`. If the file with today's date does not exist, it means that no loggable errors occurred today. The three most recent log files are kept; older files are deleted automatically.

In addition to unhandled exceptions, the following situations are logged without interrupting the application:

- Errors reading or writing the preferences file — unless the file is simply absent.
- Errors reading or writing the history database — unless the database file is simply absent.
- Errors reading or writing `.clipdb` snapshot files — unless the file is simply absent.
- Unobserved background task exceptions.
- Detection of a corrupted history database.
- The choice made by the user in the corruption dialog (Recover / Delete and Start Fresh / Disable History Tracking).
- Recovery outcome: strategy used (VACUUM INTO or manual copy), number of sessions recovered, and estimated number of sessions lost.

The data directory is also accessible from **File → Settings**, which has a link at the bottom of the dialog that opens Windows Explorer at `%LOCALAPPDATA%\Simply.ClipboardMonitor\`.

## Preview Behavior

Each preview is displayed in its own tab. When a format is selected, all tabs are updated simultaneously. If the currently active tab becomes unavailable for the selected format, the tab with the lowest priority that can display the format is selected automatically.

### Hex
- Available for all byte-addressable clipboard data (always enabled).
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
- A **Word wrap** checkbox toggles line wrapping in the text view. The setting is persisted in `preferences.json` and restored on the next launch.

### Locale
- Enabled only for `CF_LOCALE`.
- Reads the LCID stored as a 32-bit little-endian integer.
- Displays three lines: the raw LCID as a hex value, the BCP-47 language tag (e.g. `en-US`), and the culture display name (e.g. `English (United States)`).
- The tag and display name rows are hidden when the LCID is not recognized by the .NET runtime.

### HTML
- Enabled for HTML-like format names (`HTML Format`, `text/html`, etc.).
- Rendered inside an embedded **WebView2** control (requires the Microsoft Edge WebView2 Evergreen Runtime to be installed).
- `HTML Format` (CF_HTML) payloads are parsed using the `StartHTML`/`EndHTML` byte offsets in the ASCII header; other HTML formats are decoded with the standard text pipeline.
- Scripts, context menus, dev tools, and user-initiated navigation are all disabled — the preview is strictly read-only.
- If the WebView2 Runtime is not installed, a message is shown in place of the rendered page.

### RTF
- Enabled for RTF-like format names (`Rich Text Format`, `richtext`, etc.).
- Rendered using WPF's built-in `RichTextBox` / `FlowDocument` RTF importer.
- Compound GDI face names (e.g. `Aptos SemiBold`, `Segoe UI Bold Italic`) are automatically split into a base family name plus explicit `FontWeight` and `FontStyle` values after import, so bold and italic styles render correctly even when the WPF font resolver cannot find the compound name as a standalone family.
- Status line shows character count and raw byte size.

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
- There are no image-rendering tests for `CF_ENHMETAFILE` / `CF_DSPENHMETAFILE` and no UI-level tests.

## Tech Stack

- C#
- WPF
- .NET 8 (`net8.0-windows`)
- Win32 APIs via P/Invoke (`user32.dll`, `kernel32.dll`, `gdi32.dll`, `shell32.dll`, `ntdll.dll`)
- SQLite via `Microsoft.Data.Sqlite` (clipboard database persistence and history)
- ZStandard via `ZstdSharp.Port` (history blob compression)
- `Microsoft.Extensions.DependencyInjection` (constructor injection throughout)
- `Microsoft.Web.WebView2` (HTML preview rendered in an embedded Chromium-based control)
- `System.Drawing.Common` (system tray icon loading; system warning icon in the corruption dialog)

## Project Structure

- `Simply.ClipboardMonitor.sln` — solution
- `Simply.ClipboardMonitor/Simply.ClipboardMonitor.csproj` — app project
- `Simply.ClipboardMonitor.Tests/Simply.ClipboardMonitor.Tests.csproj` — xUnit test project
- `Simply.ClipboardMonitor/App.xaml` / `App.xaml.cs` — WPF application entry point; builds the DI container and registers all services, strategies, exporters, and preview tab controls
- `Simply.ClipboardMonitor/Views/MainWindow.xaml` / `MainWindow.xaml.cs` — main window (clipboard listener, format list, history, export, sort, preferences); preview tabs are populated dynamically from `IEnumerable<IPreviewTab>` injected at construction
- `Simply.ClipboardMonitor/Views/AboutDialog.xaml` / `AboutDialog.xaml.cs` — About dialog
- `Simply.ClipboardMonitor/Views/SettingsDialog.xaml` / `SettingsDialog.xaml.cs` — Settings dialog (history limits, database size display, clear history, minimize-to-tray, start-at-login, start-minimized toggles, data directory link)
- `Simply.ClipboardMonitor/Views/CrashDialog.xaml` / `CrashDialog.xaml.cs` — Crash dialog shown on unhandled exceptions; displays a clickable link to the error log file
- `Simply.ClipboardMonitor/Views/DatabaseCorruptionDialog.xaml` / `DatabaseCorruptionDialog.xaml.cs` — Modal dialog shown when `history.db` fails an integrity check; offers Recover / Delete and Start Fresh / Disable History Tracking choices (also used as a two-option dialog after a failed recovery attempt)
- `Simply.ClipboardMonitor/Views/DatabaseRecoveringWindow.xaml` / `DatabaseRecoveringWindow.xaml.cs` — Non-closeable "Recovering…" progress window shown while database recovery runs in the background

### Models
Data-transfer records and plain model classes with no service dependencies.

- `Models/ClipboardOwnerInfo.cs` — display text and optional tooltip text for the clipboard owner status bar item
- `Models/ClipboardFormatItem.cs` — format list row model (ordinal, ID, name, content size, filter-highlight flag)
- `Models/EncodingItem.cs` — encoding display name + `Encoding` pair for the encoding combo box
- `Models/FormatColumnPreference.cs` — saved column width preference
- `Models/FormatExportContext.cs` — context record passed to `IFormatExporter` implementations
- `Models/FormatPill.cs` — colored badge record (label + brush) for the history Formats column
- `Models/FormatSnapshot.cs` — point-in-time capture of one clipboard format (handle type, raw bytes, original size)
- `Models/SavedClipboardFormat.cs` — format row stored in / loaded from a `.clipdb` file
- `Models/SessionEntry.cs` — one row from the history sessions table
- `Models/TextDecodeResult.cs` — result of a single text-decode attempt (text, encoding, success/failure)
- `Models/UserPreferences.cs` — top-level user preferences (sort property/direction, monitor/history settings, limits, minimize-to-tray, start-at-login, start-minimized toggles, balloon-shown flag, text word-wrap state, global hotkey enabled flag and binding string)

### Services
Public domain service interfaces consumed by the main window and DI wiring.

- `Services/IClipboardFileRepository.cs` — save/load clipboard snapshots to `.clipdb` files
- `Services/IClipboardOwnerService.cs` — resolve the current clipboard owner HWND to a process name, full path, and command line
- `Services/IClipboardReader.cs` — enumerate and read current clipboard formats; capture full snapshots
- `Services/IClipboardWriter.cs` — restore saved formats back onto the clipboard
- `Services/IFormatClassifier.cs` — classify formats into colored pills and tooltip text for the history list
- `Services/IFormatExporter.cs` — export a clipboard format to a file (one implementation per output type)
- `Services/IHistoryMaintenance.cs` — database maintenance operations (schema migration, enforce size limits, clear history, integrity check, recovery, delete, create fresh database)
- `Services/IHistoryRepository.cs` — clipboard history persistence (add session, load sessions with optional filter, load session formats, delete a single session, total session count, duplicate detection)
- `Services/DatabaseIntegrityStatus.cs` — `Absent` / `Healthy` / `Corrupted` enum returned by `IHistoryMaintenance.CheckIntegrity()`
- `Services/RecoveryResult.cs` — result record for `IHistoryMaintenance.TryRecover()`: success flag, data-loss flag, strategy name, sessions recovered, sessions lost
- `Services/IImagePreviewService.cs` — create WPF `BitmapSource` previews from raw clipboard bytes
- `Services/IPreferencesService.cs` — load and save user preferences
- `Services/IPreviewTab.cs` — common interface for all preview tab controls; exposes `TabItem`, `Priority`, `Update(formatId, name, bytes)`, and `Reset()`
- `Services/ITextDecodingService.cs` — decode raw bytes as text with auto-detection or manual encoding override

### Views/Previews
Each preview tab is a self-contained `UserControl` that implements `IPreviewTab`. The main window adds their `TabItem` instances to `ContentTabControl` at startup; no main-window changes are needed to add a new tab.

- `Views/Previews/HexPreviewControl.xaml` / `.xaml.cs` — hex dump tab (always enabled; Priority 0)
- `Views/Previews/TextPreviewControl.xaml` / `.xaml.cs` — text decoding tab with encoding selector (Priority 1); exposes `AutoDetectedEncoding` and `ManuallyChangedEncoding` for export
- `Views/Previews/LocalePreviewControl.xaml` / `.xaml.cs` — `CF_LOCALE` decoder tab showing LCID, BCP-47 tag, and display name (Priority 1)
- `Views/Previews/HtmlPreviewControl.xaml` / `.xaml.cs` — HTML preview tab using an embedded WebView2 control (Priority 2)
- `Views/Previews/RtfPreviewControl.xaml` / `.xaml.cs` — RTF preview tab using WPF's `RichTextBox`, including compound font-name normalization (Priority 2)
- `Views/Previews/ImagePreviewControl.xaml` / `.xaml.cs` — image preview tab with fit-scale, zoom slider, and pan (Priority 1); exposes `PreviewImageSource` for export

### Services/Impl
Concrete service implementations. All classes are `internal sealed`.

- `Services/Impl/ClipboardOwnerService.cs` — resolves `GetClipboardOwner` HWND → PID → process name via `QueryFullProcessImageName`; reads command line by walking the process PEB via `NtQueryInformationProcess` (native 64-bit path) or the WOW64 PEB (32-bit path); applies a two-tier `OpenProcess` fallback (`PROCESS_QUERY_INFORMATION | PROCESS_VM_READ` → `PROCESS_QUERY_LIMITED_INFORMATION`)
- `Services/Impl/IHandleReadStrategy.cs` — internal strategy interface for reading a specific clipboard handle type
- `Services/Impl/IHandleWriteStrategy.cs` — internal strategy interface for restoring a specific clipboard handle type
- `Services/Impl/ClipboardFileRepository.cs` — `.clipdb` save/load (SQLite, SHA-256 content deduplication)
- `Services/Impl/ClipboardReaderService.cs` — Win32 clipboard reading; dispatches per-handle-type work to injected `IHandleReadStrategy` implementations
- `Services/Impl/ClipboardWriterService.cs` — Win32 clipboard writing; dispatches per-handle-type work to injected `IHandleWriteStrategy` implementations
- `Services/Impl/FormatClassifierService.cs` — produces colored pills and tooltip text for the history list
- `Services/Impl/HistoryRepository.cs` — history database (SQLite, ZStandard compression, SHA-256 deduplication, session trimming, single-session deletion with orphan cleanup, schema migration, integrity check, three-strategy recovery, fresh-database initialization, duplicate detection); implements both `IHistoryRepository` and `IHistoryMaintenance`
- `Services/Impl/ImagePreviewService.cs` — decodes DIB, HBITMAP-derived, and encoded image formats into WPF `BitmapSource` objects
- `Services/Impl/PreferencesService.cs` — JSON preferences persisted to `%LOCALAPPDATA%\Simply.ClipboardMonitor\preferences.json`
- `Services/Impl/TextDecodingService.cs` — text decoding with format-aware priority chain and UTF-16 heuristics

### Services/Impl/Strategies
One class per clipboard handle type or export file format. All classes are `internal sealed`.

- `Strategies/NoneHandleReadStrategy.cs` — "none" handle type (e.g. `CF_PALETTE`): returns a failure message; no data read
- `Strategies/HGlobalHandleReadStrategy.cs` — HGLOBAL: `GlobalLock` / `GlobalUnlock` byte copy
- `Strategies/HBitmapHandleReadStrategy.cs` — HBITMAP: converts to 32 bpp DIB via `GetDIBits`
- `Strategies/HEnhMetaFileHandleReadStrategy.cs` — HENHMETAFILE: raw EMF bytes via `GetEnhMetaFileBits`
- `Strategies/HGlobalHandleWriteStrategy.cs` — HGLOBAL restore: `GlobalAlloc` + `SetClipboardData`
- `Strategies/HBitmapHandleWriteStrategy.cs` — HBITMAP restore: `CreateDIBitmap` + `SetClipboardData`
- `Strategies/HEnhMetaFileHandleWriteStrategy.cs` — HENHMETAFILE restore: `SetEnhMetaFileBits` + `SetClipboardData`
- `Strategies/TextFormatExporter.cs` — exports as `.txt` using the auto-detected or manually selected encoding
- `Strategies/PngFormatExporter.cs` — exports image preview as `.png`
- `Strategies/JpegFormatExporter.cs` — exports as `.jpg` (raw bytes for native JPEG formats; re-encoded at quality 80 otherwise)
- `Strategies/BinaryFormatExporter.cs` — exports raw bytes as `.bin` (always-available fallback)

### Common
Internal utility types with no domain logic.

- `Common/AutoStartHelper.cs` — reads and writes the Windows current-user auto-start registry key
- `Common/HotkeyBinding.cs` — immutable value type representing a global hotkey (modifier flags + virtual key code); includes `ToString` / `TryParse` for preferences serialisation and `FormatModifiers` for the live capture display
- `Common/ErrorLogger.cs` — thread-safe rolling logger; `Log(Exception)` for errors and `LogInfo(string)` for informational events; writes to dated `.txt` files under `%LOCALAPPDATA%\Simply.ClipboardMonitor\`, retaining the three most recent files
- `Common/ClipboardFormatConstants.cs` — Windows clipboard format ID constants (`CF_TEXT`, `CF_BITMAP`, etc.), handle-type classification sets, and shared image-format detection
- `Common/DisplayHelper.cs` — shared display formatting utilities (human-readable byte-size strings)
- `Common/HexRow.cs` — single hex-dump display row (offset, hex bytes, ASCII)
- `Common/HexRowCollection.cs` — lazy-loaded, cached `IReadOnlyList<HexRow>` over a raw byte array
- `Common/NativeMethods.cs` — Win32 P/Invoke declarations (`user32.dll`, `kernel32.dll`, `gdi32.dll`, `shell32.dll`, `ntdll.dll`)
- `Common/ShellHelper.cs` — opens a URL in the default browser via `ShellExecute`
- `Common/Win32Structs.cs` — Win32 structs used by clipboard read/write and process inspection (`BITMAP`, `BITMAPINFOHEADER`, `PROCESS_BASIC_INFORMATION`)

## Tests

The solution includes an xUnit test project (`Simply.ClipboardMonitor.Tests`) targeting `net8.0-windows`. Run tests with:

```
dotnet test
```

The test suite covers:

| File | What is tested |
|------|---------------|
| `TextDecodingServiceTests` | `IsTextCompatible`, `Decode` (all format ID branches and all auto-detection paths — UTF-8 BOM, UTF-16 LE/BE BOM, heuristic, fallback), `DecodeWith`, `GetDecodedTextStats` (including `\r`, `\n`, `\r\n` variants) |
| `FormatClassifierServiceTests` | `GetFormatPillLabel` for every category and the null/OTHER fallback; `ComputePills` including the "OTHER is suppressed when any known category matches" rule; `ComputeTooltip` singular/plural/empty |
| `DisplayHelperTests` | Zero, byte, KB, MB, GB boundary values; correct decimal rounding at 1.5× scale in all three ranges |
| `HexRowCollectionTests` | Row count (empty/exact/partial), offset formatting, hex pair content and padding, ASCII printable/non-printable mapping, out-of-range indexer, row caching, enumeration |
| `HistoryRepositoryTests` | `GetSessionCount`, `AddSession` with and without trim, `LoadSessions` (order, filter match/no-match), `LoadSessionFormats` with byte round-trip, `DeleteSession` (count, format removal, other sessions unaffected, shared blob preservation), `ClearHistory`, `IsDuplicateOfLastSession`, `EnforceLimits` (verifies the *oldest* sessions are removed), `BuildFormatsText` (name and truncation) |

`HistoryRepositoryTests` are integration tests that write to a temporary SQLite file created per test class instance; each directory is deleted in `Dispose`.

`FormatClassifierServiceTests` run on an STA thread (`[StaFact]` from `Xunit.StaFact`) because `FormatClassifierService` creates frozen `SolidColorBrush` objects in its static initialiser.

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
