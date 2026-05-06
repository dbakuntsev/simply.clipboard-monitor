using Simply.ClipboardMonitor.Common;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Result returned by <see cref="IDialogService.ShowSettings"/>.
/// <para>
/// <see cref="Saved"/> is <see langword="false"/> when the user cancelled;
/// the settings fields are meaningless in that case but <see cref="HistoryWasCleared"/>
/// is always valid (the user can clear history then cancel the dialog).
/// </para>
/// </summary>
public sealed record SettingsDialogResult(
    bool          Saved,
    bool          HistoryWasCleared,
    int           MaxEntries           = 0,
    int           MaxSizeMb            = 0,
    bool          MinimizeToSystemTray = false,
    bool          StartAtLogin         = false,
    bool          StartMinimized       = false,
    bool          HotkeyEnabled        = false,
    HotkeyBinding HotkeyBinding        = default);
