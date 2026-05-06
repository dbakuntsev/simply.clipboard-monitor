using Simply.ClipboardMonitor.Common;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Current settings values passed into <see cref="IDialogService.ShowSettings"/>.</summary>
public sealed record SettingsDialogInput(
    int          MaxEntries,
    int          MaxSizeMb,
    bool         MinimizeToSystemTray,
    bool         StartAtLogin,
    bool         StartMinimized,
    bool         HotkeyEnabled,
    HotkeyBinding HotkeyBinding,
    bool         HotkeyConflict);
