using System.Windows;

namespace Simply.ClipboardMonitor.Services.Impl;

internal sealed class DialogService : IDialogService
{
    private readonly IHistoryMaintenance _historyMaintenance;

    public DialogService(IHistoryMaintenance historyMaintenance)
        => _historyMaintenance = historyMaintenance;

    private static Window Owner => Application.Current.MainWindow;

    public void ShowAbout()
        => new AboutDialog { Owner = Owner }.ShowDialog();

    public SettingsDialogResult ShowSettings(SettingsDialogInput input)
    {
        var dlg = new SettingsDialog(
            input.MaxEntries, input.MaxSizeMb, _historyMaintenance,
            input.MinimizeToSystemTray, input.StartAtLogin, input.StartMinimized,
            input.HotkeyEnabled, input.HotkeyBinding, input.HotkeyConflict)
        { Owner = Owner };

        dlg.ShowDialog();

        return dlg.DialogResult == true
            ? new SettingsDialogResult(
                Saved:                true,
                HistoryWasCleared:    dlg.HistoryWasCleared,
                MaxEntries:           dlg.MaxEntries,
                MaxSizeMb:            dlg.MaxSizeMb,
                MinimizeToSystemTray: dlg.MinimizeToSystemTray,
                StartAtLogin:         dlg.StartAtLogin,
                StartMinimized:       dlg.StartMinimized,
                HotkeyEnabled:        dlg.HotkeyEnabled,
                HotkeyBinding:        dlg.GlobalHotkeyBinding)
            : new SettingsDialogResult(Saved: false, HistoryWasCleared: dlg.HistoryWasCleared);
    }

    public CorruptionDialogResult ShowDatabaseCorruption(bool showRecoverOption)
    {
        var dlg = new DatabaseCorruptionDialog(showRecoverOption) { Owner = Owner };
        dlg.ShowDialog();
        return dlg.Result;
    }

    public Action ShowDatabaseRecovering()
    {
        var window = new DatabaseRecoveringWindow { Owner = Owner };
        window.Show();
        return window.CloseWindow;
    }
}
