namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Abstracts modal dialog and overlay-window creation so ViewModels can trigger
/// dialogs without taking a direct dependency on WPF Window types.
/// </summary>
public interface IDialogService
{
    /// <summary>Shows the About dialog modally.</summary>
    void ShowAbout();

    /// <summary>
    /// Shows the Settings dialog modally.
    /// <see cref="SettingsDialogResult.HistoryWasCleared"/> is valid regardless of whether
    /// the user saved; check <see cref="SettingsDialogResult.Saved"/> before applying
    /// other settings fields.
    /// </summary>
    SettingsDialogResult ShowSettings(SettingsDialogInput input);

    /// <summary>
    /// Shows the database corruption dialog modally and returns the user's choice.
    /// </summary>
    CorruptionDialogResult ShowDatabaseCorruption(bool showRecoverOption);

    /// <summary>
    /// Shows the non-interactive "Recovering…" overlay window and returns an action
    /// that closes it. The caller is responsible for invoking the action on the UI thread.
    /// </summary>
    Action ShowDatabaseRecovering();
}
