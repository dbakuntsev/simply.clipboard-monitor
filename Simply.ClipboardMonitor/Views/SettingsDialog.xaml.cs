using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Services;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace Simply.ClipboardMonitor;

public partial class SettingsDialog : Window
{
    private readonly IHistoryMaintenance _history;
    private readonly bool _initialHotkeyConflict;

    // Press-to-capture state
    private HotkeyBinding _hotkeyBinding;
    private bool _isCapturing;
    private bool _bindingChangedSinceOpen;

    private static readonly string DataDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Simply.ClipboardMonitor");

    /// <summary>Validated maximum entry count; meaningful only when DialogResult is true.</summary>
    public int MaxEntries { get; private set; }

    /// <summary>Validated maximum database size in MB; meaningful only when DialogResult is true.</summary>
    public int MaxSizeMb { get; private set; }

    /// <summary>True if the user pressed "Clear History" at any point during this dialog session.</summary>
    public bool HistoryWasCleared { get; private set; }

    /// <summary>Whether "Minimize to System Tray" was checked when OK was pressed.</summary>
    public bool MinimizeToSystemTray { get; private set; }

    /// <summary>Whether "Start at login" was checked when OK was pressed.</summary>
    public bool StartAtLogin { get; private set; }

    /// <summary>Whether "Start minimized" was checked when OK was pressed.</summary>
    public bool StartMinimized { get; private set; }

    /// <summary>Whether the global hotkey was enabled when OK was pressed.</summary>
    public bool HotkeyEnabled { get; private set; }

    /// <summary>The hotkey binding in effect when OK was pressed.</summary>
    public HotkeyBinding GlobalHotkeyBinding { get; private set; }

    internal SettingsDialog(
        int maxEntries, int maxSizeMb, IHistoryMaintenance history,
        bool minimizeToSystemTray, bool startAtLogin, bool startMinimized,
        bool hotkeyEnabled, HotkeyBinding hotkeyBinding, bool hotkeyConflict)
    {
        _history               = history;
        _hotkeyBinding         = hotkeyBinding;
        _initialHotkeyConflict = hotkeyConflict;

        InitializeComponent();

        MaxEntriesBox.Text                = maxEntries.ToString();
        MaxSizeMbBox.Text                 = maxSizeMb.ToString();
        MinimizeToTrayCheckBox.IsChecked  = minimizeToSystemTray;
        StartAtLoginCheckBox.IsChecked    = startAtLogin;
        StartMinimizedCheckBox.IsChecked  = startMinimized;
        HotkeyEnabledCheckBox.IsChecked   = hotkeyEnabled;
        HotkeyCaptureBox.Text             = hotkeyBinding.ToString();
        DataDirectoryPathRun.Text         = DataDirectory;
        RefreshDbSizeDisplay();
        UpdateConflictWarning();
    }

    private void RefreshDbSizeDisplay()
    {
        DbSizeTextBlock.Text = DisplayHelper.FormatFileSize(_history.GetDatabaseFileSize());
    }

    private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        _history.ClearHistory();
        HistoryWasCleared = true;
        RefreshDbSizeDisplay();
    }

    private void DataDirectoryLink_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", DataDirectory) { UseShellExecute = true });
        }
        catch { /* best effort */ }
    }

    private bool ValidatePositiveInt(System.Windows.Controls.TextBox box, string fieldName, out int value)
    {
        if (!int.TryParse(box.Text.Trim(), out value) || value < 1)
        {
            MessageBox.Show(
                $"{fieldName} must be a positive integer.",
                "Invalid Value", MessageBoxButton.OK, MessageBoxImage.Warning);
            box.Focus();
            box.SelectAll();
            return false;
        }
        return true;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidatePositiveInt(MaxEntriesBox, "Maximum number of entries", out var entries))
            return;

        if (!ValidatePositiveInt(MaxSizeMbBox, "Maximum database size", out var sizeMb))
            return;

        MaxEntries           = entries;
        MaxSizeMb            = sizeMb;
        MinimizeToSystemTray = MinimizeToTrayCheckBox.IsChecked  == true;
        StartAtLogin         = StartAtLoginCheckBox.IsChecked    == true;
        StartMinimized       = StartMinimizedCheckBox.IsChecked  == true;
        HotkeyEnabled        = HotkeyEnabledCheckBox.IsChecked   == true;
        GlobalHotkeyBinding  = _hotkeyBinding;
        DialogResult         = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    // ── Global hotkey ────────────────────────────────────────────────────────

    private void HotkeyEnabled_Changed(object sender, RoutedEventArgs e)
    {
        UpdateConflictWarning();
    }

    private void HotkeyCaptureBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        _isCapturing = true;
        HotkeyCaptureBox.Text = "Press a key combination…";
    }

    private void HotkeyCaptureBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (!_isCapturing) return;
        // User clicked away without completing a capture — revert to the current binding.
        _isCapturing = false;
        HotkeyCaptureBox.Text = _hotkeyBinding.ToString();
    }

    private void HotkeyCaptureBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (!_isCapturing) return;
        e.Handled = true;

        // Alt+key combinations report Key.System; recover the real key from SystemKey.
        var key = e.Key == Key.System ? e.SystemKey : e.Key;

        if (key == Key.Escape)
        {
            _isCapturing = false;
            HotkeyCaptureBox.Text = _hotkeyBinding.ToString();
            return;
        }

        if (IsModifierKey(key))
        {
            // Show the partial combination while the user is still holding modifiers.
            UpdateCapturePlaceholder();
            return;
        }

        // A non-modifier key was pressed.  Validate that there is at least one of
        // Alt, Ctrl or Win — bare keys and Shift-only combos are rejected.
        var mods = GetCurrentModifiers();
        if ((mods & (HotkeyBinding.MOD_ALT | HotkeyBinding.MOD_CONTROL | HotkeyBinding.MOD_WIN)) == 0)
        {
            HotkeyCaptureBox.Text = "Press a key combination…";
            return;
        }

        var vk = (uint)KeyInterop.VirtualKeyFromKey(key);
        if (vk == 0)
        {
            HotkeyCaptureBox.Text = "Press a key combination…";
            return;
        }

        // Accept the new binding.
        _hotkeyBinding           = new HotkeyBinding { Modifiers = mods, VirtualKey = vk };
        _bindingChangedSinceOpen = true;
        _isCapturing             = false;
        HotkeyCaptureBox.Text    = _hotkeyBinding.ToString();
        UpdateConflictWarning();

        // Move focus to OK so the user can immediately press Enter to confirm.
        Dispatcher.BeginInvoke(() => OkButton.Focus());
    }

    private void HotkeyCaptureBox_PreviewKeyUp(object sender, KeyEventArgs e)
    {
        if (!_isCapturing) return;
        e.Handled = true;
        UpdateCapturePlaceholder();
    }

    // ── Hotkey helpers ────────────────────────────────────────────────────────

    /// <summary>
    /// Updates the capture field text to reflect whichever modifier keys are currently held.
    /// </summary>
    private void UpdateCapturePlaceholder()
    {
        var mods = GetCurrentModifiers();
        HotkeyCaptureBox.Text = mods != 0
            ? HotkeyBinding.FormatModifiers(mods) + "+…"
            : "Press a key combination…";
    }

    /// <summary>
    /// Shows the conflict warning only when the hotkey is enabled, a conflict was reported
    /// on entry, and the binding has not been changed during this dialog session.
    /// </summary>
    private void UpdateConflictWarning()
    {
        HotkeyConflictText.Visibility =
            _initialHotkeyConflict &&
            !_bindingChangedSinceOpen &&
            HotkeyEnabledCheckBox.IsChecked == true
                ? Visibility.Visible
                : Visibility.Collapsed;
    }

    private static bool IsModifierKey(Key key) =>
        key is Key.LeftAlt   or Key.RightAlt
            or Key.LeftCtrl  or Key.RightCtrl
            or Key.LeftShift or Key.RightShift
            or Key.LWin      or Key.RWin;

    /// <summary>
    /// Reads the current physical state of all modifier keys.
    /// Uses <see cref="Keyboard.IsKeyDown"/> rather than <see cref="ModifierKeys"/>
    /// so Win key state is captured reliably.
    /// </summary>
    private static uint GetCurrentModifiers()
    {
        uint mods = 0;
        if (Keyboard.IsKeyDown(Key.LeftAlt)   || Keyboard.IsKeyDown(Key.RightAlt))   mods |= HotkeyBinding.MOD_ALT;
        if (Keyboard.IsKeyDown(Key.LeftCtrl)  || Keyboard.IsKeyDown(Key.RightCtrl))  mods |= HotkeyBinding.MOD_CONTROL;
        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) mods |= HotkeyBinding.MOD_SHIFT;
        if (Keyboard.IsKeyDown(Key.LWin)      || Keyboard.IsKeyDown(Key.RWin))       mods |= HotkeyBinding.MOD_WIN;
        return mods;
    }

}
