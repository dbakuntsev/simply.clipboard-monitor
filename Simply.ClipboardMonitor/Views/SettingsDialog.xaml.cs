using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Services;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace Simply.ClipboardMonitor;

public partial class SettingsDialog : Window
{
    private readonly IHistoryMaintenance _history;

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

    internal SettingsDialog(
        int maxEntries, int maxSizeMb, IHistoryMaintenance history,
        bool minimizeToSystemTray, bool startAtLogin, bool startMinimized)
    {
        _history = history;
        InitializeComponent();
        MaxEntriesBox.Text                = maxEntries.ToString();
        MaxSizeMbBox.Text                 = maxSizeMb.ToString();
        MinimizeToTrayCheckBox.IsChecked  = minimizeToSystemTray;
        StartAtLoginCheckBox.IsChecked    = startAtLogin;
        StartMinimizedCheckBox.IsChecked  = startMinimized;
        DataDirectoryPathRun.Text         = DataDirectory;
        RefreshDbSizeDisplay();
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
        DialogResult         = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

}
