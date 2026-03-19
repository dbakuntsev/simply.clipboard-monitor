using Simply.ClipboardMonitor.Common;
using System.Windows;

namespace Simply.ClipboardMonitor;

public partial class SettingsDialog : Window
{
    /// <summary>Validated maximum entry count; meaningful only when DialogResult is true.</summary>
    public int MaxEntries { get; private set; }

    /// <summary>Validated maximum database size in MB; meaningful only when DialogResult is true.</summary>
    public int MaxSizeMb { get; private set; }

    /// <summary>True if the user pressed "Clear History" at any point during this dialog session.</summary>
    public bool HistoryWasCleared { get; private set; }

    public SettingsDialog(int maxEntries, int maxSizeMb)
    {
        InitializeComponent();
        MaxEntriesBox.Text = maxEntries.ToString();
        MaxSizeMbBox.Text  = maxSizeMb.ToString();
        RefreshDbSizeDisplay();
    }

    private void RefreshDbSizeDisplay()
    {
        DbSizeTextBlock.Text = FormatFileSize(ClipboardHistory.GetDatabaseFileSize());
    }

    private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        ClipboardHistory.ClearHistory();
        HistoryWasCleared = true;
        RefreshDbSizeDisplay();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(MaxEntriesBox.Text.Trim(), out var entries) || entries < 1)
        {
            MessageBox.Show(
                "Maximum number of entries must be a positive integer.",
                "Invalid Value", MessageBoxButton.OK, MessageBoxImage.Warning);
            MaxEntriesBox.Focus();
            MaxEntriesBox.SelectAll();
            return;
        }

        if (!int.TryParse(MaxSizeMbBox.Text.Trim(), out var sizeMb) || sizeMb < 1)
        {
            MessageBox.Show(
                "Maximum database size must be a positive integer.",
                "Invalid Value", MessageBoxButton.OK, MessageBoxImage.Warning);
            MaxSizeMbBox.Focus();
            MaxSizeMbBox.SelectAll();
            return;
        }

        MaxEntries   = entries;
        MaxSizeMb    = sizeMb;
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private static string FormatFileSize(long bytes) => bytes switch
    {
        0                      => "Not created yet",
        >= 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024 * 1024):F1} GB",
        >= 1024L * 1024        => $"{bytes / (1024.0 * 1024):F1} MB",
        >= 1024L               => $"{bytes / 1024.0:F1} KB",
        _                      => $"{bytes} B",
    };
}
