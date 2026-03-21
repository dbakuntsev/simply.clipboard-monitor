using Simply.ClipboardMonitor.Services;
using System.Windows;

namespace Simply.ClipboardMonitor;

public partial class SettingsDialog : Window
{
    private readonly IHistoryRepository _history;

    /// <summary>Validated maximum entry count; meaningful only when DialogResult is true.</summary>
    public int MaxEntries { get; private set; }

    /// <summary>Validated maximum database size in MB; meaningful only when DialogResult is true.</summary>
    public int MaxSizeMb { get; private set; }

    /// <summary>True if the user pressed "Clear History" at any point during this dialog session.</summary>
    public bool HistoryWasCleared { get; private set; }

    internal SettingsDialog(int maxEntries, int maxSizeMb, IHistoryRepository history)
    {
        _history = history;
        InitializeComponent();
        MaxEntriesBox.Text = maxEntries.ToString();
        MaxSizeMbBox.Text  = maxSizeMb.ToString();
        RefreshDbSizeDisplay();
    }

    private void RefreshDbSizeDisplay()
    {
        DbSizeTextBlock.Text = FormatFileSize(_history.GetDatabaseFileSize());
    }

    private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        _history.ClearHistory();
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
