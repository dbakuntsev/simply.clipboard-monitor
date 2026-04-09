using System.Diagnostics;
using System.Windows;

namespace Simply.ClipboardMonitor;

/// <summary>
/// Shown when an unhandled exception is caught. Displays the path to the error log
/// as a clickable hyperlink so the user can open the file directly.
/// </summary>
public partial class CrashDialog : Window
{
    private readonly string _logFilePath;

    internal CrashDialog(string logFilePath)
    {
        _logFilePath = logFilePath;
        InitializeComponent();
        LogFilePathRun.Text = logFilePath;
    }

    private void LogFileLink_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(_logFilePath) { UseShellExecute = true });
        }
        catch { /* best effort */ }
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
