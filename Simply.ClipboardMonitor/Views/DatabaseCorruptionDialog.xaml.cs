using System.ComponentModel;
using System.Drawing;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor;

/// <summary>The choice the user made in <see cref="DatabaseCorruptionDialog"/>.</summary>
internal enum CorruptionDialogResult { Recover, DeleteAndReset, Disable }

/// <summary>
/// Modal dialog shown when <c>history.db</c> fails an integrity check.
/// Used for both the initial three-option prompt and the two-option prompt
/// shown after a failed recovery attempt.
/// </summary>
public partial class DatabaseCorruptionDialog : Window
{
    /// <summary>The action chosen by the user. Defaults to <see cref="CorruptionDialogResult.Disable"/>.</summary>
    internal CorruptionDialogResult Result { get; private set; } = CorruptionDialogResult.Disable;

    internal DatabaseCorruptionDialog(bool showRecoverOption)
    {
        InitializeComponent();

        WarningIcon.Source = Imaging.CreateBitmapSourceFromHIcon(
            SystemIcons.Warning.Handle,
            System.Windows.Int32Rect.Empty,
            BitmapSizeOptions.FromEmptyOptions());

        if (showRecoverOption)
        {
            Title = "History Database Corrupted";
            MessageTextBlock.Text =
                "The clipboard history database (history.db) is corrupted and cannot be read.";
        }
        else
        {
            Title = "Database Recovery Failed";
            MessageTextBlock.Text =
                "The database could not be repaired.";
            RecoverButton.Visibility = Visibility.Collapsed;
        }
    }

    private void RecoverButton_Click(object sender, RoutedEventArgs e)
    {
        Result = CorruptionDialogResult.Recover;
        Close();
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        Result = CorruptionDialogResult.DeleteAndReset;
        Close();
    }

    private void DisableButton_Click(object sender, RoutedEventArgs e)
    {
        Result = CorruptionDialogResult.Disable;
        Close();
    }

    // Closing via the title-bar X is treated as "Disable" (the safe default).
    protected override void OnClosing(CancelEventArgs e) => base.OnClosing(e);
}
