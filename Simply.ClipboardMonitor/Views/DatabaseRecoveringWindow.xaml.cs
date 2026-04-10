using System.ComponentModel;
using System.Windows;

namespace Simply.ClipboardMonitor;

/// <summary>
/// Non-interactive window shown while the history database is being recovered.
/// Cannot be closed by the user; call <see cref="CloseWindow"/> from the orchestrator
/// once recovery is complete.
/// </summary>
public partial class DatabaseRecoveringWindow : Window
{
    private bool _allowClose;

    public DatabaseRecoveringWindow() => InitializeComponent();

    /// <summary>Allows the window to close and closes it.</summary>
    internal void CloseWindow()
    {
        _allowClose = true;
        Close();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!_allowClose)
            e.Cancel = true;
        else
            base.OnClosing(e);
    }
}
