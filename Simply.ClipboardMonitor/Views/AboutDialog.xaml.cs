using Simply.ClipboardMonitor.Common;
using System.Windows;
using System.Windows.Navigation;

namespace Simply.ClipboardMonitor;

public partial class AboutDialog : Window
{
    public AboutDialog()
    {
        InitializeComponent();
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        ShellHelper.OpenUrl(e.Uri.AbsoluteUri);
        e.Handled = true;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
