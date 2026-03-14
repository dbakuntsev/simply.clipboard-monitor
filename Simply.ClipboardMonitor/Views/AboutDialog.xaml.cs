using Simply.ClipboardMonitor.Common;
using System.Reflection;
using System.Windows;
using System.Windows.Navigation;

namespace Simply.ClipboardMonitor;

public partial class AboutDialog : Window
{
    public string AppVersion { get; }

    public AboutDialog()
    {
        var raw = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? string.Empty;
        AppVersion = $"v{raw.Split('+')[0]}";

        DataContext = this;
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
