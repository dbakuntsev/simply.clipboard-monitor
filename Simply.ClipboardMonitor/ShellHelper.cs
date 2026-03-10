using System.Diagnostics;

namespace Simply.ClipboardMonitor;

internal static class ShellHelper
{
    internal static void OpenUrl(string url)
    {
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }
}
