using System.Diagnostics;

namespace Simply.ClipboardMonitor.Common;

internal static class ShellHelper
{
    internal static void OpenUrl(string url)
    {
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }
}
