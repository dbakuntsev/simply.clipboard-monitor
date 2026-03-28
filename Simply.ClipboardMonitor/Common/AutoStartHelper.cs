using Microsoft.Win32;

namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Manages the application's entry in the Windows current-user auto-start registry key
/// (<c>HKCU\Software\Microsoft\Windows\CurrentVersion\Run</c>).
/// </summary>
internal static class AutoStartHelper
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string AppName    = "Simply.ClipboardMonitor";

    /// <summary>
    /// Adds or removes the application from the auto-start registry key.
    /// </summary>
    internal static void SetAutoStart(bool enable)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
        if (key == null) return;

        if (enable)
        {
            var exePath = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(exePath))
                key.SetValue(AppName, $"\"{exePath}\"");
        }
        else
        {
            key.DeleteValue(AppName, throwOnMissingValue: false);
        }
    }

    /// <summary>
    /// Returns true if the application's auto-start registry entry currently exists.
    /// Use this on startup to read back the actual system state.
    /// </summary>
    internal static bool IsAutoStartEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        return key?.GetValue(AppName) != null;
    }
}
