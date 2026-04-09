using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.IO;
using System.Text.Json;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Loads and saves <see cref="UserPreferences"/> to a JSON file stored under
/// <c>%LOCALAPPDATA%\Simply.ClipboardMonitor\preferences.json</c>.
/// All I/O errors are swallowed so that a corrupt or missing file never prevents
/// the application from starting.
/// </summary>
internal sealed class PreferencesService : IPreferencesService
{
    private const string PreferencesFileName = "preferences.json";

    /// <inheritdoc/>
    public UserPreferences Load()
    {
        try
        {
            var path = GetPreferencesFilePath();
            if (!File.Exists(path))
                return new UserPreferences();

            var json        = File.ReadAllText(path);
            var preferences = JsonSerializer.Deserialize<UserPreferences>(json);
            return preferences ?? new UserPreferences();
        }
        catch (Exception ex)
        {
            if (ex is not FileNotFoundException and not DirectoryNotFoundException)
                ErrorLogger.Log(ex);
            return new UserPreferences();
        }
    }

    /// <inheritdoc/>
    public void Save(UserPreferences preferences)
    {
        try
        {
            var path      = GetPreferencesFilePath();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            var json = JsonSerializer.Serialize(preferences,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch (Exception ex)
        {
            if (ex is not FileNotFoundException and not DirectoryNotFoundException)
                ErrorLogger.Log(ex);
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    private static string GetPreferencesFilePath()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(appData, "Simply.ClipboardMonitor", PreferencesFileName);
    }
}
