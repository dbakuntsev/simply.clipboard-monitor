using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Loads and saves <see cref="UserPreferences"/> to a JSON file.</summary>
public interface IPreferencesService
{
    /// <summary>Loads preferences from disk, returning defaults on any failure.</summary>
    UserPreferences Load();

    /// <summary>Persists <paramref name="preferences"/> to disk. Silently ignores errors.</summary>
    void Save(UserPreferences preferences);
}
