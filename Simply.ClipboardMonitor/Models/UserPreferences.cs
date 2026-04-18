namespace Simply.ClipboardMonitor.Models;

public sealed class UserPreferences
{
    public string? SortProperty { get; set; }
    public string? SortDirection { get; set; }
    public List<FormatColumnPreference>? FormatColumns { get; set; }
    public bool MonitorChanges { get; set; } = true;
    public bool TrackHistory { get; set; }
    public int  HistoryMaxEntries      { get; set; } = 100;
    public int  HistoryMaxSizeMb       { get; set; } = 100;
    public bool MinimizeToSystemTray   { get; set; }
    public bool TrayBalloonShown       { get; set; }
    public bool StartAtLogin           { get; set; }
    public bool StartMinimized         { get; set; }
    public bool TextWordWrap           { get; set; }
}

