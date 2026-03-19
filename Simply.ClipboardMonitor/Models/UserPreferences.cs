namespace Simply.ClipboardMonitor.Models;

internal sealed class UserPreferences
{
    public string? SortProperty { get; set; }
    public string? SortDirection { get; set; }
    public List<FormatColumnPreference>? FormatColumns { get; set; }
    public bool MonitorChanges { get; set; } = true;
    public bool TrackHistory { get; set; }
}

internal sealed class FormatColumnPreference
{
    public string Key { get; set; } = string.Empty;
    public double? Width { get; set; }
}
