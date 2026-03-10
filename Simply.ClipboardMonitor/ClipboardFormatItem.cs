namespace Simply.ClipboardMonitor;

internal sealed class ClipboardFormatItem(int ordinal, uint formatId, string name, string contentSize, long contentSizeValue)
{
    public int Ordinal { get; } = ordinal;
    public uint FormatId { get; } = formatId;
    public string FormatNumber => FormatId.ToString("D");
    public string Name { get; } = name;
    public string ContentSize { get; } = contentSize;
    public long ContentSizeValue { get; } = contentSizeValue;
}
