namespace Simply.ClipboardMonitor.Common;

internal sealed class HexRow(string offset, string hex, string ascii)
{
    public string Offset { get; } = offset;
    public string Hex    { get; } = hex;
    public string Ascii  { get; } = ascii;
}
