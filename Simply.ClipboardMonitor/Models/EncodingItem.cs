using System.Text;

namespace Simply.ClipboardMonitor.Models;

/// <summary>Encoding item for display in the Encoding combo-box.</summary>
public sealed record EncodingItem(Encoding Encoding, string DisplayName)
{
    public override string ToString() => DisplayName;
}
