using Simply.ClipboardMonitor.Models;
using System.IO;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Exports clipboard text bytes as a UTF-8 or user-selected encoding text file.
/// Available only when an encoding has been auto-detected for the selected format.
/// </summary>
internal sealed class TextFormatExporter : IFormatExporter
{
    public string Extension  => ".txt";
    public string FilterLabel => "Text (*.txt)|*.txt";

    public bool CanExport(FormatExportContext ctx) =>
        ctx.AutoDetectedEncoding != null;

    public void Export(string path, FormatExportContext ctx)
    {
        var encoding = ctx.ManuallySelectedEncoding ?? ctx.AutoDetectedEncoding!;
        var text     = encoding.GetString(ctx.Bytes).TrimEnd('\0');
        File.WriteAllText(path, text, encoding);
    }
}
