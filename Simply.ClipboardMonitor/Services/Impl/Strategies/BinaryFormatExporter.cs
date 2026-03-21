using Simply.ClipboardMonitor.Models;
using System.IO;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Exports the raw clipboard format bytes as a binary file.
/// Always available as a fallback when no more specific format applies.
/// </summary>
internal sealed class BinaryFormatExporter : IFormatExporter
{
    public string Extension   => ".bin";
    public string FilterLabel => "Binary Data (*.bin)|*.bin";

    public bool CanExport(FormatExportContext ctx) => true;

    public void Export(string path, FormatExportContext ctx) =>
        File.WriteAllBytes(path, ctx.Bytes);
}
