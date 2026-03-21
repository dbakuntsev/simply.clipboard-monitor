using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Exports a clipboard format's raw bytes to a file in a specific output format.
/// </summary>
public interface IFormatExporter
{
    /// <summary>File extension this exporter produces, including the leading dot (e.g. ".png").</summary>
    string Extension { get; }

    /// <summary>Filter string for the SaveFileDialog (e.g. "PNG Image (*.png)|*.png").</summary>
    string FilterLabel { get; }

    /// <summary>
    /// Returns true when this exporter can produce meaningful output for the given context.
    /// Used to build the available filter list shown in the Save dialog.
    /// </summary>
    bool CanExport(FormatExportContext ctx);

    /// <summary>Writes the exported content to <paramref name="path"/>.</summary>
    void Export(string path, FormatExportContext ctx);
}
