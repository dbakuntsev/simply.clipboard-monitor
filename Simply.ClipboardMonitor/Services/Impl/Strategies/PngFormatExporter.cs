using Simply.ClipboardMonitor.Models;
using System.IO;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Exports the current image preview as a PNG file.
/// Available only when an image preview has been decoded for the selected format.
/// </summary>
internal sealed class PngFormatExporter : IFormatExporter
{
    public string Extension   => ".png";
    public string FilterLabel => "PNG Image (*.png)|*.png";

    public bool CanExport(FormatExportContext ctx) =>
        ctx.ImagePreviewSource != null;

    public void Export(string path, FormatExportContext ctx)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(ctx.ImagePreviewSource!));
        using var fs = File.Create(path);
        encoder.Save(fs);
    }
}
