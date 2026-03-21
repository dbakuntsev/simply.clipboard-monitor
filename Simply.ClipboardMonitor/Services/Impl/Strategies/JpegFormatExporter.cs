using Simply.ClipboardMonitor.Models;
using System.IO;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor.Services.Impl.Strategies;

/// <summary>
/// Exports the current image preview as a JPEG file.
/// When the source format already contains JPEG-encoded bytes, they are written directly;
/// otherwise the image preview is re-encoded at quality level 80.
/// Available only when an image preview has been decoded for the selected format.
/// </summary>
internal sealed class JpegFormatExporter : IFormatExporter
{
    public string Extension   => ".jpg";
    public string FilterLabel => "JPEG Image (*.jpg)|*.jpg";

    public bool CanExport(FormatExportContext ctx) =>
        ctx.ImagePreviewSource != null;

    public void Export(string path, FormatExportContext ctx)
    {
        var norm = ctx.FormatName.ToLowerInvariant();
        if (norm.Contains("jpeg") || norm.Contains("jpg"))
        {
            // Raw JPEG bytes can be written directly.
            File.WriteAllBytes(path, ctx.Bytes);
        }
        else
        {
            var encoder = new JpegBitmapEncoder { QualityLevel = 80 };
            encoder.Frames.Add(BitmapFrame.Create(ctx.ImagePreviewSource!));
            using var fs = File.Create(path);
            encoder.Save(fs);
        }
    }
}
