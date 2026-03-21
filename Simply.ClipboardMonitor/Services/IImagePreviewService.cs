using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Creates WPF BitmapSource previews from raw clipboard format bytes.</summary>
public interface IImagePreviewService
{
    /// <summary>True when the format is eligible for an image preview.</summary>
    bool IsImageCompatible(uint formatId, string formatName);

    /// <summary>
    /// Attempts to decode <paramref name="data"/> into a WPF BitmapSource.
    /// Returns false (with a non-null <paramref name="failureMessage"/>) on failure.
    /// </summary>
    bool TryCreatePreview(uint formatId, string formatName, byte[] data,
        out BitmapSource? preview, out string failureMessage);
}
