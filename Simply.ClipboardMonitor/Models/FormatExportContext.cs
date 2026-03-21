using System.Text;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor.Models;

/// <summary>
/// Carries the state needed by an <see cref="Services.IFormatExporter"/> to write the
/// currently selected clipboard format to a file.
/// </summary>
public sealed record FormatExportContext(
    byte[]        Bytes,
    uint          FormatId,
    string        FormatName,
    Encoding?     AutoDetectedEncoding,
    Encoding?     ManuallySelectedEncoding,
    BitmapSource? ImagePreviewSource);
