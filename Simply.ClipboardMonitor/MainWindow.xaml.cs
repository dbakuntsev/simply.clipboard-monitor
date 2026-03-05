using System.Collections.ObjectModel;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace Simply.ClipboardMonitor;

public partial class MainWindow : Window
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    private const int ERROR_NOT_LOCKED = 158;
    private const int BYTES_PER_ROW = 16;

    private static readonly Dictionary<uint, string> WellKnownFormats = new()
    {
        [1] = "CF_TEXT",
        [2] = "CF_BITMAP",
        [3] = "CF_METAFILEPICT",
        [4] = "CF_SYLK",
        [5] = "CF_DIF",
        [6] = "CF_TIFF",
        [7] = "CF_OEMTEXT",
        [8] = "CF_DIB",
        [9] = "CF_PALETTE",
        [10] = "CF_PENDATA",
        [11] = "CF_RIFF",
        [12] = "CF_WAVE",
        [13] = "CF_UNICODETEXT",
        [14] = "CF_ENHMETAFILE",
        [15] = "CF_HDROP",
        [16] = "CF_LOCALE",
        [17] = "CF_DIBV5",
        [0x0080] = "CF_OWNERDISPLAY",
        [0x0081] = "CF_DSPTEXT",
        [0x0082] = "CF_DSPBITMAP",
        [0x0083] = "CF_DSPMETAFILEPICT",
        [0x008E] = "CF_DSPENHMETAFILE"
    };

    private static readonly HashSet<uint> NonGlobalMemoryFormats =
    [
        2,       // CF_BITMAP
        3,       // CF_METAFILEPICT
        9,       // CF_PALETTE
        14,      // CF_ENHMETAFILE
        0x0082,  // CF_DSPBITMAP
        0x0083,  // CF_DSPMETAFILEPICT
        0x008E   // CF_DSPENHMETAFILE
    ];

    private readonly ObservableCollection<ClipboardFormatItem> _formats = [];
    private HwndSource? _hwndSource;
    private double _fitScale = 1.0;
    private bool _ignoreZoomChanges;
    private bool _isUiReady;

    public MainWindow()
    {
        InitializeComponent();
        _isUiReady = true;
        FormatListBox.ItemsSource = _formats;
        InitializePreviewState();
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        if (PresentationSource.FromVisual(this) is HwndSource source)
        {
            _hwndSource = source;
            _hwndSource.AddHook(WndProc);
            NativeMethods.AddClipboardFormatListener(source.Handle);
        }

        RefreshFormats();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (_hwndSource != null)
        {
            NativeMethods.RemoveClipboardFormatListener(_hwndSource.Handle);
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }

        base.OnClosed(e);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_CLIPBOARDUPDATE)
        {
            RefreshFormats();
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void RefreshFormats()
    {
        InitializePreviewState();

        var selectedId = (FormatListBox.SelectedItem as ClipboardFormatItem)?.Id;
        _formats.Clear();

        foreach (var item in EnumerateClipboardFormats())
        {
            _formats.Add(item);
        }

        if (selectedId.HasValue)
        {
            var matching = _formats.FirstOrDefault(f => f.Id == selectedId.Value);
            if (matching != null)
            {
                FormatListBox.SelectedItem = matching;
            }
        }
    }

    private void FormatListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        InitializePreviewState();

        if (FormatListBox.SelectedItem is not ClipboardFormatItem item)
        {
            return;
        }

        byte[]? bytes = null;
        string hexFailureMessage = "Clipboard data is unavailable.";
        BitmapSource? imagePreview = null;
        string imageFailureMessage = "Image preview unavailable for this format.";

        if (!TryOpenClipboard(IntPtr.Zero))
        {
            SetHexPreviewUnavailable("Unable to open clipboard.");
            return;
        }

        try
        {
            if (!TryReadClipboardDataBytes((uint)item.Id, out bytes, out hexFailureMessage))
            {
                bytes = null;
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        if (bytes != null)
        {
            SetHexPreview(bytes);
            UpdateTextPreview((uint)item.Id, item.Name, bytes);

            if (TryCreateImagePreview((uint)item.Id, item.Name, bytes, out imagePreview, out imageFailureMessage) &&
                imagePreview != null)
            {
                SetImagePreview(imagePreview);
            }
            else
            {
                SetImagePreviewUnavailable(imageFailureMessage);
            }
        }
        else
        {
            SetHexPreviewUnavailable(hexFailureMessage);
            SetTextPreviewUnavailable("Text preview requires byte-addressable clipboard data.");
            SetImagePreviewUnavailable("Image preview requires byte-addressable clipboard data.");
        }

    }

    private IEnumerable<ClipboardFormatItem> EnumerateClipboardFormats()
    {
        var results = new List<ClipboardFormatItem>();

        if (!TryOpenClipboard(IntPtr.Zero))
        {
            return results;
        }

        try
        {
            uint format = 0;
            while (true)
            {
                format = NativeMethods.EnumClipboardFormats(format);
                if (format == 0)
                {
                    break;
                }

                var name = GetFormatDisplayName(format);
                var sizeText = TryGetClipboardDataSize(format, out var sizeBytes)
                    ? $"{sizeBytes:N0} bytes"
                    : "size unavailable";

                results.Add(new ClipboardFormatItem((int)format, name, $"{format}: {name} ({sizeText})"));
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        return results;
    }

    private void SetHexPreview(byte[] data)
    {
        HexStatusTextBlock.Text = $"Showing {data.Length:N0} bytes ({((data.Length + BYTES_PER_ROW - 1) / BYTES_PER_ROW):N0} rows).";
        HexListView.ItemsSource = new HexRowCollection(data);
    }

    private void SetHexPreviewUnavailable(string message)
    {
        HexStatusTextBlock.Text = message;
        HexListView.ItemsSource = null;
    }

    private void UpdateTextPreview(uint formatId, string formatName, byte[] bytes)
    {
        if (!TryDecodeTextContent(formatId, formatName, bytes, out var text, out var message))
        {
            SetTextPreviewUnavailable(message);
            return;
        }

        TextStatusTextBlock.Text = "Decoded text preview.";
        TextContentTextBox.Text = text;
    }

    private static bool TryDecodeTextContent(uint formatId, string formatName, byte[] bytes, out string text, out string failureMessage)
    {
        var normalized = formatName.ToLowerInvariant();
        var isTextFormat = formatId is 1 or 7 or 13 ||
                           normalized.Contains("text", StringComparison.Ordinal) ||
                           normalized.Contains("html", StringComparison.Ordinal) ||
                           normalized.Contains("rtf", StringComparison.Ordinal) ||
                           normalized.Contains("xml", StringComparison.Ordinal) ||
                           normalized.Contains("json", StringComparison.Ordinal) ||
                           normalized.Contains("csv", StringComparison.Ordinal);

        if (!isTextFormat)
        {
            text = string.Empty;
            failureMessage = "Text preview unavailable for this format.";
            return false;
        }

        try
        {
            if (formatId == 13) // CF_UNICODETEXT
            {
                text = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }
            else if (formatId == 1) // CF_TEXT
            {
                text = Encoding.Default.GetString(TrimAtNull(bytes, 1));
            }
            else if (formatId == 7) // CF_OEMTEXT
            {
                var oemEncoding = Encoding.GetEncoding((int)NativeMethods.GetOEMCP());
                text = oemEncoding.GetString(TrimAtNull(bytes, 1));
            }
            else
            {
                text = DecodeWithFallback(bytes);
            }

            failureMessage = string.Empty;
            return true;
        }
        catch
        {
            text = string.Empty;
            failureMessage = "Failed to decode this format as text.";
            return false;
        }
    }

    private static byte[] TrimAtNull(byte[] bytes, int unitSize)
    {
        if (unitSize <= 1)
        {
            var index = Array.IndexOf(bytes, (byte)0);
            if (index < 0)
            {
                return bytes;
            }

            return bytes[..index];
        }

        for (var i = 0; i <= bytes.Length - unitSize; i += unitSize)
        {
            var allZero = true;
            for (var j = 0; j < unitSize; j++)
            {
                if (bytes[i + j] != 0)
                {
                    allZero = false;
                    break;
                }
            }

            if (allZero)
            {
                return bytes[..i];
            }
        }

        return bytes;
    }

    private static string DecodeWithFallback(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3).TrimEnd('\0');
        }

        try
        {
            return new UTF8Encoding(false, true).GetString(bytes).TrimEnd('\0');
        }
        catch
        {
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }
    }

    private static bool TryCreateImagePreview(uint formatId, string formatName, byte[] bytes, out BitmapSource? image, out string failureMessage)
    {
        image = null;
        var normalized = formatName.ToLowerInvariant();
        var isImageCandidate = formatId is 8 or 17 ||
                               normalized.Contains("png", StringComparison.Ordinal) ||
                               normalized.Contains("jpeg", StringComparison.Ordinal) ||
                               normalized.Contains("jpg", StringComparison.Ordinal) ||
                               normalized.Contains("gif", StringComparison.Ordinal) ||
                               normalized.Contains("dib", StringComparison.Ordinal) ||
                               normalized.Contains("bitmap", StringComparison.Ordinal) ||
                               normalized.Contains("image", StringComparison.Ordinal);

        if (!isImageCandidate)
        {
            failureMessage = "Image preview unavailable for this format.";
            return false;
        }

        if (bytes.Length == 0)
        {
            failureMessage = "Image data is empty.";
            return false;
        }

        try
        {
            if (formatId is 8 or 17)
            {
                if (!TryCreateBitmapFromDib(bytes, out image))
                {
                    failureMessage = "Failed to decode DIB image data.";
                    return false;
                }
            }
            else
            {
                image = CreateBitmapFromEncodedImage(bytes);
            }

            failureMessage = string.Empty;
            return image != null;
        }
        catch
        {
            image = null;
            failureMessage = "Failed to decode image preview for this format.";
            return false;
        }
    }

    private static BitmapSource CreateBitmapFromEncodedImage(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
        var frame = decoder.Frames[0];
        frame.Freeze();
        return frame;
    }

    private static bool TryCreateBitmapFromDib(byte[] dibBytes, out BitmapSource? bitmap)
    {
        bitmap = null;
        if (dibBytes.Length < 40)
        {
            return false;
        }

        var headerSize = BitConverter.ToUInt32(dibBytes, 0);
        if (headerSize < 40 || headerSize > dibBytes.Length)
        {
            return false;
        }

        var bitCount = BitConverter.ToUInt16(dibBytes, 14);
        var compression = BitConverter.ToUInt32(dibBytes, 16);
        var colorsUsed = BitConverter.ToUInt32(dibBytes, 32);

        uint masksSize = 0;
        if ((compression == 3 || compression == 6) && headerSize == 40)
        {
            masksSize = compression == 6 ? 16u : 12u;
        }

        uint colorTableEntries = colorsUsed;
        if (colorTableEntries == 0 && bitCount <= 8)
        {
            colorTableEntries = 1u << bitCount;
        }

        var colorTableSize = colorTableEntries * 4;
        var pixelOffset = 14u + headerSize + masksSize + colorTableSize;
        if (pixelOffset > dibBytes.Length + 14u)
        {
            return false;
        }

        var fileBytes = new byte[dibBytes.Length + 14];
        fileBytes[0] = (byte)'B';
        fileBytes[1] = (byte)'M';
        WriteUInt32(fileBytes, 2, (uint)fileBytes.Length);
        WriteUInt32(fileBytes, 10, pixelOffset);
        Buffer.BlockCopy(dibBytes, 0, fileBytes, 14, dibBytes.Length);

        bitmap = CreateBitmapFromEncodedImage(fileBytes);
        return true;
    }

    private static void WriteUInt32(byte[] buffer, int offset, uint value)
    {
        buffer[offset] = (byte)(value & 0xFF);
        buffer[offset + 1] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private void SetImagePreview(BitmapSource image)
    {
        ImagePreview.Source = image;
        ImageStatusTextBlock.Text = string.Empty;
        ZoomSlider.IsEnabled = true;
        SetZoomValue(1);
        UpdateFitScale();
    }

    private void SetImagePreviewUnavailable(string message)
    {
        ImagePreview.Source = null;
        ImageStatusTextBlock.Text = message;
        ZoomSlider.IsEnabled = false;
        SetZoomValue(1);
        ApplyImageScale(1);
    }

    private void SetTextPreviewUnavailable(string message)
    {
        TextStatusTextBlock.Text = message;
        TextContentTextBox.Text = string.Empty;
    }

    private void InitializePreviewState()
    {
        SetHexPreviewUnavailable("Select a clipboard format to preview.");
        SetTextPreviewUnavailable("Text preview unavailable for this format.");
        SetImagePreviewUnavailable("Image preview unavailable for this format.");
    }

    private void ZoomSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (_ignoreZoomChanges || !_isUiReady)
        {
            return;
        }

        ApplyImageScale();
    }

    private void ResetZoomButton_OnClick(object sender, RoutedEventArgs e)
    {
        SetZoomValue(1);
        ApplyImageScale();
    }

    private void ImageScrollViewer_OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateFitScale();
    }

    private void UpdateFitScale()
    {
        if (ImagePreview.Source is not BitmapSource source || source.PixelWidth <= 0 || source.PixelHeight <= 0)
        {
            _fitScale = 1;
            ApplyImageScale(1);
            return;
        }

        var viewportWidth = ImageScrollViewer.ViewportWidth;
        var viewportHeight = ImageScrollViewer.ViewportHeight;
        if (viewportWidth <= 0 || viewportHeight <= 0)
        {
            _fitScale = 1;
            ApplyImageScale();
            return;
        }

        var fitX = viewportWidth / source.PixelWidth;
        var fitY = viewportHeight / source.PixelHeight;
        _fitScale = Math.Min(1.0, Math.Min(fitX, fitY));
        ApplyImageScale();
    }

    private void ApplyImageScale(double? explicitScale = null)
    {
        if (ImageScaleTransform == null || ZoomValueTextBlock == null || ZoomSlider == null)
        {
            return;
        }

        var scale = explicitScale ?? (_fitScale * ZoomSlider.Value);
        if (scale <= 0)
        {
            scale = 0.01;
        }

        ImageScaleTransform.ScaleX = scale;
        ImageScaleTransform.ScaleY = scale;
        ZoomValueTextBlock.Text = $"{scale * 100:0}%";
    }

    private void SetZoomValue(double value)
    {
        _ignoreZoomChanges = true;
        ZoomSlider.Value = value;
        _ignoreZoomChanges = false;
    }

    private static bool TryReadClipboardDataBytes(uint formatId, out byte[] bytes, out string failureMessage)
    {
        bytes = [];

        if (NonGlobalMemoryFormats.Contains(formatId))
        {
            failureMessage = "Selected format uses a non-memory handle type (not HGLOBAL).";
            return false;
        }

        var handle = NativeMethods.GetClipboardData(formatId);
        if (handle == IntPtr.Zero)
        {
            failureMessage = "No data handle is available for this format.";
            return false;
        }

        var globalSize = NativeMethods.GlobalSize(handle);
        if (globalSize == UIntPtr.Zero)
        {
            failureMessage = "Clipboard format size is unavailable.";
            return false;
        }

        var size64 = (ulong)globalSize;
        if (size64 > int.MaxValue)
        {
            failureMessage = "Clipboard data is too large to render in this viewer.";
            return false;
        }

        var dataPtr = NativeMethods.GlobalLock(handle);
        if (dataPtr == IntPtr.Zero)
        {
            failureMessage = "Failed to lock clipboard data.";
            return false;
        }

        try
        {
            bytes = new byte[(int)size64];
            Marshal.Copy(dataPtr, bytes, 0, bytes.Length);
            failureMessage = string.Empty;
            return true;
        }
        finally
        {
            if (!NativeMethods.GlobalUnlock(handle))
            {
                var error = Marshal.GetLastWin32Error();
                if (error != 0 && error != ERROR_NOT_LOCKED)
                {
                    // Best effort unlock on clipboard-provided memory.
                }
            }
        }
    }

    private static bool TryGetClipboardDataSize(uint format, out ulong sizeBytes)
    {
        sizeBytes = 0;

        if (NonGlobalMemoryFormats.Contains(format))
        {
            return false;
        }

        var handle = NativeMethods.GetClipboardData(format);
        if (handle == IntPtr.Zero)
        {
            return false;
        }

        var size = NativeMethods.GlobalSize(handle);
        if (size == UIntPtr.Zero)
        {
            return false;
        }

        sizeBytes = (ulong)size;
        return true;
    }

    private static bool TryOpenClipboard(IntPtr owner)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            if (NativeMethods.OpenClipboard(owner))
            {
                return true;
            }

            Thread.Sleep(20);
        }

        return false;
    }

    private static string GetFormatDisplayName(uint format)
    {
        if (WellKnownFormats.TryGetValue(format, out var wellKnownName))
        {
            return wellKnownName;
        }

        var buffer = new StringBuilder(256);
        var chars = NativeMethods.GetClipboardFormatName(format, buffer, buffer.Capacity);
        if (chars > 0)
        {
            return buffer.ToString();
        }

        return "Unknown";
    }

    private sealed class ClipboardFormatItem(int id, string name, string displayName)
    {
        public int Id { get; } = id;
        public string Name { get; } = name;
        public string DisplayName { get; } = displayName;
    }

    private sealed class HexRowCollection(byte[] data) : IList
    {
        private readonly Dictionary<int, HexRow> _cache = [];

        public int Count => (data.Length + BYTES_PER_ROW - 1) / BYTES_PER_ROW;
        public bool IsReadOnly => true;
        public bool IsFixedSize => true;
        public bool IsSynchronized => false;
        public object SyncRoot => this;

        public object? this[int index]
        {
            get
            {
                if (index < 0 || index >= Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index));
                }

                if (_cache.TryGetValue(index, out var cached))
                {
                    return cached;
                }

                var offset = index * BYTES_PER_ROW;
                var rowLength = Math.Min(BYTES_PER_ROW, data.Length - offset);
                var hexBuilder = new StringBuilder();
                var asciiBuilder = new StringBuilder();

                for (var i = 0; i < BYTES_PER_ROW; i++)
                {
                    if (i < rowLength)
                    {
                        var b = data[offset + i];
                        hexBuilder.Append(b.ToString("X2"));
                        asciiBuilder.Append(b >= 32 && b <= 126 ? (char)b : '.');
                    }
                    else
                    {
                        hexBuilder.Append("  ");
                    }

                    if (i != BYTES_PER_ROW - 1)
                    {
                        hexBuilder.Append(' ');
                    }
                }

                var row = new HexRow(offset.ToString("X8"), hexBuilder.ToString(), asciiBuilder.ToString());
                _cache[index] = row;
                return row;
            }
            set => throw new NotSupportedException();
        }

        public int Add(object? value) => throw new NotSupportedException();
        public void Clear() => throw new NotSupportedException();
        public bool Contains(object? value) => false;
        public int IndexOf(object? value) => -1;
        public void Insert(int index, object? value) => throw new NotSupportedException();
        public void Remove(object? value) => throw new NotSupportedException();
        public void RemoveAt(int index) => throw new NotSupportedException();
        public void CopyTo(Array array, int index) => throw new NotSupportedException();
        public IEnumerator GetEnumerator()
        {
            for (var i = 0; i < Count; i++)
            {
                yield return this[i];
            }
        }
    }

    private sealed class HexRow(string offset, string hex, string ascii)
    {
        public string Offset { get; } = offset;
        public string Hex { get; } = hex;
        public string Ascii { get; } = ascii;
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool CloseClipboard();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint EnumClipboardFormats(uint format);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int GetClipboardFormatName(uint format, StringBuilder lpszFormatName, int cchMaxCount);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern UIntPtr GlobalSize(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern uint GetOEMCP();
    }
}
