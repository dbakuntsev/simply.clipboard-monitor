using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.Buffers.Binary;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Simply.ClipboardMonitor;

public partial class MainWindow : Window
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    private const int ERROR_NOT_LOCKED = 158;
    private const string PreferencesFileName = "preferences.json";
    private const string DefaultSortProperty = nameof(ClipboardFormatItem.Ordinal);

    public static readonly RoutedUICommand RefreshClipboardCommand = new(
        "Refresh", "RefreshClipboard", typeof(MainWindow),
        new InputGestureCollection { new KeyGesture(Key.F5) });

    public static readonly RoutedUICommand LoadCommand = new(
        "Load", "Load", typeof(MainWindow),
        new InputGestureCollection { new KeyGesture(Key.O, ModifierKeys.Control) });

    public static readonly RoutedUICommand SaveCommand = new(
        "Save", "Save", typeof(MainWindow),
        new InputGestureCollection { new KeyGesture(Key.S, ModifierKeys.Control) });

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

    // CF_BITMAP / CF_DSPBITMAP: HBITMAP — converted to DIB bytes via GetDIBits.
    private static readonly HashSet<uint> HBitmapFormats = [2, 0x0082];

    // CF_ENHMETAFILE / CF_DSPENHMETAFILE: HENHMETAFILE — raw bytes via GetEnhMetaFileBits.
    private static readonly HashSet<uint> HEnhMetaFileFormats = [14, 0x008E];

    // Formats whose handles cannot be usefully read as raw bytes (HPALETTE, etc.).
    private static readonly HashSet<uint> NonGlobalMemoryFormats = [9]; // CF_PALETTE

    private readonly ObservableCollection<ClipboardFormatItem> _formats = [];
    private HwndSource? _hwndSource;
    private double _fitScale = 1.0;
    private bool _ignoreZoomChanges;
    // Set to true after InitializeComponent to suppress events fired during XAML initialization.
    private bool _isUiReady;
    private string _currentSortProperty = DefaultSortProperty;
    private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
    private List<FormatColumnPreference> _storedColumnPreferences = [];
    private bool _isMonitoring;
    private bool _isPanning;
    private Point _panStartMouse;
    private double _panStartHorizontalOffset;
    private double _panStartVerticalOffset;
    // Tracks the path of the most recently loaded or saved .clipdb file.
    private string? _currentFilePath;

    public MainWindow()
    {
        InitializeComponent();
        CommandBindings.Add(new CommandBinding(RefreshClipboardCommand, RefreshClipboardCommand_Executed, RefreshClipboardCommand_CanExecute));
        CommandBindings.Add(new CommandBinding(LoadCommand, LoadCommand_Executed));
        CommandBindings.Add(new CommandBinding(SaveCommand, SaveCommand_Executed));
        _isUiReady = true;
        FormatListBox.ItemsSource = _formats;
        LoadPreferences();
        MenuItemMonitorChanges.IsChecked = _isMonitoring;
        UpdateStatusBar();
        ApplyFormatColumnPreferences();
        FormatListBox.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler(FormatListBox_OnColumnHeaderClick));
        ApplyFormatSort();
        InitializePreviewState();
        PreviewKeyDown += (_, _) => UpdateImageScrollViewerCursor();
        PreviewKeyUp   += (_, _) => UpdateImageScrollViewerCursor();
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

        SavePreferences();
        base.OnClosed(e);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_CLIPBOARDUPDATE)
        {
            if (_isMonitoring)
            {
                RefreshFormats();
            }
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void RefreshFormats()
    {
        InitializePreviewState();

        var selectedId = (FormatListBox.SelectedItem as ClipboardFormatItem)?.FormatId;
        _formats.Clear();

        foreach (var item in EnumerateClipboardFormats())
        {
            _formats.Add(item);
        }

        if (selectedId.HasValue)
        {
            var matching = _formats.FirstOrDefault(f => f.FormatId == selectedId.Value);
            if (matching != null)
            {
                FormatListBox.SelectedItem = matching;
            }
        }
    }

    private void FormatListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FormatListBox.SelectedItem is not ClipboardFormatItem item)
        {
            return;
        }

        ContentTabControl.Visibility = Visibility.Visible;
        NoSelectionPanel.Visibility = Visibility.Collapsed;
        TextTabItem.IsEnabled = IsTextCompatible(item.FormatId, item.Name);
        ImageTabItem.IsEnabled = IsImageCompatible(item.FormatId, item.Name);

        if (!TextTabItem.IsEnabled && ContentTabControl.SelectedItem == TextTabItem)
            ContentTabControl.SelectedItem = HexTabItem;
        if (!ImageTabItem.IsEnabled && ContentTabControl.SelectedItem == ImageTabItem)
            ContentTabControl.SelectedItem = HexTabItem;

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
            TryReadClipboardDataBytes(item.FormatId, out bytes, out hexFailureMessage);
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        if (bytes != null)
        {
            SetHexPreview(bytes);
            UpdateTextPreview(item.FormatId, item.Name, bytes);

            if (TryCreateImagePreview(item.FormatId, item.Name, bytes, out imagePreview, out imageFailureMessage) &&
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
            var ordinal = 1;
            while (true)
            {
                format = NativeMethods.EnumClipboardFormats(format);
                if (format == 0)
                {
                    break;
                }

                var name = GetFormatDisplayName(format);
                var hasSize = TryGetClipboardDataSize(format, out var sizeBytes);
                var contentSizeValue = hasSize ? (long)sizeBytes : -1L;
                var contentSize = hasSize ? sizeBytes.ToString("N0") : "n/a";

                results.Add(new ClipboardFormatItem(ordinal, format, name, contentSize, contentSizeValue));
                ordinal++;
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
        HexStatusTextBlock.Text = $"Showing {data.Length:N0} bytes ({((data.Length + HexRowCollection.BytesPerRow - 1) / HexRowCollection.BytesPerRow):N0} rows).";
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
        if (!IsTextCompatible(formatId, formatName))
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
        if (!IsImageCompatible(formatId, formatName))
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
            if (formatId is 8 or 17 || HBitmapFormats.Contains(formatId))
            {
                // CF_DIB, CF_DIBV5, and HBITMAP formats all yield a BITMAPINFOHEADER + pixels block.
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

    private static bool IsTextCompatible(uint formatId, string formatName)
    {
        var normalized = formatName.ToLowerInvariant();
        return formatId is 1 or 7 or 13 ||
               normalized.Contains("text", StringComparison.Ordinal) ||
               normalized.Contains("html", StringComparison.Ordinal) ||
               normalized.Contains("rtf", StringComparison.Ordinal) ||
               normalized.Contains("xml", StringComparison.Ordinal) ||
               normalized.Contains("json", StringComparison.Ordinal) ||
               normalized.Contains("csv", StringComparison.Ordinal);
    }

    private static bool IsImageCompatible(uint formatId, string formatName)
    {
        // CF_DIB, CF_DIBV5, CF_BITMAP, CF_DSPBITMAP — all produce DIB bytes.
        if (formatId is 8 or 17 || HBitmapFormats.Contains(formatId))
            return true;

        var normalized = formatName.ToLowerInvariant();
        return normalized.Contains("png",    StringComparison.Ordinal) ||
               normalized.Contains("jpeg",   StringComparison.Ordinal) ||
               normalized.Contains("jpg",    StringComparison.Ordinal) ||
               normalized.Contains("gif",    StringComparison.Ordinal) ||
               normalized.Contains("dib",    StringComparison.Ordinal) ||
               normalized.Contains("bitmap", StringComparison.Ordinal) ||
               normalized.Contains("image",  StringComparison.Ordinal);
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
        BinaryPrimitives.WriteUInt32LittleEndian(fileBytes.AsSpan(2), (uint)fileBytes.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(fileBytes.AsSpan(10), pixelOffset);
        Buffer.BlockCopy(dibBytes, 0, fileBytes, 14, dibBytes.Length);

        bitmap = CreateBitmapFromEncodedImage(fileBytes);
        return true;
    }

    private void SetImagePreview(BitmapSource image)
    {
        ImagePreview.Source = image;
        ImageStatusTextBlock.Text = string.Empty;
        ZoomSlider.IsEnabled = true;
        ImageDimensionsWidthTextBlock.Text = $"{image.PixelWidth}px";
        ImageDimensionsHeightTextBlock.Text = $"{image.PixelHeight}px";
        ImageDimensionsStackPanel.Visibility = Visibility.Visible;
        SetZoomValue(1);
        UpdateFitScale();
    }

    private void SetImagePreviewUnavailable(string message)
    {
        ImagePreview.Source = null;
        ImageStatusTextBlock.Text = message;
        ZoomSlider.IsEnabled = false;
        ImageDimensionsStackPanel.Visibility = Visibility.Collapsed;
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
        TextTabItem.IsEnabled = true;
        ImageTabItem.IsEnabled = true;
        ContentTabControl.SelectedIndex = 0;
        ContentTabControl.Visibility = Visibility.Collapsed;
        NoSelectionPanel.Visibility = Visibility.Visible;
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
            ApplyImageScale();
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
        if (ImageScaleTransform == null || ZoomTextBox == null || ZoomSlider == null)
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
        ZoomTextBox.Text = $"{scale * 100:0}%";
    }

    private void SetZoomValue(double value)
    {
        _ignoreZoomChanges = true;
        ZoomSlider.Value = value;
        _ignoreZoomChanges = false;
    }

    private void ZoomTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        CommitZoomTextBox();
    }

    private void ZoomTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            CommitZoomTextBox();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            ZoomTextBox.Text = $"{_fitScale * ZoomSlider.Value * 100:0}%";
            ZoomTextBox.SelectAll();
            e.Handled = true;
        }
    }

    private void CommitZoomTextBox()
    {
        var text = ZoomTextBox.Text.Trim().TrimEnd('%').TrimEnd();
        if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var pct) && pct > 0)
        {
            ApplyZoomFromPercentage(pct);
        }
        else
        {
            ZoomTextBox.Text = $"{_fitScale * ZoomSlider.Value * 100:0}%";
        }
    }

    private void ApplyZoomFromPercentage(double pct)
    {
        var newSliderValue = _fitScale > 0 ? pct / 100.0 / _fitScale : pct / 100.0;
        newSliderValue = Math.Clamp(newSliderValue, ZoomSlider.Minimum, ZoomSlider.Maximum);
        ZoomSlider.Value = newSliderValue;
    }

    private void ZoomUpButton_Click(object sender, RoutedEventArgs e)
    {
        StepZoom(+1);
    }

    private void ZoomDownButton_Click(object sender, RoutedEventArgs e)
    {
        StepZoom(-1);
    }

    private void StepZoom(int direction)
    {
        var currentPct = Math.Round(_fitScale * ZoomSlider.Value * 100);
        ApplyZoomFromPercentage(currentPct + direction);
    }

    private void ImageScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == 0)
        {
            return;
        }

        if (ImagePreview.Source is not BitmapSource source)
        {
            return;
        }

        e.Handled = true;

        var mousePos    = e.GetPosition(ImageScrollViewer);
        var currentScale = _fitScale * ZoomSlider.Value;
        var contentX    = ImageScrollViewer.HorizontalOffset + mousePos.X;
        var contentY    = ImageScrollViewer.VerticalOffset   + mousePos.Y;
        var scaledW     = source.PixelWidth  * currentScale;
        var scaledH     = source.PixelHeight * currentScale;
        var imageLeft   = Math.Max(0.0, (ImageScrollViewer.ViewportWidth  - scaledW) / 2.0);
        var imageTop    = Math.Max(0.0, (ImageScrollViewer.ViewportHeight - scaledH) / 2.0);
        var imgLocalX   = (contentX - imageLeft) / currentScale;
        var imgLocalY   = (contentY - imageTop)  / currentScale;

        StepZoom(15 * Math.Sign(e.Delta));

        Dispatcher.BeginInvoke(DispatcherPriority.Background, () =>
        {
            var newScale   = _fitScale * ZoomSlider.Value;
            var newScaledW = source.PixelWidth  * newScale;
            var newScaledH = source.PixelHeight * newScale;
            var newImgLeft = Math.Max(0.0, (ImageScrollViewer.ViewportWidth  - newScaledW) / 2.0);
            var newImgTop  = Math.Max(0.0, (ImageScrollViewer.ViewportHeight - newScaledH) / 2.0);
            ImageScrollViewer.ScrollToHorizontalOffset(imgLocalX * newScale + newImgLeft - mousePos.X);
            ImageScrollViewer.ScrollToVerticalOffset  (imgLocalY * newScale + newImgTop  - mousePos.Y);
        });
    }

    private void ImageScrollViewer_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle)
        {
            return;
        }

        _isPanning = true;
        _panStartMouse = e.GetPosition(ImageScrollViewer);
        _panStartHorizontalOffset = ImageScrollViewer.HorizontalOffset;
        _panStartVerticalOffset   = ImageScrollViewer.VerticalOffset;
        ImageScrollViewer.Cursor = Cursors.ScrollAll;
        Mouse.Capture(ImageScrollViewer);
        e.Handled = true;
    }

    private void ImageScrollViewer_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle || !_isPanning)
        {
            return;
        }

        _isPanning = false;
        Mouse.Capture(null);
        UpdateImageScrollViewerCursor();
        e.Handled = true;
    }

    private void ImageScrollViewer_MouseMove(object sender, MouseEventArgs e)
    {
        if (_isPanning)
        {
            var pos = e.GetPosition(ImageScrollViewer);
            ImageScrollViewer.ScrollToHorizontalOffset(_panStartHorizontalOffset + (_panStartMouse.X - pos.X));
            ImageScrollViewer.ScrollToVerticalOffset  (_panStartVerticalOffset   + (_panStartMouse.Y - pos.Y));
            return;
        }

        UpdateImageScrollViewerCursor();
    }

    private void ImageScrollViewer_MouseLeave(object sender, MouseEventArgs e)
    {
        if (_isPanning)
        {
            return;
        }

        ImageScrollViewer.Cursor = null;
    }

    private void UpdateImageScrollViewerCursor()
    {
        if (_isPanning)
        {
            return;
        }

        if (!ImageScrollViewer.IsMouseOver)
        {
            return;
        }

        ImageScrollViewer.Cursor = (Keyboard.Modifiers & ModifierKeys.Control) != 0
            ? Cursors.Hand
            : null;
    }

    private static bool TryReadClipboardDataBytes(uint formatId, out byte[]? bytes, out string failureMessage)
    {
        bytes = null;

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

        if (HBitmapFormats.Contains(formatId))
            return TryReadHBitmapAsBytes(handle, out bytes, out failureMessage);

        if (HEnhMetaFileFormats.Contains(formatId))
            return TryReadEnhMetaFileAsBytes(handle, out bytes, out failureMessage);

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

    private static bool TryReadHBitmapAsBytes(IntPtr hBitmap, out byte[]? bytes, out string failureMessage)
    {
        bytes = null;

        if (NativeMethods.GetObject(hBitmap, Marshal.SizeOf<BITMAP>(), out var bm) == 0)
        {
            failureMessage = "Failed to read bitmap metadata.";
            return false;
        }

        // Request 32 bpp BI_RGB output so there is no colour table to worry about.
        var header = new BITMAPINFOHEADER
        {
            biSize        = (uint)Marshal.SizeOf<BITMAPINFOHEADER>(),
            biWidth       = bm.bmWidth,
            biHeight      = bm.bmHeight,   // positive → bottom-up DIB
            biPlanes      = 1,
            biBitCount    = 32,
            biCompression = 0,             // BI_RGB
        };

        var hdc = NativeMethods.GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero)
        {
            failureMessage = "Failed to acquire a device context.";
            return false;
        }

        try
        {
            var height = (uint)Math.Abs(bm.bmHeight);

            // First call: let GDI fill in biSizeImage.
            NativeMethods.GetDIBits(hdc, hBitmap, 0, height, null, ref header, 0 /* DIB_RGB_COLORS */);

            if (header.biSizeImage == 0)
            {
                var stride = (bm.bmWidth * 32 + 31) / 32 * 4;
                header.biSizeImage = (uint)(stride * height);
            }

            var pixelData = new byte[header.biSizeImage];
            if (NativeMethods.GetDIBits(hdc, hBitmap, 0, height, pixelData, ref header, 0) == 0)
            {
                failureMessage = "Failed to retrieve bitmap pixel data.";
                return false;
            }

            // Assemble a CF_DIB-compatible block: BITMAPINFOHEADER immediately followed by pixels.
            var headerSize = Marshal.SizeOf<BITMAPINFOHEADER>();
            bytes = new byte[headerSize + pixelData.Length];

            var pin = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try   { Marshal.StructureToPtr(header, pin.AddrOfPinnedObject(), fDeleteOld: false); }
            finally { pin.Free(); }

            Buffer.BlockCopy(pixelData, 0, bytes, headerSize, pixelData.Length);
            failureMessage = string.Empty;
            return true;
        }
        finally
        {
            NativeMethods.ReleaseDC(IntPtr.Zero, hdc);
        }
    }

    private static bool TryReadEnhMetaFileAsBytes(IntPtr hEmf, out byte[]? bytes, out string failureMessage)
    {
        bytes = null;

        var size = NativeMethods.GetEnhMetaFileBits(hEmf, 0, null);
        if (size == 0)
        {
            failureMessage = "Failed to determine EMF data size.";
            return false;
        }

        bytes = new byte[size];
        if (NativeMethods.GetEnhMetaFileBits(hEmf, size, bytes) != size)
        {
            bytes = null;
            failureMessage = "Failed to read EMF data.";
            return false;
        }

        failureMessage = string.Empty;
        return true;
    }

    private static bool TryGetClipboardDataSize(uint format, out ulong sizeBytes)
    {
        sizeBytes = 0;

        if (NonGlobalMemoryFormats.Contains(format))
            return false;

        var handle = NativeMethods.GetClipboardData(format);
        if (handle == IntPtr.Zero)
            return false;

        if (HBitmapFormats.Contains(format))
        {
            if (NativeMethods.GetObject(handle, Marshal.SizeOf<BITMAP>(), out var bm) == 0)
                return false;
            var stride = (bm.bmWidth * bm.bmBitsPixel + 31) / 32 * 4;
            sizeBytes = (ulong)(Math.Abs(bm.bmHeight) * stride * bm.bmPlanes);
            return sizeBytes > 0;
        }

        if (HEnhMetaFileFormats.Contains(format))
        {
            var emfSize = NativeMethods.GetEnhMetaFileBits(handle, 0, null);
            if (emfSize == 0)
                return false;
            sizeBytes = emfSize;
            return true;
        }

        var size = NativeMethods.GlobalSize(handle);
        if (size == UIntPtr.Zero)
            return false;

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

    private void UpdateStatusBar()
    {
        StatusText.Text = _isMonitoring ? "Monitoring..." : "Press F5 to refresh";
    }

    private void UpdateFileStatusBar(string action, string path)
    {
        FileStatusText.Text = $"{action}: {Path.GetFileName(path)}";
        FileStatusSeparator.Visibility = Visibility.Visible;
    }

    // Returns the clipboard handle-type tag that matches how Windows stores this format.
    private static string GetHandleType(uint formatId) =>
        HBitmapFormats.Contains(formatId)        ? "hbitmap"       :
        HEnhMetaFileFormats.Contains(formatId)   ? "henhmetafile"  :
        NonGlobalMemoryFormats.Contains(formatId) ? "none"         :
        "hglobal";

    private void RefreshClipboardCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        RefreshFormats();
    }

    private void RefreshClipboardCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        e.CanExecute = !_isMonitoring;
    }

    private void LoadCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        MenuItemLoad_Click(sender, e);
    }

    private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        MenuItemSave_Click(sender, e);
    }

    private void MenuItemExit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ClipboardMenu_SubmenuOpened(object sender, RoutedEventArgs e)
    {
        MenuItemClear.IsEnabled = NativeMethods.CountClipboardFormats() > 0;
    }

    private void MenuItemClear_Click(object sender, RoutedEventArgs e)
    {
        Clipboard.Clear();
        if (!_isMonitoring)
            RefreshFormats();
    }

    private void MenuItemMonitorChanges_Click(object sender, RoutedEventArgs e)
    {
        _isMonitoring = MenuItemMonitorChanges.IsChecked;
        UpdateStatusBar();
        SavePreferences();

        if (_isMonitoring)
        {
            RefreshFormats();
        }
    }

    private void MenuItemSubmitFeedback_Click(object sender, RoutedEventArgs e)
    {
        ShellHelper.OpenUrl("https://github.com/dbakuntsev/simply.clipboard-monitor/issues");
    }

    private void MenuItemAbout_Click(object sender, RoutedEventArgs e)
    {
        new AboutDialog { Owner = this }.ShowDialog();
    }

    private void MenuItemLoad_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title           = "Open Clipboard Database",
            Filter          = "Clipboard Database (*.clipdb)|*.clipdb",
            DefaultExt      = "clipdb",
            CheckFileExists = true,
        };

        if (dlg.ShowDialog(this) != true)
            return;

        try
        {
            LoadFromFile(dlg.FileName);
        }
        catch (Exception ex)
        {
            ShowFileOperationError("load", ex);
        }
    }

    private void MenuItemSave_Click(object sender, RoutedEventArgs e)
    {
        if (_currentFilePath == null)
        {
            MenuItemSaveAs_Click(sender, e);
            return;
        }

        try
        {
            SaveToFile(_currentFilePath);
        }
        catch (Exception ex)
        {
            ShowFileOperationError("save", ex);
        }
    }

    private void MenuItemSaveAs_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            Title           = "Save Clipboard Database",
            Filter          = "Clipboard Database (*.clipdb)|*.clipdb",
            DefaultExt      = "clipdb",
            OverwritePrompt = true,
            FileName        = _currentFilePath != null ? Path.GetFileName(_currentFilePath) : string.Empty,
        };

        if (dlg.ShowDialog(this) != true)
            return;

        try
        {
            SaveToFile(dlg.FileName);
        }
        catch (Exception ex)
        {
            ShowFileOperationError("save", ex);
        }
    }

    // ── Save implementation ─────────────────────────────────────────────────

    private void SaveToFile(string path)
    {
        if (_formats.Count == 0)
        {
            MessageBox.Show(this,
                "There are no clipboard formats to save.\n\nCopy something to the clipboard first, then save.",
                "Nothing to Save", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!TryOpenClipboard(IntPtr.Zero))
        {
            MessageBox.Show(this,
                "Unable to open the clipboard for reading.\n\nClose any application that may be using the clipboard and try again.",
                "Save Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        List<SavedClipboardFormat> formats;
        try
        {
            formats = new List<SavedClipboardFormat>(_formats.Count);
            foreach (var item in _formats)
            {
                var handleType = GetHandleType(item.FormatId);
                byte[]? data = null;
                if (handleType != "none")
                    TryReadClipboardDataBytes(item.FormatId, out data, out _);

                formats.Add(new SavedClipboardFormat(item.Ordinal, item.FormatId, item.Name, handleType, data));
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        ClipboardDatabase.Save(path, formats);
        _currentFilePath = path;
        UpdateFileStatusBar("Saved", path);
    }

    // ── Load implementation ─────────────────────────────────────────────────

    private void LoadFromFile(string path)
    {
        var formats = ClipboardDatabase.Load(path);

        if (formats.Count == 0)
        {
            MessageBox.Show(this,
                "The clipboard database is empty — no formats were stored in the file.",
                "Nothing to Load", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!TryOpenClipboard(_hwndSource?.Handle ?? IntPtr.Zero))
        {
            MessageBox.Show(this,
                "Unable to open the clipboard for writing.\n\nClose any application that may be using the clipboard and try again.",
                "Load Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            NativeMethods.EmptyClipboard();
            foreach (var fmt in formats)
                RestoreFormat(fmt);
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        _currentFilePath = path;
        UpdateFileStatusBar("Loaded", path);

        // When monitoring is active, WM_CLIPBOARDUPDATE fires automatically
        // and refreshes the list; refresh manually only when monitoring is off.
        if (!_isMonitoring)
            RefreshFormats();
    }

    private static void RestoreFormat(SavedClipboardFormat fmt)
    {
        // Custom format IDs (≥ 0xC000) are assigned dynamically per Windows session;
        // re-register by name to get the current-session ID.
        uint actualId = fmt.FormatId >= 0xC000
            ? NativeMethods.RegisterClipboardFormat(fmt.FormatName)
            : fmt.FormatId;

        if (actualId == 0)
            return;

        switch (fmt.HandleType)
        {
            case "hglobal":
                RestoreAsHGlobal(actualId, fmt.Data);
                break;
            case "hbitmap":
                RestoreAsHBitmap(actualId, fmt.Data);
                break;
            case "henhmetafile":
                RestoreAsHEnhMetaFile(actualId, fmt.Data);
                break;
            // "none" (e.g. CF_PALETTE): no bytes were captured; skip silently.
        }
    }

    private static void RestoreAsHGlobal(uint formatId, byte[]? data)
    {
        if (data is not { Length: > 0 })
            return;

        const uint GMEM_MOVEABLE = 0x0002;
        var hGlobal = NativeMethods.GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)(uint)data.Length);
        if (hGlobal == IntPtr.Zero)
            return;

        var ptr = NativeMethods.GlobalLock(hGlobal);
        if (ptr == IntPtr.Zero)
        {
            NativeMethods.GlobalFree(hGlobal);
            return;
        }

        try   { Marshal.Copy(data, 0, ptr, data.Length); }
        finally { NativeMethods.GlobalUnlock(hGlobal); }

        if (NativeMethods.SetClipboardData(formatId, hGlobal) == IntPtr.Zero)
            NativeMethods.GlobalFree(hGlobal);
    }

    private static void RestoreAsHEnhMetaFile(uint formatId, byte[]? data)
    {
        if (data is not { Length: > 0 })
            return;

        var hemf = NativeMethods.SetEnhMetaFileBits((uint)data.Length, data);
        if (hemf == IntPtr.Zero)
            return;

        if (NativeMethods.SetClipboardData(formatId, hemf) == IntPtr.Zero)
            NativeMethods.DeleteEnhMetaFile(hemf);
    }

    private static void RestoreAsHBitmap(uint formatId, byte[]? data)
    {
        var headerSize = Marshal.SizeOf<BITMAPINFOHEADER>();
        if (data == null || data.Length <= headerSize)
            return;

        // The stored block is BITMAPINFOHEADER + pixel data (produced by GetDIBits, 32 bpp BI_RGB).
        var header    = MemoryMarshal.Read<BITMAPINFOHEADER>(data.AsSpan(0, headerSize));
        var pixelData = data.AsSpan(headerSize).ToArray();

        var hdc = NativeMethods.GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero)
            return;

        try
        {
            const uint CBM_INIT        = 4;
            const uint DIB_RGB_COLORS  = 0;
            var hBitmap = NativeMethods.CreateDIBitmap(hdc, ref header, CBM_INIT, pixelData, ref header, DIB_RGB_COLORS);
            if (hBitmap == IntPtr.Zero)
                return;

            if (NativeMethods.SetClipboardData(formatId, hBitmap) == IntPtr.Zero)
                NativeMethods.DeleteObject(hBitmap);
        }
        finally
        {
            NativeMethods.ReleaseDC(IntPtr.Zero, hdc);
        }
    }

    // ── Error reporting ─────────────────────────────────────────────────────

    private void ShowFileOperationError(string operation, Exception ex)
    {
        var (message, remedy) = ex switch
        {
            SqliteException { Message: var m } when m.Contains("not a database", StringComparison.OrdinalIgnoreCase)
                => ($"The file is not a valid clipboard database.",
                    "Select a .clipdb file created by this application."),

            SqliteException { Message: var m }
                => ($"Database error: {m}", null),

            UnauthorizedAccessException
                => ("Access to the file was denied.", "Check that you have permission to read/write the selected file."),

            IOException { Message: var m }
                => ($"File I/O error: {m}", "Check that the file is not in use by another application."),

            OutOfMemoryException
                => ("The clipboard database is too large to load.", "Try loading a smaller file."),

            _ => (ex.Message, null),
        };

        var text = remedy != null
            ? $"Failed to {operation} clipboard database:\n\n{message}\n\n{remedy}"
            : $"Failed to {operation} clipboard database:\n\n{message}";

        MessageBox.Show(this, text, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private void FormatListBox_OnColumnHeaderClick(object sender, RoutedEventArgs e)
    {
        if (e.OriginalSource is not GridViewColumnHeader header || header.Role == GridViewColumnHeaderRole.Padding)
        {
            return;
        }

        if (header.Tag is not string requestedProperty || !IsSupportedSortProperty(requestedProperty))
        {
            return;
        }

        if (string.Equals(_currentSortProperty, requestedProperty, StringComparison.Ordinal))
        {
            _currentSortDirection = _currentSortDirection == ListSortDirection.Ascending
                ? ListSortDirection.Descending
                : ListSortDirection.Ascending;
        }
        else
        {
            _currentSortProperty = requestedProperty;
            _currentSortDirection = ListSortDirection.Ascending;
        }

        ApplyFormatSort();
        SavePreferences();
    }

    private void ApplyFormatSort()
    {
        var view = CollectionViewSource.GetDefaultView(FormatListBox.ItemsSource);
        if (view == null)
        {
            return;
        }

        using (view.DeferRefresh())
        {
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(_currentSortProperty, _currentSortDirection));
        }

        UpdateSortIndicators();
    }

    private static bool IsSupportedSortProperty(string propertyName)
    {
        return propertyName is nameof(ClipboardFormatItem.Ordinal)
            or nameof(ClipboardFormatItem.FormatId)
            or nameof(ClipboardFormatItem.Name)
            or nameof(ClipboardFormatItem.ContentSizeValue);
    }

    private static bool IsSupportedColumnKey(string key)
    {
        return IsSupportedSortProperty(key);
    }

    private void LoadPreferences()
    {
        _storedColumnPreferences = [];

        try
        {
            var path = GetPreferencesFilePath();
            if (!File.Exists(path))
            {
                return;
            }

            var json = File.ReadAllText(path);
            var preferences = JsonSerializer.Deserialize<UserPreferences>(json);
            if (preferences == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(preferences.SortProperty) && IsSupportedSortProperty(preferences.SortProperty))
            {
                _currentSortProperty = preferences.SortProperty;
            }

            if (Enum.TryParse<ListSortDirection>(preferences.SortDirection, ignoreCase: true, out var direction))
            {
                _currentSortDirection = direction;
            }

            if (preferences.FormatColumns != null)
            {
                foreach (var column in preferences.FormatColumns)
                {
                    if (column == null || string.IsNullOrWhiteSpace(column.Key) || !IsSupportedColumnKey(column.Key))
                    {
                        continue;
                    }

                    _storedColumnPreferences.Add(new FormatColumnPreference
                    {
                        Key = column.Key,
                        Width = column.Width
                    });
                }
            }

            _isMonitoring = preferences.MonitorChanges;
        }
        catch
        {
            _currentSortProperty = DefaultSortProperty;
            _currentSortDirection = ListSortDirection.Ascending;
            _storedColumnPreferences = [];
            _isMonitoring = true;
        }
    }

    private void SavePreferences()
    {
        try
        {
            var path = GetPreferencesFilePath();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var preferences = new UserPreferences
            {
                SortProperty = _currentSortProperty,
                SortDirection = _currentSortDirection.ToString(),
                FormatColumns = CaptureFormatColumnPreferences(),
                MonitorChanges = _isMonitoring
            };

            var json = JsonSerializer.Serialize(preferences, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch
        {
            // Ignore preference persistence failures.
        }
    }

    private void ApplyFormatColumnPreferences()
    {
        var gridView = GetFormatGridView();
        if (gridView == null || _storedColumnPreferences.Count == 0)
        {
            return;
        }

        var columnsByKey = new Dictionary<string, GridViewColumn>(StringComparer.Ordinal);
        foreach (var column in gridView.Columns)
        {
            if (TryGetColumnKey(column, out var key) && !columnsByKey.ContainsKey(key))
            {
                columnsByKey[key] = column;
            }
        }

        var orderedColumns = new List<GridViewColumn>();
        foreach (var preference in _storedColumnPreferences)
        {
            if (!columnsByKey.TryGetValue(preference.Key, out var column))
            {
                continue;
            }

            if (preference.Width.HasValue && preference.Width.Value > 0)
            {
                column.Width = preference.Width.Value;
            }

            orderedColumns.Add(column);
            columnsByKey.Remove(preference.Key);
        }

        foreach (var column in gridView.Columns)
        {
            if (!TryGetColumnKey(column, out var key))
            {
                continue;
            }

            if (columnsByKey.Remove(key))
            {
                orderedColumns.Add(column);
            }
        }

        if (orderedColumns.Count != gridView.Columns.Count)
        {
            return;
        }

        gridView.Columns.Clear();
        foreach (var column in orderedColumns)
        {
            gridView.Columns.Add(column);
        }
    }

    private List<FormatColumnPreference> CaptureFormatColumnPreferences()
    {
        var result = new List<FormatColumnPreference>();
        var gridView = GetFormatGridView();
        if (gridView == null)
        {
            return result;
        }

        foreach (var column in gridView.Columns)
        {
            if (!TryGetColumnKey(column, out var key) || !IsSupportedColumnKey(key))
            {
                continue;
            }

            result.Add(new FormatColumnPreference
            {
                Key = key,
                Width = column.Width > 0 ? column.Width : null
            });
        }

        return result;
    }

    private GridView? GetFormatGridView()
    {
        return FormatListBox.View as GridView;
    }

    private static bool TryGetColumnKey(GridViewColumn column, out string key)
    {
        key = string.Empty;

        if (column.Header is not GridViewColumnHeader header || header.Tag is not string tag || string.IsNullOrWhiteSpace(tag))
        {
            return false;
        }

        key = tag;
        return true;
    }

    private void UpdateSortIndicators()
    {
        OrdinalColumnHeader.Content = "#";
        IdColumnHeader.Content = "ID";
        FormatColumnHeader.Content = "Format";
        SizeColumnHeader.Content = "Size";

        var currentHeader = GetHeaderBySortProperty(_currentSortProperty);
        if (currentHeader == null)
        {
            return;
        }

        var direction = _currentSortDirection == ListSortDirection.Ascending ? "↑" : "↓";
        currentHeader.Content = $"{GetHeaderBaseTitle(_currentSortProperty)} {direction}";
    }

    private GridViewColumnHeader? GetHeaderBySortProperty(string propertyName)
    {
        return propertyName switch
        {
            nameof(ClipboardFormatItem.Ordinal) => OrdinalColumnHeader,
            nameof(ClipboardFormatItem.FormatId) => IdColumnHeader,
            nameof(ClipboardFormatItem.Name) => FormatColumnHeader,
            nameof(ClipboardFormatItem.ContentSizeValue) => SizeColumnHeader,
            _ => null
        };
    }

    private static string GetHeaderBaseTitle(string propertyName)
    {
        return propertyName switch
        {
            nameof(ClipboardFormatItem.Ordinal) => "#",
            nameof(ClipboardFormatItem.FormatId) => "ID",
            nameof(ClipboardFormatItem.Name) => "Format",
            nameof(ClipboardFormatItem.ContentSizeValue) => "Size",
            _ => propertyName
        };
    }

    private static string GetPreferencesFilePath()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(appData, "Simply.ClipboardMonitor", PreferencesFileName);
    }

}




