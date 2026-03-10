using System.Collections.ObjectModel;
using System.Collections;
using System.ComponentModel;
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

namespace Simply.ClipboardMonitor;

public partial class MainWindow : Window
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    private const int ERROR_NOT_LOCKED = 158;
    private const int BYTES_PER_ROW = 16;
    private const string PreferencesFileName = "preferences.json";
    private const string DefaultSortProperty = nameof(ClipboardFormatItem.Ordinal);

    public static readonly RoutedUICommand RefreshClipboardCommand = new(
        "Refresh", "RefreshClipboard", typeof(MainWindow),
        new InputGestureCollection { new KeyGesture(Key.F5) });

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
    private string _currentSortProperty = DefaultSortProperty;
    private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
    private List<FormatColumnPreference> _storedColumnPreferences = [];
    private bool _isMonitoring;

    public MainWindow()
    {
        InitializeComponent();
        CommandBindings.Add(new CommandBinding(RefreshClipboardCommand, RefreshClipboardCommand_Executed, RefreshClipboardCommand_CanExecute));
        _isUiReady = true;
        FormatListBox.ItemsSource = _formats;
        LoadPreferences();
        MenuItemMonitorChanges.IsChecked = _isMonitoring;
        UpdateStatusBar();
        ApplyFormatColumnPreferences();
        FormatListBox.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler(FormatListBox_OnColumnHeaderClick));
        ApplyFormatSort();
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

                results.Add(new ClipboardFormatItem(ordinal, (int)format, format, name, contentSize, contentSizeValue));
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

    private void UpdateStatusBar()
    {
        StatusText.Text = _isMonitoring ? "Monitoring..." : "Ready";
    }

    private void RefreshClipboardCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        RefreshFormats();
    }

    private void RefreshClipboardCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        e.CanExecute = !_isMonitoring;
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

    private sealed class ClipboardFormatItem(int ordinal, int id, uint formatNumberValue, string name, string contentSize, long contentSizeValue)
    {
        public int Ordinal { get; } = ordinal;
        public int Id { get; } = id;
        public uint FormatNumberValue { get; } = formatNumberValue;
        public string FormatNumber => FormatNumberValue.ToString("D");
        public string Name { get; } = name;
        public string ContentSize { get; } = contentSize;
        public long ContentSizeValue { get; } = contentSizeValue;
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
            or nameof(ClipboardFormatItem.FormatNumberValue)
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
            nameof(ClipboardFormatItem.FormatNumberValue) => IdColumnHeader,
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
            nameof(ClipboardFormatItem.FormatNumberValue) => "ID",
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

    private sealed class UserPreferences
    {
        public string? SortProperty { get; set; }
        public string? SortDirection { get; set; }
        public List<FormatColumnPreference>? FormatColumns { get; set; }
        public bool MonitorChanges { get; set; } = true;
    }

    private sealed class FormatColumnPreference
    {
        public string Key { get; set; } = string.Empty;
        public double? Width { get; set; }
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

        [DllImport("user32.dll")]
        internal static extern int CountClipboardFormats();

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




