using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System.Threading.Channels;
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
using System.Windows.Media;
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

    public static readonly RoutedUICommand ExportCommand = new(
        "Export Selected Format", "ExportSelectedFormat", typeof(MainWindow),
        new InputGestureCollection { new KeyGesture(Key.E, ModifierKeys.Control) });

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

    // Text-preview encoding state
    private byte[]?                      _currentTextBytes;
    private bool                         _suppressEncodingChange;
    private IReadOnlyList<EncodingItem>? _encodingItems;
    private Encoding?                    _currentAutoDetectedEncoding;
    private bool                         _textEncodingManuallyChanged;

    // Export state — raw bytes and identity of the currently selected format
    private byte[]? _currentFormatBytes;
    private uint    _currentFormatId;
    private string  _currentFormatName = string.Empty;

    // History state
    private bool _isTrackingHistory;
    private readonly ObservableCollection<HistoryItem> _historyItems = [];
    // Non-null when showing a history session (maps format name → snapshot).
    // Null means the live clipboard or the static snapshot is used instead.
    private IReadOnlyDictionary<string, FormatSnapshot>? _historySnapshots;
    // Non-null when monitoring is off: holds a byte-level copy of every format
    // taken at the moment monitoring was disabled (or on a manual refresh).
    // Used instead of the live clipboard so format previews still work.
    private Dictionary<string, FormatSnapshot>? _staticSnapshot;
    // Single-reader channel that serialises all DB writes through one background consumer.
    private readonly Channel<(IReadOnlyList<FormatSnapshot> Snapshots, DateTime Timestamp)> _historyChannel =
        Channel.CreateUnbounded<(IReadOnlyList<FormatSnapshot>, DateTime)>(
            new UnboundedChannelOptions { SingleReader = true });

    private sealed class HistoryItem
    {
        public required long     SessionId   { get; init; }
        public required DateTime Timestamp   { get; init; }
        public required string   FormatsText { get; init; }
        public required long     TotalSize   { get; init; }
        public string DateText => Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
        public string SizeText => TotalSize switch
        {
            >= 1024L * 1024 => $"{TotalSize / (1024.0 * 1024):F1} MB",
            >= 1024         => $"{TotalSize / 1024.0:F1} KB",
            _               => $"{TotalSize} B",
        };
    }

    public MainWindow()
    {
        // Required for legacy OEM/ANSI code pages (e.g. CP437 used by CF_OEMTEXT) which
        // are not included in the default .NET encoding set.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        InitializeComponent();
        CommandBindings.Add(new CommandBinding(RefreshClipboardCommand, RefreshClipboardCommand_Executed, RefreshClipboardCommand_CanExecute));
        CommandBindings.Add(new CommandBinding(LoadCommand, LoadCommand_Executed));
        CommandBindings.Add(new CommandBinding(SaveCommand, SaveCommand_Executed));
        CommandBindings.Add(new CommandBinding(ExportCommand, ExportCommand_Executed, ExportCommand_CanExecute));
        _ = Task.Run(ProcessHistoryChannelAsync);
        _isUiReady = true;
        FormatListBox.ItemsSource  = _formats;
        HistoryListView.ItemsSource = _historyItems;
        LoadPreferences();
        MenuItemMonitorChanges.IsChecked = _isMonitoring;
        MenuItemTrackHistory.IsChecked   = _isTrackingHistory;
        if (_isTrackingHistory)
            ShowHistoryPanel();
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

        if (!_isMonitoring)
            CaptureStaticSnapshot();

        if (_isTrackingHistory)
            LoadHistoryFromDatabase();
    }

    protected override void OnClosed(EventArgs e)
    {
        if (_hwndSource != null)
        {
            NativeMethods.RemoveClipboardFormatListener(_hwndSource.Handle);
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }

        _historyChannel.Writer.TryComplete();
        SavePreferences();
        base.OnClosed(e);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_CLIPBOARDUPDATE)
        {
            if (_isMonitoring)
            {
                // Read the sequence number before RefreshFormats() so that any
                // GetClipboardData call (including delayed-rendering triggers) that
                // increments it during the refresh is detectable afterwards.
                var seqAtArrival = NativeMethods.GetClipboardSequenceNumber();
                RefreshFormats();
                if (_isTrackingHistory)
                    CaptureHistorySession(seqAtArrival);
            }
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void RefreshFormats()
    {
        _historySnapshots = null;   // switch to live-clipboard mode
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

        _currentFormatId    = item.FormatId;
        _currentFormatName  = item.Name;
        _currentFormatBytes = null;

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

        if (_historySnapshots != null)
        {
            // History playback mode — bytes come from the in-memory snapshot dict.
            if (_historySnapshots.TryGetValue(item.Name, out var snapshot))
            {
                bytes = snapshot.Data;
                if (bytes == null && snapshot.HandleType != "none")
                    hexFailureMessage = "No data was captured for this format.";
            }
            else
            {
                hexFailureMessage = "Format data not found in the selected history entry.";
            }
        }
        else if (_staticSnapshot != null)
        {
            // Monitoring-off mode — bytes come from the snapshot taken when monitoring
            // was disabled (or on the last manual refresh), so the clipboard does not
            // need to be opened and previews stay consistent with the displayed formats.
            if (_staticSnapshot.TryGetValue(item.Name, out var snapshot))
            {
                bytes = snapshot.Data;
                if (bytes == null && snapshot.HandleType != "none")
                    hexFailureMessage = "No data was captured for this format.";
            }
            else
            {
                hexFailureMessage = "Format data not found in the static snapshot.";
            }
        }
        else
        {
            // Live mode — read from the clipboard.
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
        }

        _currentFormatBytes = bytes;

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
        _currentTextBytes            = bytes;
        _currentAutoDetectedEncoding = null;
        _textEncodingManuallyChanged = false;

        if (!TryDecodeTextContent(formatId, formatName, bytes, out var text, out var message, out var detectedEncoding))
        {
            SetTextPreviewUnavailable(message);
            return;
        }

        _currentAutoDetectedEncoding    = detectedEncoding;
        TextContentTextBox.Foreground   = SystemColors.WindowTextBrush;
        TextContentTextBox.Text         = text;
        TextStatusTextBlock.Text        = DecodedTextStatusText(text);
        PopulateEncodingComboBox(detectedEncoding);
    }

    private void PopulateEncodingComboBox(Encoding? preselect)
    {
        _encodingItems ??= Encoding.GetEncodings()
            .Select(info =>
            {
                try   { return new EncodingItem(info.GetEncoding(), $"{info.DisplayName} ({info.Name})"); }
                catch { return null; }
            })
            .OfType<EncodingItem>()
            .Append(new EncodingItem(
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                "Unicode (UTF-8, strict) (utf-8-strict)"))
            .OrderBy(e => e.DisplayText)
            .ToList();

        if (TextEncodingComboBox.ItemsSource == null)
            TextEncodingComboBox.ItemsSource = _encodingItems;

        TextEncodingComboBox.IsEnabled = true;

        _suppressEncodingChange         = true;
        TextEncodingComboBox.SelectedItem = preselect != null
            ? _encodingItems.FirstOrDefault(e => e.Encoding.CodePage == preselect.CodePage)
            : null;
        _suppressEncodingChange = false;
    }

    private void TextEncodingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressEncodingChange)
            return;
        if (TextEncodingComboBox.SelectedItem is not EncodingItem item || _currentTextBytes == null)
            return;

        _textEncodingManuallyChanged = true;

        try
        {
            var decoded                   = item.Encoding.GetString(_currentTextBytes).TrimEnd('\0');
            TextContentTextBox.Text       = decoded;
            TextContentTextBox.Foreground = SystemColors.WindowTextBrush;
            TextStatusTextBlock.Text      = DecodedTextStatusText(decoded);
        }
        catch (Exception ex)
        {
            TextContentTextBox.Text       = $"Cannot decode as {item.DisplayText}:\n{ex.Message}";
            TextContentTextBox.Foreground = Brushes.Crimson;
            TextStatusTextBlock.Text      = "Decoding failed.";
        }
    }

    private static string DecodedTextStatusText(string text)
    {
        int chars  = text.Length;
        int nonWs  = text.Count(c => !char.IsWhiteSpace(c));
        int newlines = 0;
        for (int i = 0; i < text.Length; i++)
        {
            // Approximate the number of lines by counting standalone \n and \r as well as sequences \r\n as a single newline.
            if (text[i] == '\r') 
            { 
                newlines++; 
                if (i + 1 < text.Length && text[i + 1] == '\n') 
                    i++; 
            }
            else if (text[i] == '\n') 
                newlines++;
        }
        int lines  = chars == 0 ? 0 : newlines + 1;
        return $"{chars:N0} character{(chars == 1 ? "" : "s")} ({nonWs:N0} non-whitespace) · {lines:N0} line{(lines == 1 ? "" : "s")}";
    }

    private static bool TryDecodeTextContent(uint formatId, string formatName, byte[] bytes,
        out string text, out string failureMessage, out Encoding? detectedEncoding)
    {
        if (!IsTextCompatible(formatId, formatName))
        {
            text             = string.Empty;
            failureMessage   = "Text preview unavailable for this format.";
            detectedEncoding = null;
            return false;
        }

        try
        {
            if (formatId == 13) // CF_UNICODETEXT
            {
                text             = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
                detectedEncoding = Encoding.Unicode;
            }
            else if (formatId == 1) // CF_TEXT
            {
                text             = Encoding.Default.GetString(TrimAtNull(bytes, 1));
                detectedEncoding = Encoding.Default;
            }
            else if (formatId == 7) // CF_OEMTEXT
            {
                var oemEncoding  = Encoding.GetEncoding((int)NativeMethods.GetOEMCP());
                text             = oemEncoding.GetString(TrimAtNull(bytes, 1));
                detectedEncoding = oemEncoding;
            }
            else
            {
                text = DecodeWithFallback(bytes, out detectedEncoding);
            }

            failureMessage = string.Empty;
            return true;
        }
        catch
        {
            text             = string.Empty;
            failureMessage   = "Failed to decode this format as text.";
            detectedEncoding = null;
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

    private enum Utf16Variant { None, LittleEndian, BigEndian }

    private sealed class EncodingItem
    {
        public Encoding Encoding    { get; }
        public string   DisplayText { get; }

        public EncodingItem(Encoding encoding, string displayText)
        {
            Encoding    = encoding;
            DisplayText = displayText;
        }

        public override string ToString() => DisplayText;
    }

    /// <summary>
    /// Detects UTF-16 encoding via BOM or heuristic null-byte pattern.
    /// For LE, ASCII-range characters place 0x00 at every odd byte position.
    /// For BE, they place 0x00 at every even byte position.
    /// </summary>
    private static Utf16Variant DetectUtf16(byte[] bytes)
    {
        if (bytes.Length < 4 || bytes.Length % 2 != 0)
            return Utf16Variant.None;

        // Explicit BOM takes priority
        if (bytes[0] == 0xFF && bytes[1] == 0xFE) return Utf16Variant.LittleEndian;
        if (bytes[0] == 0xFE && bytes[1] == 0xFF) return Utf16Variant.BigEndian;

        // Heuristic: sample up to the first 512 bytes
        var sampleLen = Math.Min(bytes.Length, 512);
        var pairs = sampleLen / 2;
        int nullAtOdd = 0, nullAtEven = 0;
        for (var i = 0; i < sampleLen - 1; i += 2)
        {
            if (bytes[i]     == 0) nullAtEven++;
            if (bytes[i + 1] == 0) nullAtOdd++;
        }

        // Require ≥80 % of high bytes to be null AND <5 % of low bytes to be null
        const double highThreshold = 0.80;
        const double lowMaxRatio   = 0.05;
        if ((double)nullAtOdd  / pairs >= highThreshold && (double)nullAtEven / pairs < lowMaxRatio)
            return Utf16Variant.LittleEndian;
        if ((double)nullAtEven / pairs >= highThreshold && (double)nullAtOdd  / pairs < lowMaxRatio)
            return Utf16Variant.BigEndian;

        return Utf16Variant.None;
    }

    private static string DecodeWithFallback(byte[] bytes, out Encoding usedEncoding)
    {
        // UTF-8 BOM
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            usedEncoding = Encoding.UTF8;
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3).TrimEnd('\0');
        }

        // UTF-16 LE / BE — BOM or heuristic
        var utf16 = DetectUtf16(bytes);
        if (utf16 == Utf16Variant.LittleEndian)
        {
            usedEncoding = Encoding.Unicode;
            var start    = (bytes[0] == 0xFF && bytes[1] == 0xFE) ? 2 : 0;
            var trimmed  = TrimAtNull(bytes[start..], 2);
            return Encoding.Unicode.GetString(trimmed).TrimEnd('\0');
        }
        if (utf16 == Utf16Variant.BigEndian)
        {
            usedEncoding = Encoding.BigEndianUnicode;
            var start    = (bytes[0] == 0xFE && bytes[1] == 0xFF) ? 2 : 0;
            var trimmed  = TrimAtNull(bytes[start..], 2);
            return Encoding.BigEndianUnicode.GetString(trimmed).TrimEnd('\0');
        }

        // Strict UTF-8, then system ANSI
        try
        {
            usedEncoding = Encoding.UTF8;
            return new UTF8Encoding(false, true).GetString(bytes).TrimEnd('\0');
        }
        catch
        {
            usedEncoding = Encoding.Default;
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
        _currentTextBytes               = null;
        _currentAutoDetectedEncoding    = null;
        _textEncodingManuallyChanged    = false;
        TextStatusTextBlock.Text        = message;
        TextContentTextBox.Foreground   = SystemColors.WindowTextBrush;
        TextContentTextBox.Text         = string.Empty;
        TextEncodingComboBox.IsEnabled  = false;
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
        _currentFormatBytes = null;
        _currentFormatId    = 0;
        _currentFormatName  = string.Empty;
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
        CaptureStaticSnapshot();
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

    private void ExportCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        e.CanExecute = _currentFormatBytes is { Length: > 0 };
    }

    private void ExportCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        MenuItemExportSelectedFormat_Click(sender, e);
    }

    private void MenuItemExit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ClipboardMenu_SubmenuOpened(object sender, RoutedEventArgs e)
    {
        MenuItemClear.IsEnabled        = NativeMethods.CountClipboardFormats() > 0;
        MenuItemTrackHistory.IsEnabled = _isMonitoring;
    }

    private void MenuItemClear_Click(object sender, RoutedEventArgs e)
    {
        Clipboard.Clear();
        if (!_isMonitoring)
        {
            RefreshFormats();
            CaptureStaticSnapshot();
        }
    }

    private void MenuItemMonitorChanges_Click(object sender, RoutedEventArgs e)
    {
        _isMonitoring = MenuItemMonitorChanges.IsChecked;
        UpdateStatusBar();

        // History tracking requires monitoring — stop it if monitoring is turned off.
        if (!_isMonitoring && _isTrackingHistory)
        {
            _isTrackingHistory               = false;
            MenuItemTrackHistory.IsChecked   = false;
            HideHistoryPanel();
            _historyItems.Clear();
            _historySnapshots = null;
        }

        SavePreferences();

        if (_isMonitoring)
        {
            _staticSnapshot = null;
            RefreshFormats();
        }
        else
        {
            CaptureStaticSnapshot();
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

    // ── Export implementation ────────────────────────────────────────────────

    private void MenuItemExportSelectedFormat_Click(object sender, RoutedEventArgs e)
    {
        if (_currentFormatBytes is not { Length: > 0 })
            return;

        bool textAvailable  = _currentAutoDetectedEncoding != null;
        bool imageAvailable = ImagePreview.Source is BitmapSource;

        // Build filter list in display order
        var filters = new List<(string FilterPart, string Ext)>();
        if (textAvailable)
            filters.Add(("Text (*.txt)|*.txt", ".txt"));
        if (imageAvailable)
        {
            filters.Add(("PNG Image (*.png)|*.png", ".png"));
            filters.Add(("JPEG Image (*.jpg)|*.jpg", ".jpg"));
        }
        filters.Add(("Binary Data (*.bin)|*.bin", ".bin"));

        // Determine the 1-based default filter index
        int defaultFilter;
        if (textAvailable)
        {
            defaultFilter = filters.FindIndex(f => f.Ext == ".txt") + 1;
        }
        else if (imageAvailable)
        {
            var norm = _currentFormatName.ToLowerInvariant();
            defaultFilter = (norm.Contains("jpeg") || norm.Contains("jpg"))
                ? filters.FindIndex(f => f.Ext == ".jpg") + 1
                : filters.FindIndex(f => f.Ext == ".png") + 1;
        }
        else
        {
            defaultFilter = filters.FindIndex(f => f.Ext == ".bin") + 1;
        }

        var selectedClipboardItem = FormatListBox.SelectedItem as ClipboardFormatItem;
        string fileName = $"clipboard-{new string((selectedClipboardItem?.Name ?? "CF_UNKNOWN").Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch).ToArray())}-{DateTime.Now:yyyyMMddTHHmmss}";

        var dlg = new SaveFileDialog
        {
            Title           = $"Export Selected Clipboard Format: {selectedClipboardItem?.Name ?? "CF_UNKNOWN"}",
            Filter          = string.Join("|", filters.Select(f => f.FilterPart)),
            FilterIndex     = defaultFilter,
            DefaultExt      = filters[defaultFilter - 1].Ext.TrimStart('.'),
            FileName        = fileName,
            OverwritePrompt = true,
        };

        if (dlg.ShowDialog(this) != true)
            return;

        var selectedExt = filters[dlg.FilterIndex - 1].Ext;
        try
        {
            switch (selectedExt)
            {
                case ".txt": ExportAsText(dlg.FileName);  break;
                case ".png": ExportAsPng(dlg.FileName);   break;
                case ".jpg": ExportAsJpeg(dlg.FileName);  break;
                case ".bin": File.WriteAllBytes(dlg.FileName, _currentFormatBytes); break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Failed to export:\n\n{ex.Message}",
                "Export Failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ExportAsText(string path)
    {
        Encoding encoding;
        if (ContentTabControl.SelectedItem == TextTabItem
            && _textEncodingManuallyChanged
            && TextEncodingComboBox.SelectedItem is EncodingItem { Encoding: var manualEncoding })
        {
            encoding = manualEncoding;
        }
        else
        {
            encoding = _currentAutoDetectedEncoding!;
        }

        var text = encoding.GetString(_currentTextBytes!).TrimEnd('\0');
        File.WriteAllText(path, text, encoding);
    }

    private void ExportAsPng(string path)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create((BitmapSource)ImagePreview.Source!));
        using var fs = File.Create(path);
        encoder.Save(fs);
    }

    private void ExportAsJpeg(string path)
    {
        var norm = _currentFormatName.ToLowerInvariant();
        if (norm.Contains("jpeg") || norm.Contains("jpg"))
        {
            // Source is already JPEG — write raw bytes to preserve original quality.
            File.WriteAllBytes(path, _currentFormatBytes!);
        }
        else
        {
            var encoder = new JpegBitmapEncoder { QualityLevel = 80 };
            encoder.Frames.Add(BitmapFrame.Create((BitmapSource)ImagePreview.Source!));
            using var fs = File.Create(path);
            encoder.Save(fs);
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

            _isMonitoring      = preferences.MonitorChanges;
            _isTrackingHistory = preferences.TrackHistory && _isMonitoring;
        }
        catch
        {
            _currentSortProperty     = DefaultSortProperty;
            _currentSortDirection    = ListSortDirection.Ascending;
            _storedColumnPreferences = [];
            _isMonitoring            = true;
            _isTrackingHistory       = false;
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
                SortProperty   = _currentSortProperty,
                SortDirection  = _currentSortDirection.ToString(),
                FormatColumns  = CaptureFormatColumnPreferences(),
                MonitorChanges = _isMonitoring,
                TrackHistory   = _isTrackingHistory,
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

    // ── Clipboard History ────────────────────────────────────────────────────

    private void MenuItemTrackHistory_Click(object sender, RoutedEventArgs e)
    {
        _isTrackingHistory = MenuItemTrackHistory.IsChecked;

        if (_isTrackingHistory)
        {
            ShowHistoryPanel();
            LoadHistoryFromDatabase();
        }
        else
        {
            HideHistoryPanel();
            _historyItems.Clear();

            if (_historySnapshots != null)
            {
                _historySnapshots = null;
                RefreshFormats();
            }
        }

        SavePreferences();
    }

    private void HistoryListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not HistoryItem entry)
            return;

        List<FormatSnapshot> snapshots;
        try
        {
            snapshots = ClipboardHistory.LoadSessionFormats(entry.SessionId);
        }
        catch
        {
            return;
        }

        _historySnapshots = snapshots.ToDictionary(s => s.FormatName, StringComparer.Ordinal);
        InitializePreviewState();
        _formats.Clear();

        foreach (var snap in snapshots)
        {
            var contentSize      = snap.OriginalSize > 0 ? snap.OriginalSize.ToString("N0") : "n/a";
            var contentSizeValue = snap.OriginalSize > 0 ? snap.OriginalSize : -1L;
            _formats.Add(new ClipboardFormatItem(snap.Ordinal, snap.FormatId, snap.FormatName, contentSize, contentSizeValue));
        }
    }

    private void CaptureHistorySession(uint seqAtArrival)
    {
        // _formats was just populated by RefreshFormats() — snapshot it on the UI thread
        // before opening the clipboard so we do not need to enumerate formats again.
        if (_formats.Count == 0)
            return;

        var timestamp   = DateTime.Now;
        var formatItems = _formats.ToList();  // snapshot on UI thread

        if (!TryOpenClipboard(IntPtr.Zero))
            return;

        var snapshots = new List<FormatSnapshot>(formatItems.Count);
        try
        {
            foreach (var item in formatItems)
            {
                var handleType = GetHandleType(item.FormatId);
                byte[]? data   = null;
                if (handleType != "none")
                    TryReadClipboardDataBytes(item.FormatId, out data, out _);

                // Prefer the actual byte length; fall back to the size already measured
                // by RefreshFormats for formats whose data could not be read as a byte[].
                var originalSize = data?.LongLength
                    ?? (item.ContentSizeValue >= 0 ? item.ContentSizeValue : 0L);

                snapshots.Add(new FormatSnapshot(item.Ordinal, item.FormatId, item.Name,
                                                 handleType, data, originalSize));
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        // If the sequence number advanced since WM_CLIPBOARDUPDATE arrived, a
        // delayed-rendering provider rendered data and posted a fresh
        // WM_CLIPBOARDUPDATE.  Skip this capture; the next message will capture
        // the fully-rendered clipboard and avoids creating a duplicate entry.
        if (NativeMethods.GetClipboardSequenceNumber() != seqAtArrival)
            return;

        _historyChannel.Writer.TryWrite((snapshots, timestamp));
    }

    /// <summary>
    /// Reads the raw bytes for every format currently in <see cref="_formats"/> and
    /// stores them in <see cref="_staticSnapshot"/> so that format previews remain
    /// functional while monitoring is off (called on startup, on manual refresh, and
    /// whenever monitoring is disabled).
    /// </summary>
    private void CaptureStaticSnapshot()
    {
        if (_formats.Count == 0)
        {
            _staticSnapshot = null;
            return;
        }

        var formatItems = _formats.ToList();

        if (!TryOpenClipboard(IntPtr.Zero))
            return;

        var snapshot = new Dictionary<string, FormatSnapshot>(StringComparer.Ordinal);
        try
        {
            foreach (var item in formatItems)
            {
                var handleType = GetHandleType(item.FormatId);
                byte[]? data   = null;
                if (handleType != "none")
                    TryReadClipboardDataBytes(item.FormatId, out data, out _);

                var originalSize = data?.LongLength
                    ?? (item.ContentSizeValue >= 0 ? item.ContentSizeValue : 0L);

                snapshot[item.Name] = new FormatSnapshot(item.Ordinal, item.FormatId, item.Name,
                                                         handleType, data, originalSize);
            }
        }
        finally
        {
            NativeMethods.CloseClipboard();
        }

        _staticSnapshot = snapshot;
    }

    /// <summary>
    /// Long-running background consumer: drains the history channel one item at a time,
    /// guaranteeing writes are serialised in arrival order with no DB contention.
    /// </summary>
    private async Task ProcessHistoryChannelAsync()
    {
        await foreach (var (snapshots, timestamp) in _historyChannel.Reader.ReadAllAsync())
        {
            WriteHistorySession(snapshots, timestamp);
        }
    }

    /// <summary>
    /// Called by the channel consumer on a background thread.
    /// Compresses blobs, writes to history.db, then posts the new list entry to the UI.
    /// </summary>
    private void WriteHistorySession(IReadOnlyList<FormatSnapshot> snapshots, DateTime timestamp)
    {
        try
        {
            var sessionId   = ClipboardHistory.AddSession(snapshots, timestamp);
            var formatsText = ClipboardHistory.BuildFormatsText(snapshots);
            var totalSize   = snapshots.Sum(s => s.OriginalSize);

            Dispatcher.InvokeAsync(() =>
            {
                var entry = new HistoryItem
                {
                    SessionId   = sessionId,
                    Timestamp   = timestamp,
                    FormatsText = formatsText,
                    TotalSize   = totalSize,
                };
                _historyItems.Insert(0, entry);
                HistoryListView.SelectedItem = entry;
            });
        }
        catch
        {
            // Background write failure — silently ignored.
        }
    }

    private void LoadHistoryFromDatabase()
    {
        try
        {
            var sessions = ClipboardHistory.LoadSessions();
            _historyItems.Clear();
            foreach (var session in sessions)
            {
                _historyItems.Add(new HistoryItem
                {
                    SessionId   = session.SessionId,
                    Timestamp   = session.Timestamp,
                    FormatsText = session.FormatsText,
                    TotalSize   = session.TotalSize,
                });
            }
        }
        catch
        {
            // Silently ignore load failures.
        }
    }

    private void ShowHistoryPanel()
    {
        LeftPanelGrid.RowDefinitions[1].Height = new GridLength(5);
        LeftPanelGrid.RowDefinitions[2].Height = new GridLength(180);
    }

    private void HideHistoryPanel()
    {
        LeftPanelGrid.RowDefinitions[1].Height = new GridLength(0);
        LeftPanelGrid.RowDefinitions[2].Height = new GridLength(0);
    }

}




