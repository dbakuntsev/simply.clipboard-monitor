using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System.Threading.Channels;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Threading;

namespace Simply.ClipboardMonitor;

public partial class MainWindow : Window
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
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

    // ── Injected services ──────────────────────────────────────────────────
    private readonly IClipboardReader         _clipboardReader;
    private readonly IClipboardWriter         _clipboardWriter;
    private readonly ITextDecodingService     _textDecoding;
    private readonly IImagePreviewService     _imagePreviews;
    private readonly IFormatClassifier        _formatClassifier;
    private readonly IPreferencesService      _preferences;
    private readonly IHistoryRepository       _history;
    private readonly IHistoryMaintenance      _historyMaintenance;
    private readonly IClipboardFileRepository _clipboardFiles;
    private readonly IReadOnlyList<IFormatExporter> _formatExporters;

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

    // System tray state
    private const int WM_TRAYICON = 0x8001; // WM_APP + 1
    private System.Drawing.Icon? _trayIcon;
    private bool   _trayIconAdded;
    private bool   _minimizeToSystemTray;
    private bool   _trayBalloonShown;
    private bool   _isExiting;

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

    // History limits (kept in sync with UserPreferences)
    private const int DefaultHistoryMaxEntries = 100;
    private const int DefaultHistoryMaxSizeMb  = 100;
    private int _historyMaxEntries = DefaultHistoryMaxEntries;
    private int _historyMaxSizeMb  = DefaultHistoryMaxSizeMb;

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
    // Channel that serialises all DB operations through one background consumer.
    // Each item is an Action so both session writes and limit-enforcement can share the queue.
    private readonly Channel<Action> _historyChannel =
        Channel.CreateUnbounded<Action>();

    private sealed class HistoryItem
    {
        public required long     SessionId   { get; init; }
        public required DateTime Timestamp   { get; init; }
        public required string   FormatsText { get; init; }
        public required long     TotalSize   { get; init; }
        public required IReadOnlyList<(uint FormatId, string FormatName)> Formats { get; init; }

        public IReadOnlyList<FormatPill> Pills         { get; init; } = [];
        public string                    FormatsTooltip { get; init; } = string.Empty;

        public string DateText => Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
        public string SizeText => DisplayHelper.FormatFileSize(TotalSize);
    }

    public MainWindow(
        IClipboardReader         clipboardReader,
        IClipboardWriter         clipboardWriter,
        ITextDecodingService     textDecoding,
        IImagePreviewService     imagePreviews,
        IFormatClassifier        formatClassifier,
        IPreferencesService      preferences,
        IHistoryRepository       history,
        IHistoryMaintenance      historyMaintenance,
        IClipboardFileRepository clipboardFiles,
        IEnumerable<IFormatExporter> formatExporters)
    {
        _clipboardReader    = clipboardReader;
        _clipboardWriter    = clipboardWriter;
        _textDecoding       = textDecoding;
        _imagePreviews      = imagePreviews;
        _formatClassifier   = formatClassifier;
        _preferences        = preferences;
        _history            = history;
        _historyMaintenance = historyMaintenance;
        _clipboardFiles     = clipboardFiles;
        _formatExporters    = formatExporters.ToList().AsReadOnly();

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
        FormatListBox.ItemsSource   = _formats;
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

        if (_minimizeToSystemTray)
            InitializeTrayIcon();

        RefreshFormats();

        if (!_isMonitoring)
            CaptureStaticSnapshot();

        if (_isTrackingHistory)
            LoadHistoryFromDatabase();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_minimizeToSystemTray && !_isExiting)
        {
            e.Cancel = true;
            Hide();
            ShowTrayBalloonIfNeeded();
            return;
        }

        base.OnClosing(e);
    }

    protected override void OnClosed(EventArgs e)
    {
        DisposeTrayIcon();

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

    // ── System tray ─────────────────────────────────────────────────────────

    private unsafe void InitializeTrayIcon()
    {
        if (_trayIconAdded || _hwndSource == null)
            return;

        using (var stream = this.GetType().Assembly.GetManifestResourceStream(this.GetType().Namespace + ".Resources.icon.ico"))
        {
            if (stream != null)
                _trayIcon = new System.Drawing.Icon(stream);
        }

        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize           = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd             = _hwndSource.Handle,
            uID              = 1,
            uFlags           = NativeMethods.NIF_MESSAGE | NativeMethods.NIF_ICON | NativeMethods.NIF_TIP,
            uCallbackMessage = (uint)WM_TRAYICON,
            hIcon            = (_trayIcon ?? System.Drawing.SystemIcons.Application).Handle,
        };
        CopyToFixed(nid.szTip, 128, "Simply.ClipboardMonitor");
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_ADD, ref nid);
        _trayIconAdded = true;
    }

    private unsafe void DisposeTrayIcon()
    {
        if (!_trayIconAdded || _hwndSource == null)
            return;

        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd   = _hwndSource.Handle,
            uID    = 1,
        };
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_DELETE, ref nid);
        _trayIconAdded  = false;
        _trayIcon?.Dispose();
        _trayIcon = null;
    }

    private unsafe void ShowTrayContextMenu()
    {
        if (_hwndSource == null)
            return;

        var hMenu = NativeMethods.CreatePopupMenu();
        if (hMenu == IntPtr.Zero)
            return;

        try
        {
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_STRING,    new UIntPtr(1), "Show/Hide Window");
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero,   null);
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_STRING,    new UIntPtr(2), "Exit");

            NativeMethods.GetCursorPos(out var pt);

            // Required so the menu dismisses when the user clicks elsewhere.
            NativeMethods.SetForegroundWindow(_hwndSource.Handle);

            const uint TPM_RETURNCMD   = 0x0100;
            const uint TPM_RIGHTBUTTON = 0x0002;
            var cmd = NativeMethods.TrackPopupMenuEx(
                hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                pt.x, pt.y, _hwndSource.Handle, IntPtr.Zero);

            // Post a benign message so the SetForegroundWindow trick works correctly.
            NativeMethods.PostMessage(_hwndSource.Handle, 0 /* WM_NULL */, IntPtr.Zero, IntPtr.Zero);

            switch (cmd)
            {
                case 1: ToggleTrayWindowVisibility(); break;
                case 2: ExitApplication();            break;
            }
        }
        finally
        {
            NativeMethods.DestroyMenu(hMenu);
        }
    }

    private void ToggleTrayWindowVisibility()
    {
        if (IsVisible)
        {
            Hide();
        }
        else
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        }
    }

    private unsafe void ShowTrayBalloonIfNeeded()
    {
        if (_trayBalloonShown || !_trayIconAdded || _hwndSource == null)
            return;

        _trayBalloonShown = true;
        SavePreferences();

        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize      = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd        = _hwndSource.Handle,
            uID         = 1,
            uFlags      = NativeMethods.NIF_INFO,
            dwInfoFlags = NativeMethods.NIIF_INFO,
        };
        CopyToFixed(nid.szInfoTitle, 64,  "Simply.ClipboardMonitor");
        CopyToFixed(nid.szInfo,      256, "The application is still running in the system tray.");
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_MODIFY, ref nid);
    }

    private void ExitApplication()
    {
        _isExiting = true;
        DisposeTrayIcon();
        Close();
    }

    private static unsafe void CopyToFixed(char* dest, int maxChars, string value)
    {
        int len = Math.Min(value.Length, maxChars - 1);
        value.AsSpan(0, len).CopyTo(new Span<char>(dest, len));
        dest[len] = '\0';
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
                var seqAtArrival = _clipboardReader.GetSequenceNumber();
                RefreshFormats();
                if (_isTrackingHistory)
                    CaptureHistorySession(seqAtArrival);
            }
            handled = true;
        }
        else if (msg == WM_TRAYICON)
        {
            const int WM_LBUTTONUP = 0x0202;
            const int WM_RBUTTONUP = 0x0205;
            var notification = (int)(lParam.ToInt64() & 0xFFFF);
            if (notification == WM_LBUTTONUP)
            {
                ToggleTrayWindowVisibility();
                handled = true;
            }
            else if (notification == WM_RBUTTONUP)
            {
                ShowTrayContextMenu();
                handled = true;
            }
        }

        return IntPtr.Zero;
    }

    private void RefreshFormats()
    {
        _historySnapshots = null;   // switch to live-clipboard mode
        InitializePreviewState();

        var selectedId = (FormatListBox.SelectedItem as ClipboardFormatItem)?.FormatId;
        _formats.Clear();

        foreach (var item in _clipboardReader.EnumerateFormats())
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
        TextTabItem.IsEnabled  = _textDecoding.IsTextCompatible(item.FormatId, item.Name);
        ImageTabItem.IsEnabled = _imagePreviews.IsImageCompatible(item.FormatId, item.Name);

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
            if (!_clipboardReader.TryOpenClipboard(IntPtr.Zero))
            {
                SetHexPreviewUnavailable("Unable to open clipboard.");
                return;
            }

            try
            {
                var handleType = _clipboardReader.GetHandleType(item.FormatId);
                _clipboardReader.TryReadFormatBytes(item.FormatId, handleType,
                    out bytes, out hexFailureMessage);
            }
            finally
            {
                _clipboardReader.CloseClipboard();
            }
        }

        _currentFormatBytes = bytes;

        if (bytes != null)
        {
            SetHexPreview(bytes);
            UpdateTextPreview(item.FormatId, item.Name, bytes);

            if (_imagePreviews.TryCreatePreview(item.FormatId, item.Name, bytes,
                    out imagePreview, out imageFailureMessage) && imagePreview != null)
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

        var result = _textDecoding.Decode(formatId, formatName, bytes);

        if (!result.Success)
        {
            SetTextPreviewUnavailable(result.FailureMessage ?? "Text preview unavailable for this format.");
            return;
        }

        _currentAutoDetectedEncoding    = result.DetectedEncoding;
        TextContentTextBox.Foreground   = SystemColors.WindowTextBrush;
        TextContentTextBox.Text         = result.Text ?? string.Empty;
        TextStatusTextBlock.Text        = _textDecoding.GetDecodedTextStats(result.Text ?? string.Empty);
        PopulateEncodingComboBox(result.DetectedEncoding);
    }

    private void PopulateEncodingComboBox(Encoding? preselect)
    {
        _encodingItems ??= _textDecoding.GetAvailableEncodings();

        if (TextEncodingComboBox.ItemsSource == null)
            TextEncodingComboBox.ItemsSource = _encodingItems;

        TextEncodingComboBox.IsEnabled = true;

        _suppressEncodingChange = true;
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

        var result = _textDecoding.DecodeWith(_currentTextBytes, item.Encoding);

        if (result.Success)
        {
            TextContentTextBox.Text       = result.Text ?? string.Empty;
            TextContentTextBox.Foreground = SystemColors.WindowTextBrush;
            TextStatusTextBlock.Text      = _textDecoding.GetDecodedTextStats(result.Text ?? string.Empty);
        }
        else
        {
            TextContentTextBox.Text       = $"Cannot decode as {item.DisplayName}:\n{result.FailureMessage}";
            TextContentTextBox.Foreground = Brushes.Crimson;
            TextStatusTextBlock.Text      = "Decoding failed.";
        }
    }

    private void SetImagePreview(BitmapSource image)
    {
        ImagePreview.Source = image;
        ImageStatusTextBlock.Text = string.Empty;
        ZoomSlider.IsEnabled = true;
        ImageDimensionsWidthTextBlock.Text  = $"{image.PixelWidth}px";
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
        TextTabItem.IsEnabled  = true;
        ImageTabItem.IsEnabled = true;
        ContentTabControl.SelectedIndex = 0;
        ContentTabControl.Visibility = Visibility.Collapsed;
        NoSelectionPanel.Visibility  = Visibility.Visible;
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

        var viewportWidth  = ImageScrollViewer.ViewportWidth;
        var viewportHeight = ImageScrollViewer.ViewportHeight;
        if (viewportWidth <= 0 || viewportHeight <= 0)
        {
            _fitScale = 1;
            ApplyImageScale();
            return;
        }

        var fitX  = viewportWidth  / source.PixelWidth;
        var fitY  = viewportHeight / source.PixelHeight;
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
        ZoomSlider.Value   = value;
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

        var mousePos     = e.GetPosition(ImageScrollViewer);
        var currentScale = _fitScale * ZoomSlider.Value;
        var contentX     = ImageScrollViewer.HorizontalOffset + mousePos.X;
        var contentY     = ImageScrollViewer.VerticalOffset   + mousePos.Y;
        var scaledW      = source.PixelWidth  * currentScale;
        var scaledH      = source.PixelHeight * currentScale;
        var imageLeft    = Math.Max(0.0, (ImageScrollViewer.ViewportWidth  - scaledW) / 2.0);
        var imageTop     = Math.Max(0.0, (ImageScrollViewer.ViewportHeight - scaledH) / 2.0);
        var imgLocalX    = (contentX - imageLeft) / currentScale;
        var imgLocalY    = (contentY - imageTop)  / currentScale;

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
        _panStartMouse            = e.GetPosition(ImageScrollViewer);
        _panStartHorizontalOffset = ImageScrollViewer.HorizontalOffset;
        _panStartVerticalOffset   = ImageScrollViewer.VerticalOffset;
        ImageScrollViewer.Cursor  = Cursors.ScrollAll;
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

    private void UpdateStatusBar()
    {
        if (_isMonitoring && _isTrackingHistory)
        {
            var size = _historyMaintenance.GetDatabaseFileSize();
            StatusText.Text = $"Monitoring... · Tracking history ({DisplayHelper.FormatFileSize(size)} storage size)...";
        }
        else
        {
            StatusText.Text = _isMonitoring ? "Monitoring..." : "Press F5 to refresh";
        }
    }


    private void UpdateFileStatusBar(string action, string path)
    {
        FileStatusText.Text = $"{action}: {Path.GetFileName(path)}";
        FileStatusSeparator.Visibility = Visibility.Visible;
    }

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
        ExitApplication();
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
            _isTrackingHistory             = false;
            MenuItemTrackHistory.IsChecked = false;
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

    private void MenuItemSettings_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SettingsDialog(_historyMaxEntries, _historyMaxSizeMb, _historyMaintenance, _minimizeToSystemTray) { Owner = this };
        var result = dlg.ShowDialog();

        if (dlg.HistoryWasCleared)
        {
            _historyItems.Clear();
            if (_historySnapshots != null)
            {
                _historySnapshots = null;
                RefreshFormats();
            }
        }

        if (result == true)
        {
            _historyMaxEntries = dlg.MaxEntries;
            _historyMaxSizeMb  = dlg.MaxSizeMb;

            if (dlg.MinimizeToSystemTray != _minimizeToSystemTray)
            {
                _minimizeToSystemTray = dlg.MinimizeToSystemTray;
                if (_minimizeToSystemTray)
                    InitializeTrayIcon();
                else
                    DisposeTrayIcon();
            }

            SavePreferences();

            // Enforce new limits against any existing database, regardless of whether
            // history tracking is currently on (old data may still need trimming).
            if (_historyChannel.Reader.Count == 0)
            {
                var maxEntries      = _historyMaxEntries;
                var maxBytes        = (long)_historyMaxSizeMb * 1024L * 1024L;
                var trackingHistory = _isTrackingHistory;
                _historyChannel.Writer.TryWrite(() =>
                {
                    var removed = _historyMaintenance.EnforceLimits(maxEntries, maxBytes);
                    if (removed || trackingHistory)
                    {
                        Dispatcher.InvokeAsync(() =>
                        {
                            if (removed && trackingHistory)
                            {
                                if (_historySnapshots != null)
                                {
                                    _historySnapshots = null;
                                    RefreshFormats();
                                }
                                LoadHistoryFromDatabase();
                            }
                            UpdateStatusBar();
                        });
                    }
                });
            }
        }
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

        Encoding? manuallySelected = _textEncodingManuallyChanged
            && ContentTabControl.SelectedItem == TextTabItem
            && TextEncodingComboBox.SelectedItem is EncodingItem { Encoding: var enc }
            ? enc : null;

        var ctx = new FormatExportContext(
            Bytes:                    _currentFormatBytes,
            FormatId:                 _currentFormatId,
            FormatName:               _currentFormatName,
            AutoDetectedEncoding:     _currentAutoDetectedEncoding,
            ManuallySelectedEncoding: manuallySelected,
            ImagePreviewSource:       ImagePreview.Source as BitmapSource);

        var applicable = _formatExporters.Where(exp => exp.CanExport(ctx)).ToList();
        if (applicable.Count == 0)
            return;

        // Default: text > jpeg/png (by format name hint) > binary
        int defaultIndex = applicable.FindIndex(exp => exp.Extension == ".txt");
        if (defaultIndex < 0)
        {
            var norm          = _currentFormatName.ToLowerInvariant();
            var preferredExt  = norm.Contains("jpeg") || norm.Contains("jpg") ? ".jpg" : ".png";
            defaultIndex      = applicable.FindIndex(exp => exp.Extension == preferredExt);
        }
        if (defaultIndex < 0)
            defaultIndex = applicable.Count - 1;  // fallback to last (binary)

        var selectedClipboardItem = FormatListBox.SelectedItem as ClipboardFormatItem;
        var fileName = $"clipboard-{new string((selectedClipboardItem?.Name ?? "CF_UNKNOWN").Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch).ToArray())}-{DateTime.Now:yyyyMMddTHHmmss}";

        var dlg = new SaveFileDialog
        {
            Title           = $"Export Selected Clipboard Format: {selectedClipboardItem?.Name ?? "CF_UNKNOWN"}",
            Filter          = string.Join("|", applicable.Select(exp => exp.FilterLabel)),
            FilterIndex     = defaultIndex + 1,
            DefaultExt      = applicable[defaultIndex].Extension.TrimStart('.'),
            FileName        = fileName,
            OverwritePrompt = true,
        };

        if (dlg.ShowDialog(this) != true)
            return;

        try
        {
            applicable[dlg.FilterIndex - 1].Export(dlg.FileName, ctx);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Failed to export:\n\n{ex.Message}",
                "Export Failed", MessageBoxButton.OK, MessageBoxImage.Error);
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

        if (!_clipboardReader.TryOpenClipboard(IntPtr.Zero))
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
                var handleType = _clipboardReader.GetHandleType(item.FormatId);
                byte[]? data   = null;
                if (handleType != "none")
                    _clipboardReader.TryReadFormatBytes(item.FormatId, handleType, out data, out _);

                formats.Add(new SavedClipboardFormat(item.Ordinal, item.FormatId, item.Name, handleType, data));
            }
        }
        finally
        {
            _clipboardReader.CloseClipboard();
        }

        _clipboardFiles.Save(path, formats);
        _currentFilePath = path;
        UpdateFileStatusBar("Saved", path);
    }

    // ── Load implementation ─────────────────────────────────────────────────

    private void LoadFromFile(string path)
    {
        var formats = _clipboardFiles.Load(path);

        if (formats.Count == 0)
        {
            MessageBox.Show(this,
                "The clipboard database is empty — no formats were stored in the file.",
                "Nothing to Load", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (!_clipboardReader.TryOpenClipboard(_hwndSource?.Handle ?? IntPtr.Zero))
        {
            MessageBox.Show(this,
                "Unable to open the clipboard for writing.\n\nClose any application that may be using the clipboard and try again.",
                "Load Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            NativeMethods.EmptyClipboard();
            _clipboardWriter.RestoreFormats(formats);
        }
        finally
        {
            _clipboardReader.CloseClipboard();
        }

        _currentFilePath = path;
        UpdateFileStatusBar("Loaded", path);

        if (!_isMonitoring)
            RefreshFormats();
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
            _currentSortProperty  = requestedProperty;
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

        var preferences = _preferences.Load();

        try
        {
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
                        Key   = column.Key,
                        Width = column.Width
                    });
                }
            }

            _isMonitoring         = preferences.MonitorChanges;
            _isTrackingHistory    = preferences.TrackHistory && _isMonitoring;
            _historyMaxEntries    = preferences.HistoryMaxEntries > 0 ? preferences.HistoryMaxEntries : DefaultHistoryMaxEntries;
            _historyMaxSizeMb     = preferences.HistoryMaxSizeMb  > 0 ? preferences.HistoryMaxSizeMb  : DefaultHistoryMaxSizeMb;
            _minimizeToSystemTray = preferences.MinimizeToSystemTray;
            _trayBalloonShown     = preferences.TrayBalloonShown;
        }
        catch
        {
            _currentSortProperty     = DefaultSortProperty;
            _currentSortDirection    = ListSortDirection.Ascending;
            _storedColumnPreferences = [];
            _isMonitoring            = true;
            _isTrackingHistory       = false;
            _historyMaxEntries       = DefaultHistoryMaxEntries;
            _historyMaxSizeMb        = DefaultHistoryMaxSizeMb;
        }
    }

    private void SavePreferences()
    {
        var preferences = new UserPreferences
        {
            SortProperty         = _currentSortProperty,
            SortDirection        = _currentSortDirection.ToString(),
            FormatColumns        = CaptureFormatColumnPreferences(),
            MonitorChanges       = _isMonitoring,
            TrackHistory         = _isTrackingHistory,
            HistoryMaxEntries    = _historyMaxEntries,
            HistoryMaxSizeMb     = _historyMaxSizeMb,
            MinimizeToSystemTray = _minimizeToSystemTray,
            TrayBalloonShown     = _trayBalloonShown,
        };

        _preferences.Save(preferences);
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
        var result   = new List<FormatColumnPreference>();
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
                Key   = key,
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
        IdColumnHeader.Content      = "ID";
        FormatColumnHeader.Content  = "Format";
        SizeColumnHeader.Content    = "Size";

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
            nameof(ClipboardFormatItem.Ordinal)          => OrdinalColumnHeader,
            nameof(ClipboardFormatItem.FormatId)         => IdColumnHeader,
            nameof(ClipboardFormatItem.Name)             => FormatColumnHeader,
            nameof(ClipboardFormatItem.ContentSizeValue) => SizeColumnHeader,
            _ => null
        };
    }

    private static string GetHeaderBaseTitle(string propertyName)
    {
        return propertyName switch
        {
            nameof(ClipboardFormatItem.Ordinal)          => "#",
            nameof(ClipboardFormatItem.FormatId)         => "ID",
            nameof(ClipboardFormatItem.Name)             => "Format",
            nameof(ClipboardFormatItem.ContentSizeValue) => "Size",
            _ => propertyName
        };
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

        UpdateStatusBar();
        SavePreferences();
    }

    private void HistoryListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not HistoryItem entry)
            return;

        List<FormatSnapshot> snapshots;
        try
        {
            snapshots = _history.LoadSessionFormats(entry.SessionId);
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

        var snapshots = _clipboardReader.CaptureAllFormats(formatItems);
        if (snapshots.Count == 0)
            return;

        // If the sequence number advanced since WM_CLIPBOARDUPDATE arrived, a
        // delayed-rendering provider rendered data and posted a fresh
        // WM_CLIPBOARDUPDATE.  Skip this capture; the next message will capture
        // the fully-rendered clipboard and avoids creating a duplicate entry.
        if (_clipboardReader.GetSequenceNumber() != seqAtArrival)
            return;

        _historyChannel.Writer.TryWrite(() => WriteHistorySession(snapshots, timestamp));
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

        var snapshots = _clipboardReader.CaptureAllFormats(_formats.ToList());
        _staticSnapshot = snapshots.Count == 0
            ? null
            : snapshots.ToDictionary(s => s.FormatName, StringComparer.Ordinal);
    }

    /// <summary>
    /// Long-running background consumer: drains the history channel one action at a time,
    /// guaranteeing all DB operations are serialised in arrival order with no contention.
    /// </summary>
    private async Task ProcessHistoryChannelAsync()
    {
        await foreach (var action in _historyChannel.Reader.ReadAllAsync())
        {
            action();
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
            var maxDatabaseBytes = (long)_historyMaxSizeMb * 1024L * 1024L;
            var (sessionId, trimmed) = _history.AddSession(snapshots, timestamp, _historyMaxEntries, maxDatabaseBytes);
            var formatsText = _history.BuildFormatsText(snapshots);
            var totalSize   = snapshots.Sum(s => s.OriginalSize);

            var formats = snapshots.Select(s => (s.FormatId, s.FormatName)).ToList();

            // Build pills and tooltip on the background thread (they are read-only
            // after construction and the service is thread-safe).
            var pills   = _formatClassifier.ComputePills(formats);
            var tooltip = _formatClassifier.ComputeTooltip(formats);

            Dispatcher.InvokeAsync(() =>
            {
                if (trimmed)
                {
                    // Old sessions were deleted — reload the list so stale entries
                    // are removed, then select the newly-added entry at the top.
                    LoadHistoryFromDatabase();
                    if (_historyItems.Count > 0)
                        HistoryListView.SelectedItem = _historyItems[0];
                }
                else
                {
                    var entry = new HistoryItem
                    {
                        SessionId      = sessionId,
                        Timestamp      = timestamp,
                        FormatsText    = formatsText,
                        TotalSize      = totalSize,
                        Formats        = formats,
                        Pills          = pills,
                        FormatsTooltip = tooltip,
                    };
                    _historyItems.Insert(0, entry);
                    HistoryListView.SelectedItem = entry;
                }
                UpdateStatusBar();
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
            var sessions = _history.LoadSessions();
            _historyItems.Clear();
            foreach (var session in sessions)
            {
                var pills   = _formatClassifier.ComputePills(session.Formats);
                var tooltip = _formatClassifier.ComputeTooltip(session.Formats);

                _historyItems.Add(new HistoryItem
                {
                    SessionId      = session.SessionId,
                    Timestamp      = session.Timestamp,
                    FormatsText    = session.FormatsText,
                    TotalSize      = session.TotalSize,
                    Formats        = session.Formats,
                    Pills          = pills,
                    FormatsTooltip = tooltip,
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
