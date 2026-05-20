using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using Simply.ClipboardMonitor.Common;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;
using Simply.ClipboardMonitor.ViewModels;
using Simply.ClipboardMonitor.Views.Previews;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
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
    private const int WM_HOTKEY          = 0x0312;
    private const int HotkeyId           = 1;
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

    // ── Grouped state ────────────────────────────────────────────────────────

    private record struct TrayState
    {
        public System.Drawing.Icon? Icon          { get; set; }
        public bool                 IconAdded     { get; set; }
        public bool                 MinimizeToTray { get; set; }
        public bool                 BalloonShown  { get; set; }
        public bool                 IsExiting     { get; set; }
    }

    private record struct FormatSelectionState
    {
        public byte[]? Bytes      { get; set; }
        public uint    FormatId   { get; set; }
        public string  FormatName { get; set; }
    }

    // ── Injected services ──────────────────────────────────────────────────
    private readonly IClipboardReader         _clipboardReader;
    private readonly IClipboardWriter         _clipboardWriter;
    private readonly ITextDecodingService     _textDecoding;
    private readonly IFormatClassifier        _formatClassifier;
    private readonly IPreferencesService      _preferences;
    private readonly IHistoryRepository       _history;
    private readonly IHistoryMaintenance      _historyMaintenance;
    private readonly IClipboardFileRepository _clipboardFiles;
    private readonly IDialogService           _dialogService;
    private readonly IReadOnlyList<IFormatExporter> _formatExporters;
    private readonly IReadOnlyList<IPreviewTab>     _previewTabs;
    private readonly IClipboardOwnerService         _clipboardOwner;
    private readonly IFormatNotificationMatcher     _formatNotificationMatcher;

    // ── ViewModel ─────────────────────────────────────────────────────────
    private readonly MainWindowViewModel _vm;

    private readonly ObservableCollection<ClipboardFormatItem> _formats = [];
    private HwndSource? _hwndSource;
    private string _currentSortProperty = DefaultSortProperty;
    private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
    private List<FormatColumnPreference> _storedColumnPreferences = [];

    // Tracks the path of the most recently loaded or saved .clipdb file.
    private string? _currentFilePath;

    // System tray state
    private const int WM_TRAYICON = 0x8001; // WM_APP + 1
    private TrayState _trayState;

    // Global hotkey state
    private bool         _hotkeyEnabled = true;
    private HotkeyBinding _hotkeyBinding = HotkeyBinding.Default;
    private bool         _hotkeyRegistered;
    private bool         _hotkeyConflict;

    // Auto-start / start-minimized state
    private bool _startAtLogin;
    private bool _startMinimized;

    // Format-arrival notification state
    private bool    _formatNotificationsEnabled;
    private string? _formatNotificationPatterns;

    // Selected format state — raw bytes and identity of the currently selected format
    private FormatSelectionState _formatState = new() { FormatName = string.Empty };

    // History filter debounce state
    private CancellationTokenSource? _filterCts;

    // Drag-from-history state
    private System.Windows.Point? _historyDragStartPoint;

    // Non-null when showing a history session (maps format name → snapshot).
    // Null means the live clipboard or the static snapshot is used instead.
    private IReadOnlyDictionary<string, FormatSnapshot>? _historySnapshots;
    // Non-null when monitoring is off: holds a byte-level copy of every format
    // taken at the moment monitoring was disabled (or on a manual refresh).
    // Used instead of the live clipboard so format previews still work.
    private Dictionary<string, FormatSnapshot>? _staticSnapshot;

    public MainWindow(
        IClipboardReader         clipboardReader,
        IClipboardWriter         clipboardWriter,
        ITextDecodingService     textDecoding,
        IFormatClassifier        formatClassifier,
        IPreferencesService      preferences,
        IHistoryRepository       history,
        IHistoryMaintenance      historyMaintenance,
        IClipboardFileRepository clipboardFiles,
        IClipboardOwnerService   clipboardOwner,
        IDialogService           dialogService,
        IFormatNotificationMatcher formatNotificationMatcher,
        IEnumerable<IFormatExporter> formatExporters,
        IEnumerable<IPreviewTab>     previewTabs)
    {
        _clipboardReader           = clipboardReader;
        _clipboardWriter           = clipboardWriter;
        _textDecoding              = textDecoding;
        _formatClassifier          = formatClassifier;
        _preferences               = preferences;
        _history                   = history;
        _historyMaintenance        = historyMaintenance;
        _clipboardFiles            = clipboardFiles;
        _clipboardOwner            = clipboardOwner;
        _dialogService             = dialogService;
        _formatNotificationMatcher = formatNotificationMatcher;
        _formatExporters           = formatExporters.ToList().AsReadOnly();
        _previewTabs               = previewTabs.ToList().AsReadOnly();

        // Required for legacy OEM/ANSI code pages (e.g. CP437 used by CF_OEMTEXT) which
        // are not included in the default .NET encoding set.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        InitializeComponent();

        // ── Create ViewModel (after InitializeComponent so XAML elements exist for callbacks) ──
        _vm = new MainWindowViewModel(
            history, historyMaintenance, formatClassifier, textDecoding, dialogService, Dispatcher,
            setHistoryPanelVisible:          show => { if (show) ShowHistoryPanel(); else HideHistoryPanel(); },
            setHistoryLoadingOverlayVisible: show => { if (show) ShowHistoryLoadingOverlay(); else HideHistoryLoadingOverlay(); },
            clearFormatPanelToEmpty:         () =>  { _historySnapshots = null; InitializePreviewState(); _formats.Clear(); },
            clearSnapshotsAndRefreshToLive:  () =>  { if (_historySnapshots != null) { _historySnapshots = null; RefreshFormats(); } },
            setSelectedHistoryItem:          item => { HistoryListView.SelectedItem = item; if (item != null) HistoryListView.ScrollIntoView(item); },
            setMainWindowEnabled:            enabled => IsEnabled = enabled,
            showWarning:                     (msg, title) => MessageBox.Show(this, msg, title, MessageBoxButton.OK, MessageBoxImage.Warning),
            savePreferences:                 SavePreferences);

        DataContext = _vm;

        // Register each preview tab's TabItem with the tab control.
        foreach (var tab in _previewTabs)
            ContentTabControl.Items.Add(tab.TabItem);

        CommandBindings.Add(new CommandBinding(RefreshClipboardCommand, RefreshClipboardCommand_Executed, RefreshClipboardCommand_CanExecute));
        CommandBindings.Add(new CommandBinding(LoadCommand, LoadCommand_Executed));
        CommandBindings.Add(new CommandBinding(SaveCommand, SaveCommand_Executed));
        CommandBindings.Add(new CommandBinding(ExportCommand, ExportCommand_Executed, ExportCommand_CanExecute));
        _ = Task.Run(_vm.ProcessHistoryChannelAsync);
        FormatListBox.ItemsSource = _formats;
        LoadPreferences();
        if (_vm.IsTrackingHistory)
            ShowHistoryPanel();
        _vm.UpdateStatusBarText();
        UpdateStatusBar();
        ApplyFormatColumnPreferences();
        FormatListBox.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler(FormatListBox_OnColumnHeaderClick));
        HistoryListView.PreviewMouseLeftButtonDown += HistoryListView_PreviewMouseLeftButtonDown;
        HistoryListView.PreviewMouseMove           += HistoryListView_PreviewMouseMove;
        HistoryListView.MouseLeftButtonUp          += (_, _) => _historyDragStartPoint = null;
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

        RecomputeTrayIconPresence();
        if (_trayState.MinimizeToTray && _hotkeyEnabled)
            TryRegisterHotkey();
        UpdateHotkeyStatusBar();

        if (_startMinimized)
        {
            if (_trayState.MinimizeToTray)
                // Defer past the Show() call so the window is hidden before it appears.
                Dispatcher.BeginInvoke(() => { Hide(); ShowTrayBalloonIfNeeded(); });
            else
                WindowState = WindowState.Minimized;
        }

        RefreshFormats();

        if (!_vm.IsMonitoring)
            CaptureStaticSnapshot();

        // Run integrity check and schema migration on the history channel so they are
        // serialised with all subsequent DB writes.  LoadHistoryFromDatabase and
        // CaptureStartupSession are deferred until after migration completes — they run
        // as the onReady callback so they never race against an incompatible schema.
        var isTracking = _vm.IsTrackingHistory;
        if (isTracking)
            ShowHistoryLoadingOverlay();
        Action? onReady = isTracking
            ? () => Dispatcher.Invoke(() =>
              {
                  HideHistoryLoadingOverlay();
                  CaptureStartupSession();
                  _vm.LoadHistoryFromDatabase(selectLive: true);
              })
            : null;
        _vm.EnqueueDatabaseInitialization(isTracking, onReady);
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_trayState.MinimizeToTray && !_trayState.IsExiting)
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
        UnregisterHotkey();
        DisposeTrayIcon();

        if (_hwndSource != null)
        {
            NativeMethods.RemoveClipboardFormatListener(_hwndSource.Handle);
            _hwndSource.RemoveHook(WndProc);
            _hwndSource = null;
        }

        _vm.HistoryChannel.Writer.TryComplete();
        SavePreferences();
        base.OnClosed(e);
    }

    // ── System tray ─────────────────────────────────────────────────────────

    private unsafe void InitializeTrayIcon()
    {
        if (_trayState.IconAdded || _hwndSource == null)
            return;

        using (var stream = this.GetType().Assembly.GetManifestResourceStream(this.GetType().Namespace + ".Resources.icon.ico"))
        {
            if (stream != null)
                _trayState.Icon = new System.Drawing.Icon(stream);
        }

        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize           = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd             = _hwndSource.Handle,
            uID              = 1,
            uFlags           = NativeMethods.NIF_MESSAGE | NativeMethods.NIF_ICON | NativeMethods.NIF_TIP,
            uCallbackMessage = (uint)WM_TRAYICON,
            hIcon            = (_trayState.Icon ?? System.Drawing.SystemIcons.Application).Handle,
        };
        CopyToFixed(nid.szTip, 128, "Simply.ClipboardMonitor");
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_ADD, ref nid);
        _trayState.IconAdded = true;
    }

    private unsafe void DisposeTrayIcon()
    {
        if (!_trayState.IconAdded || _hwndSource == null)
            return;

        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd   = _hwndSource.Handle,
            uID    = 1,
        };
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_DELETE, ref nid);
        _trayState.IconAdded  = false;
        _trayState.Icon?.Dispose();
        _trayState.Icon = null;
    }

    private void ApplyTraySettings(bool newValue)
    {
        _trayState.MinimizeToTray = newValue;
        RecomputeTrayIconPresence();
        if (_trayState.MinimizeToTray)
        {
            RefreshHotkeyRegistration();
        }
        else
        {
            UnregisterHotkey();
            _hotkeyConflict = false;
            UpdateHotkeyStatusBar();
        }
    }

    /// <summary>
    /// Adds or removes the system-tray icon to match the current feature set.
    /// The icon must be present whenever the user enables hide-to-tray <em>or</em>
    /// format-arrival notifications (the latter relies on Shell_NotifyIcon balloons).
    /// </summary>
    private void RecomputeTrayIconPresence()
    {
        var needed = _trayState.MinimizeToTray || _formatNotificationsEnabled;
        if (needed && !_trayState.IconAdded)
            InitializeTrayIcon();
        else if (!needed && _trayState.IconAdded)
            DisposeTrayIcon();
    }

    private void ApplyAutoStartPreference(bool newValue)
    {
        _startAtLogin = newValue;
        AutoStartHelper.SetAutoStart(_startAtLogin);
        AutoStartPill.Visibility = _startAtLogin ? Visibility.Visible : Visibility.Collapsed;
        UpdateOwnerPillSeparator();
        UpdateHotkeyStatusBar();
    }

    private enum TrayCmd : uint
    {
        ShowHideWindow        = 1,
        Exit                  = 2,
        Load                  = 10,
        Save                  = 11,
        SaveAs                = 12,
        ExportSelectedFormat  = 13,
        Settings              = 14,
        Clear                 = 20,
        Refresh               = 21,
        MonitorChanges        = 22,
        TrackHistory          = 23,
        Resume                = 24,
        Pause1m               = 30,
        Pause5m               = 31,
        Pause10m              = 32,
        Pause30m              = 33,
        Pause1h               = 34,
        SubmitFeedback        = 40,
        About                 = 41,
        PausedUntilLabel      = 50,
    }

    private unsafe void ShowTrayContextMenu()
    {
        if (_hwndSource == null)
            return;

        var hMenu = NativeMethods.CreatePopupMenu();
        if (hMenu == IntPtr.Zero)
            return;

        static UIntPtr Cmd(TrayCmd c) => (UIntPtr)(uint)c;

        try
        {
            // ── Show/Hide Window (top) ───────────────────────────────────────
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_STRING,    Cmd(TrayCmd.ShowHideWindow), "Show/Hide Window");
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero,                null);

            // ── File submenu ─────────────────────────────────────────────────
            var hFileMenu = NativeMethods.CreatePopupMenu();
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Load),   "Load...");
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Save),   "Save");
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.SaveAs), "Save As...");
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero, null);
            NativeMethods.AppendMenu(hFileMenu,
                NativeMethods.MF_STRING | (_formatState.Bytes is { Length: > 0 } ? 0u : NativeMethods.MF_GRAYED),
                Cmd(TrayCmd.ExportSelectedFormat), "Export Selected Format...");
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero, null);
            NativeMethods.AppendMenu(hFileMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Settings), "Settings...");
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_POPUP, (UIntPtr)(nuint)(nint)hFileMenu, "File");

            // ── Clipboard submenu ────────────────────────────────────────────
            var hClipboardMenu = NativeMethods.CreatePopupMenu();
            NativeMethods.AppendMenu(hClipboardMenu,
                NativeMethods.MF_STRING | (NativeMethods.CountClipboardFormats() > 0 ? 0u : NativeMethods.MF_GRAYED),
                Cmd(TrayCmd.Clear), "Clear");
            NativeMethods.AppendMenu(hClipboardMenu,
                NativeMethods.MF_STRING | (_vm.IsMonitoring ? NativeMethods.MF_GRAYED : 0u),
                Cmd(TrayCmd.Refresh), "Refresh");
            NativeMethods.AppendMenu(hClipboardMenu,
                NativeMethods.MF_STRING | (_vm.IsMonitoring ? NativeMethods.MF_CHECKED : 0u),
                Cmd(TrayCmd.MonitorChanges), "Monitor Changes");
            NativeMethods.AppendMenu(hClipboardMenu,
                NativeMethods.MF_STRING | (_vm.IsTrackingHistory ? NativeMethods.MF_CHECKED : 0u)
                                        | (_vm.IsMonitoring ? 0u : NativeMethods.MF_GRAYED),
                Cmd(TrayCmd.TrackHistory), "Track History");
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_POPUP, (UIntPtr)(nuint)(nint)hClipboardMenu, "Clipboard");

            // ── Help submenu ─────────────────────────────────────────────────
            var hHelpMenu = NativeMethods.CreatePopupMenu();
            NativeMethods.AppendMenu(hHelpMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.SubmitFeedback), "Submit Feedback");
            NativeMethods.AppendMenu(hHelpMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.About),          "About");
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_POPUP, (UIntPtr)(nuint)(nint)hHelpMenu, "Help");

            // ── Pause / Resume (top-level) ───────────────────────────────────
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero, null);
            if (_vm.IsPaused)
            {
                NativeMethods.AppendMenu(hMenu, NativeMethods.MF_STRING | NativeMethods.MF_GRAYED,
                    Cmd(TrayCmd.PausedUntilLabel), $"Paused until {_vm.PauseUntilLabel}");
            }
            var hPauseMenu = NativeMethods.CreatePopupMenu();
            NativeMethods.AppendMenu(hPauseMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Pause1m),  _vm.IsPaused ? "+1 minute"   : "1 minute");
            NativeMethods.AppendMenu(hPauseMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Pause5m),  _vm.IsPaused ? "+5 minutes"  : "5 minutes");
            NativeMethods.AppendMenu(hPauseMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Pause10m), _vm.IsPaused ? "+10 minutes" : "10 minutes");
            NativeMethods.AppendMenu(hPauseMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Pause30m), _vm.IsPaused ? "+30 minutes" : "30 minutes");
            NativeMethods.AppendMenu(hPauseMenu, NativeMethods.MF_STRING, Cmd(TrayCmd.Pause1h),  _vm.IsPaused ? "+1 hour"     : "1 hour");
            NativeMethods.AppendMenu(hMenu,
                NativeMethods.MF_POPUP | (_vm.IsMonitoring ? 0u : NativeMethods.MF_GRAYED),
                (UIntPtr)(nuint)(nint)hPauseMenu, "Pause");
            NativeMethods.AppendMenu(hMenu,
                NativeMethods.MF_STRING | (_vm.IsPaused ? 0u : NativeMethods.MF_GRAYED),
                Cmd(TrayCmd.Resume), "Resume");

            // ── Exit (bottom) ────────────────────────────────────────────────
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_SEPARATOR, UIntPtr.Zero, null);
            NativeMethods.AppendMenu(hMenu, NativeMethods.MF_STRING,    Cmd(TrayCmd.Exit), "Exit");

            NativeMethods.GetCursorPos(out var pt);

            // Required so the menu dismisses when the user clicks elsewhere.
            NativeMethods.SetForegroundWindow(_hwndSource.Handle);

            const uint TPM_RETURNCMD   = 0x0100;
            const uint TPM_RIGHTBUTTON = 0x0002;
            var selected = NativeMethods.TrackPopupMenuEx(
                hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                pt.x, pt.y, _hwndSource.Handle, IntPtr.Zero);

            // Post a benign message so the SetForegroundWindow trick works correctly.
            NativeMethods.PostMessage(_hwndSource.Handle, 0 /* WM_NULL */, IntPtr.Zero, IntPtr.Zero);

            switch ((TrayCmd)selected)
            {
                case TrayCmd.ShowHideWindow: ToggleTrayWindowVisibility(); break;
                case TrayCmd.Exit:           ExitApplication();            break;

                case TrayCmd.Load:                EnsureWindowVisible(); MenuItemLoad_Click(this, new RoutedEventArgs());                  break;
                case TrayCmd.Save:                if (_currentFilePath == null) EnsureWindowVisible();
                                                  MenuItemSave_Click(this, new RoutedEventArgs());                                         break;
                case TrayCmd.SaveAs:              EnsureWindowVisible(); MenuItemSaveAs_Click(this, new RoutedEventArgs());                break;
                case TrayCmd.ExportSelectedFormat: EnsureWindowVisible(); MenuItemExportSelectedFormat_Click(this, new RoutedEventArgs()); break;
                case TrayCmd.Settings:            EnsureWindowVisible(); MenuItemSettings_Click(this, new RoutedEventArgs());              break;

                case TrayCmd.Clear:          MenuItemClear_Click(this, new RoutedEventArgs()); break;
                case TrayCmd.Refresh:        RefreshFormats(); CaptureStaticSnapshot();        break;
                case TrayCmd.MonitorChanges: MenuItemMonitorChanges.IsChecked = !_vm.IsMonitoring;
                                             MenuItemMonitorChanges_Click(this, new RoutedEventArgs()); break;
                case TrayCmd.TrackHistory:   MenuItemTrackHistory.IsChecked = !_vm.IsTrackingHistory;
                                             MenuItemTrackHistory_Click(this, new RoutedEventArgs());   break;

                case TrayCmd.Pause1m:  _vm.ActivatePause(1);  break;
                case TrayCmd.Pause5m:  _vm.ActivatePause(5);  break;
                case TrayCmd.Pause10m: _vm.ActivatePause(10); break;
                case TrayCmd.Pause30m: _vm.ActivatePause(30); break;
                case TrayCmd.Pause1h:  _vm.ActivatePause(60); break;

                case TrayCmd.Resume: _vm.CancelPause(); break;

                case TrayCmd.SubmitFeedback: MenuItemSubmitFeedback_Click(this, new RoutedEventArgs()); break;
                case TrayCmd.About:          EnsureWindowVisible(); _dialogService.ShowAbout(); break;
            }
        }
        finally
        {
            NativeMethods.DestroyMenu(hMenu);
        }
    }

    private void EnsureWindowVisible()
    {
        if (!IsVisible)
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
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

    /// <summary>
    /// Brings the main window to the foreground without toggling.  Shows it if hidden,
    /// restores it if minimized, and preserves a Maximized state if the user had it.
    /// Used by the format-notification balloon click handler.
    /// </summary>
    private void ShowAndActivateMainWindow()
    {
        if (!IsVisible) Show();
        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
        Activate();
    }

    private unsafe void ShowTrayBalloonIfNeeded()
    {
        if (_trayState.BalloonShown || !_trayState.IconAdded || _hwndSource == null)
            return;

        _trayState.BalloonShown = true;
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

    /// <summary>
    /// Called from <see cref="WndProc"/> immediately after a clipboard update has
    /// been refreshed.  Shows a single tray balloon listing every watched format
    /// name present on the new clipboard contents (or does nothing when the
    /// feature is disabled, no patterns match, or the tray icon is unavailable).
    /// </summary>
    private void TryShowFormatNotification()
    {
        if (!_formatNotificationMatcher.HasActivePatterns || _formats.Count == 0)
            return;

        var matched = _formatNotificationMatcher.Match(_formats.Select(f => f.Name));
        if (matched.Count == 0)
            return;

        if (!_trayState.IconAdded || _hwndSource == null)
            return; // Silently skip when the tray icon is missing (e.g. couldn't initialize).

        var ownerLine = _clipboardOwner.Resolve()?.DisplayText ?? "Owner: (unknown)";

        // Sum sizes of the matched formats only — total over all formats can be misleading
        // because formats often share their underlying payload (e.g. text in multiple encodings).
        var matchedSet  = new HashSet<string>(matched, StringComparer.OrdinalIgnoreCase);
        var matchedSize = _formats.Where(f => matchedSet.Contains(f.Name)).Sum(f => f.ContentSizeValue);

        var title = matched.Count == 1
            ? $"Clipboard format: {matched[0]}"
            : $"Clipboard formats matched ({matched.Count})";

        var body = $"{string.Join(", ", matched)}\n{ownerLine}\nSize: {DisplayHelper.FormatFileSize(matchedSize)}";

        ShowFormatNotificationBalloon(title, body);
    }

    private unsafe void ShowFormatNotificationBalloon(string title, string body)
    {
        var nid = new NativeMethods.NOTIFYICONDATA
        {
            cbSize      = (uint)sizeof(NativeMethods.NOTIFYICONDATA),
            hWnd        = _hwndSource!.Handle,
            uID         = 1,
            uFlags      = NativeMethods.NIF_INFO,
            dwInfoFlags = NativeMethods.NIIF_INFO,
        };
        CopyToFixed(nid.szInfoTitle, 64,  title);
        CopyToFixed(nid.szInfo,      256, body);
        NativeMethods.Shell_NotifyIcon(NativeMethods.NIM_MODIFY, ref nid);
    }

    private void ExitApplication()
    {
        _trayState.IsExiting = true;
        DisposeTrayIcon();
        Close();
    }

    // ── Global hotkey ────────────────────────────────────────────────────────

    /// <summary>
    /// Called when the registered global hotkey fires.
    /// Hides the window if it is visible and active; otherwise shows and activates it.
    /// </summary>
    private void OnGlobalHotkey()
    {
        if (IsVisible && IsActive)
        {
            Hide();
        }
        else
        {
            if (!IsVisible) Show();
            WindowState = WindowState.Normal;
            Activate();
        }
    }

    /// <summary>
    /// Attempts to register the hotkey with Windows.
    /// Sets <see cref="_hotkeyRegistered"/> and <see cref="_hotkeyConflict"/> accordingly,
    /// then refreshes the status bar.
    /// </summary>
    private void TryRegisterHotkey()
    {
        if (_hwndSource == null) return;

        _hotkeyRegistered = NativeMethods.RegisterHotKey(
            _hwndSource.Handle, HotkeyId,
            _hotkeyBinding.Modifiers | HotkeyBinding.MOD_NOREPEAT,
            _hotkeyBinding.VirtualKey);

        _hotkeyConflict = !_hotkeyRegistered;
        UpdateHotkeyStatusBar();
    }

    /// <summary>
    /// Unregisters the hotkey if it is currently registered.  Safe to call when not registered.
    /// </summary>
    private void UnregisterHotkey()
    {
        if (!_hotkeyRegistered || _hwndSource == null) return;
        NativeMethods.UnregisterHotKey(_hwndSource.Handle, HotkeyId);
        _hotkeyRegistered = false;
    }

    /// <summary>
    /// Unregisters any existing registration and re-registers using the current
    /// <see cref="_hotkeyEnabled"/> and <see cref="_hotkeyBinding"/> values.
    /// No-op when Minimize to Tray is off or the hotkey is disabled.
    /// </summary>
    private void RefreshHotkeyRegistration()
    {
        UnregisterHotkey();
        if (_trayState.MinimizeToTray && _hotkeyEnabled)
            TryRegisterHotkey();
        else
        {
            _hotkeyConflict = false;
            UpdateHotkeyStatusBar();
        }
    }

    /// <summary>
    /// Updates the hotkey status bar item visibility and conflict indicator.
    /// The item is shown only when Minimize to Tray is on and the hotkey is enabled.
    /// </summary>
    private void UpdateHotkeyStatusBar()
    {
        var show = _trayState.MinimizeToTray && _hotkeyEnabled;
        bool ownerVisible = ClipboardOwnerStatusItem.Visibility == Visibility.Visible;
        bool autoStartVisible = AutoStartPill.Visibility == Visibility.Visible;
        HotkeyStatusItem.Visibility      = show ? Visibility.Visible : Visibility.Collapsed;
        HotkeyStatusSeparator.Visibility = show && (ownerVisible || autoStartVisible) ? Visibility.Visible : Visibility.Collapsed;
        if (show)
        {
            HotkeyStatusText.Text         = $"Hotkey: {_hotkeyBinding}";
            HotkeyConflictIcon.Visibility = _hotkeyConflict ? Visibility.Visible : Visibility.Collapsed;
        }
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
            if (_vm.IsMonitoring && !_vm.IsPaused)
            {
                // Read the sequence number before RefreshFormats() so that any
                // GetClipboardData call (including delayed-rendering triggers) that
                // increments it during the refresh is detectable afterwards.
                var seqAtArrival = _clipboardReader.GetSequenceNumber();
                RefreshFormats();
                if (_vm.IsTrackingHistory)
                    CaptureHistorySession(seqAtArrival);

                // Skip notification when the sequence number advanced — a delayed-rendering
                // provider (e.g. Firefox for HTML) re-posted WM_CLIPBOARDUPDATE during our
                // reads, and the next message will be the final, stable one.  Without this
                // guard the user sees the same balloon twice for a single clipboard set.
                if (_clipboardReader.GetSequenceNumber() == seqAtArrival)
                    TryShowFormatNotification();
            }
            handled = true;
        }
        else if (msg == WM_HOTKEY && (int)wParam == HotkeyId)
        {
            OnGlobalHotkey();
            handled = true;
        }
        else if (msg == WM_TRAYICON)
        {
            const int WM_LBUTTONUP         = 0x0202;
            const int WM_RBUTTONUP         = 0x0205;
            const int NIN_BALLOONUSERCLICK = 0x0405; // WM_USER + 5
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
            else if (notification == NIN_BALLOONUSERCLICK)
            {
                ShowAndActivateMainWindow();
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

        UpdateClipboardOwner();

        // Push the current live clipboard formats onto the [Live Clipboard] sentinel row
        // so its pills, tooltip, and size always reflect reality, even while a real history
        // entry is selected.
        _vm.UpdateLiveRow(_formats);
    }

    private void FormatListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FormatListBox.SelectedItem is not ClipboardFormatItem item)
            return;

        _formatState.FormatId   = item.FormatId;
        _formatState.FormatName = item.Name;
        _formatState.Bytes      = null;

        ContentTabControl.Visibility = Visibility.Visible;
        NoSelectionPanel.Visibility  = Visibility.Collapsed;

        // Resolve bytes (null when data is unavailable; tabs handle null gracefully).
        TryGetFormatBytes(item, out var bytes, out _);
        _formatState.Bytes = bytes;

        foreach (var tab in _previewTabs)
            tab.Update(item.FormatId, item.Name, bytes);

        AutoSelectBestTab();
    }

    /// <summary>
    /// Keeps the currently selected tab if it is still enabled; otherwise switches to the
    /// enabled tab with the lowest <see cref="IPreviewTab.Priority"/> value.
    /// </summary>
    private void AutoSelectBestTab()
    {
        if (ContentTabControl.SelectedItem is TabItem current && current.IsEnabled)
            return;

        var best = _previewTabs
            .Where(t => t.TabItem.IsEnabled)
            .OrderBy(t => t.Priority)
            .FirstOrDefault();

        if (best != null)
            ContentTabControl.SelectedItem = best.TabItem;
    }

    /// <summary>
    /// Resolves the raw byte content for <paramref name="item"/> from the active data source
    /// (history snapshot, static snapshot, or live clipboard).
    /// Returns <see langword="false"/> when the live clipboard could not be opened — the caller
    /// should display <paramref name="hexFailureMessage"/> and return.
    /// </summary>
    private bool TryGetFormatBytes(ClipboardFormatItem item,
        out byte[]? bytes, out string hexFailureMessage)
    {
        hexFailureMessage = "Clipboard data is unavailable.";
        bytes = null;

        if (_historySnapshots != null)
        {
            // History playback mode.
            if (_historySnapshots.TryGetValue(item.Name, out var snapshot))
            {
                bytes = snapshot.Data;
                if (bytes == null && snapshot.HandleType != HandleTypes.None)
                    hexFailureMessage = "No data was captured for this format.";
            }
            else
            {
                hexFailureMessage = "Format data not found in the selected history entry.";
            }
            return true;
        }

        if (_staticSnapshot != null)
        {
            // Monitoring-off mode — bytes come from the snapshot taken when monitoring
            // was disabled (or on the last manual refresh), so the clipboard does not
            // need to be opened and previews stay consistent with the displayed formats.
            if (_staticSnapshot.TryGetValue(item.Name, out var snapshot))
            {
                bytes = snapshot.Data;
                if (bytes == null && snapshot.HandleType != HandleTypes.None)
                    hexFailureMessage = "No data was captured for this format.";
            }
            else
            {
                hexFailureMessage = "Format data not found in the static snapshot.";
            }
            return true;
        }

        // Live mode — read from the clipboard.
        // Use local vars because out-params cannot be captured by lambdas.
        byte[]? liveBytes  = null;
        string  liveMsg    = string.Empty;
        if (!ExecuteWithClipboard(IntPtr.Zero, () =>
        {
            var handleType = _clipboardReader.GetHandleType(item.FormatId);
            _clipboardReader.TryReadFormatBytes(item.FormatId, handleType, out liveBytes, out liveMsg);
        }))
        {
            hexFailureMessage = "Unable to open clipboard.";
            return false;
        }

        bytes             = liveBytes;
        hexFailureMessage = liveMsg;
        return true;
    }

    private void InitializePreviewState()
    {
        foreach (var tab in _previewTabs)
            tab.Reset();
        ContentTabControl.SelectedItem  = _previewTabs.OrderBy(t => t.Priority).First().TabItem;
        ContentTabControl.Visibility    = Visibility.Collapsed;
        NoSelectionPanel.Visibility     = Visibility.Visible;
        _formatState.Bytes      = null;
        _formatState.FormatId   = 0;
        _formatState.FormatName = string.Empty;
    }

    private void UpdateStatusBar()
    {
        _vm.UpdateStatusBarText();  // computes StatusBarText → PropertyChanged → StatusText.Text
        AutoStartPill.Visibility = _startAtLogin ? Visibility.Visible : Visibility.Collapsed;
        UpdateOwnerPillSeparator();
        UpdateHotkeyStatusBar();
    }


    private void UpdateFileStatusBar(string action, string path)
    {
        FileStatusText.Text = $"{action}: {Path.GetFileName(path)}";
        FileStatusSeparator.Visibility = Visibility.Visible;
    }

    // ── Clipboard owner ──────────────────────────────────────────────────────

    /// <summary>
    /// Resolves the current clipboard owner and updates the status bar item.
    /// Called every time the live clipboard is refreshed.
    /// </summary>
    private void UpdateClipboardOwner()
    {
        var info = _clipboardOwner.Resolve();
        if (info == null)
        {
            ClipboardOwnerStatusItem.Visibility = Visibility.Collapsed;
            ClipboardOwnerTextBlock.Text        = string.Empty;
            ClipboardOwnerTextBlock.ToolTip     = null;
        }
        else
        {
            ClipboardOwnerTextBlock.Text    = info.DisplayText;
            ClipboardOwnerTextBlock.ToolTip = info.TooltipText is { } tip
                ? new TextBlock { Text = tip, TextWrapping = TextWrapping.Wrap, MaxWidth = 650 }
                : null;
            ClipboardOwnerStatusItem.Visibility = Visibility.Visible;
        }
        UpdateOwnerPillSeparator();
        UpdateHotkeyStatusBar();
    }

    /// <summary>
    /// Shows the separator between the owner item and the AUTO-START pill only when
    /// both are visible at the same time.
    /// </summary>
    private void UpdateOwnerPillSeparator()
    {
        bool ownerVisible    = ClipboardOwnerStatusItem.Visibility == Visibility.Visible;
        bool autoStartVisible = AutoStartPill.Visibility            == Visibility.Visible;
        OwnerAutoStartSeparator.Visibility =
            (ownerVisible && autoStartVisible) ? Visibility.Visible : Visibility.Collapsed;
    }

    private void RefreshClipboardCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        RefreshFormats();
        CaptureStaticSnapshot();
    }

    private void RefreshClipboardCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        e.CanExecute = !_vm.IsMonitoring;
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
        e.CanExecute = _formatState.Bytes is { Length: > 0 };
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
        MenuItemTrackHistory.IsEnabled = _vm.IsMonitoring;
        MenuItemPause.IsEnabled        = _vm.IsMonitoring;
        MenuItemResume.IsEnabled       = _vm.IsPaused;
    }

    private void MenuItemClear_Click(object sender, RoutedEventArgs e)
    {
        Clipboard.Clear();
        if (!_vm.IsMonitoring)
        {
            RefreshFormats();
            CaptureStaticSnapshot();
        }
    }

    private void MenuItemMonitorChanges_Click(object sender, RoutedEventArgs e)
    {
        // IsMonitoring already updated by the TwoWay binding before this handler fires.
        if (!_vm.IsMonitoring)
            _vm.OnMonitoringDisabled();

        UpdateStatusBar();
        SavePreferences();

        if (_vm.IsMonitoring)
        {
            _staticSnapshot = null;
            RefreshFormats();
        }
        else
        {
            CaptureStaticSnapshot();
        }
    }

    // ── Clipboard pause ──────────────────────────────────────────────────────

    private void PauseMenu_SubmenuOpened(object sender, RoutedEventArgs e)
    {
        MenuItemPause1m.Header  = _vm.IsPaused ? "+1 minute"   : "1 minute";
        MenuItemPause5m.Header  = _vm.IsPaused ? "+5 minutes"  : "5 minutes";
        MenuItemPause10m.Header = _vm.IsPaused ? "+10 minutes" : "10 minutes";
        MenuItemPause30m.Header = _vm.IsPaused ? "+30 minutes" : "30 minutes";
        MenuItemPause1h.Header  = _vm.IsPaused ? "+1 hour"     : "1 hour";
    }

    private void MenuItemPause_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem item && item.Tag is string tagStr && int.TryParse(tagStr, out var minutes))
            _vm.ActivatePause(minutes);
    }

    private void MenuItemResume_Click(object sender, RoutedEventArgs e)
    {
        _vm.CancelPause();
    }

    private void MenuItemSubmitFeedback_Click(object sender, RoutedEventArgs e)
    {
        ShellHelper.OpenUrl("https://github.com/dbakuntsev/simply.clipboard-monitor/issues");
    }

    private void MenuItemSettings_Click(object sender, RoutedEventArgs e)
    {
        var result = _dialogService.ShowSettings(new SettingsDialogInput(
            _vm.MaxEntries, _vm.MaxSizeMb,
            _trayState.MinimizeToTray, _startAtLogin, _startMinimized,
            _hotkeyEnabled, _hotkeyBinding, _hotkeyConflict,
            _formatNotificationsEnabled, _formatNotificationPatterns));

        if (result.HistoryWasCleared)
            _vm.OnHistoryCleared();

        if (!result.Saved)
            return;

        _vm.MaxEntries = result.MaxEntries;
        _vm.MaxSizeMb  = result.MaxSizeMb;

        // Update hotkey state first so ApplyTraySettings reads the new values.
        var hotkeySettingsChanged = result.HotkeyEnabled != _hotkeyEnabled ||
                                    result.HotkeyBinding != _hotkeyBinding;
        if (hotkeySettingsChanged)
        {
            _hotkeyEnabled = result.HotkeyEnabled;
            _hotkeyBinding = result.HotkeyBinding;
        }

        // Update format-notification state before tray adjustments so that
        // RecomputeTrayIconPresence sees the new value.
        _formatNotificationsEnabled = result.FormatNotificationsEnabled;
        _formatNotificationPatterns = result.FormatNotificationPatterns;
        _formatNotificationMatcher.Configure(_formatNotificationsEnabled, _formatNotificationPatterns);

        if (result.MinimizeToSystemTray != _trayState.MinimizeToTray)
            ApplyTraySettings(result.MinimizeToSystemTray);
        else if (hotkeySettingsChanged)
            RefreshHotkeyRegistration();

        // Format-notification toggling may change the tray-icon requirement
        // even when MinimizeToSystemTray didn't change.
        RecomputeTrayIconPresence();

        if (result.StartAtLogin != _startAtLogin)
            ApplyAutoStartPreference(result.StartAtLogin);

        _startMinimized = result.StartMinimized;

        SavePreferences();

        // Enforce new limits against any existing database, regardless of whether
        // history tracking is currently on (old data may still need trimming).
        _vm.EnqueueEnforceLimits();
    }

    private void MenuItemAbout_Click(object sender, RoutedEventArgs e)
        => _dialogService.ShowAbout();

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
        if (_formatState.Bytes is not { Length: > 0 })
            return;

        var textTab  = _previewTabs.OfType<TextPreviewControl>().FirstOrDefault();
        var imageTab = _previewTabs.OfType<ImagePreviewControl>().FirstOrDefault();

        Encoding? manuallySelected = textTab?.ManuallyChangedEncoding is { } enc
            && ContentTabControl.SelectedItem == textTab.TabItem
            ? enc : null;

        var ctx = new FormatExportContext(
            Bytes:                    _formatState.Bytes,
            FormatId:                 _formatState.FormatId,
            FormatName:               _formatState.FormatName,
            AutoDetectedEncoding:     textTab?.AutoDetectedEncoding,
            ManuallySelectedEncoding: manuallySelected,
            ImagePreviewSource:       imageTab?.PreviewImageSource);

        var applicable = _formatExporters.Where(exp => exp.CanExport(ctx)).ToList();
        if (applicable.Count == 0)
            return;

        // Default: text > jpeg/png (by format name hint) > binary
        int defaultIndex = applicable.FindIndex(exp => exp.Extension == ".txt");
        if (defaultIndex < 0)
        {
            var norm          = _formatState.FormatName.ToLowerInvariant();
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
            ShowError($"Failed to export:\n\n{ex.Message}", "Export Failed");
        }
    }

    // ── Save / Load implementation ───────────────────────────────────────────

    /// <summary>
    /// Opens the clipboard, runs <paramref name="action"/>, then closes it.
    /// Returns <see langword="false"/> (without running the action) if the clipboard
    /// could not be opened.
    /// </summary>
    private bool ExecuteWithClipboard(IntPtr hwnd, Action action)
    {
        if (!_clipboardReader.TryOpenClipboard(hwnd))
            return false;
        try { action(); }
        finally { _clipboardReader.CloseClipboard(); }
        return true;
    }

    private void SaveToFile(string path)
    {
        if (_formats.Count == 0)
        {
            ShowInfo(
                "There are no clipboard formats to save.\n\nCopy something to the clipboard first, then save.",
                "Nothing to Save");
            return;
        }

        var formats = new List<SavedClipboardFormat>(_formats.Count);
        if (!ExecuteWithClipboard(IntPtr.Zero, () =>
        {
            foreach (var item in _formats)
            {
                var handleType = _clipboardReader.GetHandleType(item.FormatId);
                byte[]? data   = null;
                if (handleType != HandleTypes.None)
                    _clipboardReader.TryReadFormatBytes(item.FormatId, handleType, out data, out _);

                formats.Add(new SavedClipboardFormat(item.Ordinal, item.FormatId, item.Name, handleType, data));
            }
        }))
        {
            ShowWarning(
                "Unable to open the clipboard for reading.\n\nClose any application that may be using the clipboard and try again.",
                "Save Failed");
            return;
        }

        _clipboardFiles.Save(path, formats);
        _currentFilePath = path;
        UpdateFileStatusBar("Saved", path);
    }

    private void LoadFromFile(string path)
    {
        var formats = _clipboardFiles.Load(path);

        if (formats.Count == 0)
        {
            ShowInfo(
                "The clipboard database is empty — no formats were stored in the file.",
                "Nothing to Load");
            return;
        }

        if (!ExecuteWithClipboard(_hwndSource?.Handle ?? IntPtr.Zero, () =>
        {
            NativeMethods.EmptyClipboard();
            _clipboardWriter.RestoreFormats(formats);
        }))
        {
            ShowWarning(
                "Unable to open the clipboard for writing.\n\nClose any application that may be using the clipboard and try again.",
                "Load Failed");
            return;
        }

        _currentFilePath = path;
        UpdateFileStatusBar("Loaded", path);

        if (!_vm.IsMonitoring)
            RefreshFormats();
    }

    // ── Error reporting ─────────────────────────────────────────────────────

    private void ShowInfo(string message, string title) =>
        MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Information);

    private void ShowWarning(string message, string title) =>
        MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Warning);

    private void ShowError(string message, string title) =>
        MessageBox.Show(this, message, title, MessageBoxButton.OK, MessageBoxImage.Error);

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

        ShowError(text, "Error");
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

            _vm.IsMonitoring      = preferences.MonitorChanges;
            _vm.IsTrackingHistory = preferences.TrackHistory && preferences.MonitorChanges;
            _vm.MaxEntries        = preferences.HistoryMaxEntries > 0 ? preferences.HistoryMaxEntries : _vm.MaxEntries;
            _vm.MaxSizeMb         = preferences.HistoryMaxSizeMb  > 0 ? preferences.HistoryMaxSizeMb  : _vm.MaxSizeMb;
            _trayState.MinimizeToTray = preferences.MinimizeToSystemTray;
            _trayState.BalloonShown   = preferences.TrayBalloonShown;
            _startMinimized           = preferences.StartMinimized;
            _hotkeyEnabled            = preferences.HotkeyEnabled;
            _hotkeyBinding            = HotkeyBinding.TryParse(preferences.HotkeyBinding, out var hb)
                                            ? hb : HotkeyBinding.Default;
            _formatNotificationsEnabled = preferences.FormatNotificationsEnabled;
            _formatNotificationPatterns = preferences.FormatNotificationPatterns;
            _formatNotificationMatcher.Configure(_formatNotificationsEnabled, _formatNotificationPatterns);
        }
        catch
        {
            _currentSortProperty     = DefaultSortProperty;
            _currentSortDirection    = ListSortDirection.Ascending;
            _storedColumnPreferences = [];
            _vm.IsMonitoring      = true;
            _vm.IsTrackingHistory = false;
        }

        // Always read auto-start from the registry so the UI reflects the actual system state,
        // even if the entry was manually added or removed outside the app.
        _startAtLogin = AutoStartHelper.IsAutoStartEnabled();

        // Apply and subscribe to the word wrap preference for the text preview tab.
        var textTab = _previewTabs.OfType<TextPreviewControl>().FirstOrDefault();
        if (textTab != null)
        {
            textTab.WordWrap = preferences.TextWordWrap;
            textTab.WordWrapChanged += (_, _) => SavePreferences();
        }
    }

    private void SavePreferences()
    {
        var preferences = new UserPreferences
        {
            SortProperty         = _currentSortProperty,
            SortDirection        = _currentSortDirection.ToString(),
            FormatColumns        = CaptureFormatColumnPreferences(),
            MonitorChanges       = _vm.IsMonitoring,
            TrackHistory         = _vm.IsTrackingHistory,
            HistoryMaxEntries    = _vm.MaxEntries,
            HistoryMaxSizeMb     = _vm.MaxSizeMb,
            MinimizeToSystemTray = _trayState.MinimizeToTray,
            TrayBalloonShown     = _trayState.BalloonShown,
            StartAtLogin         = _startAtLogin,
            StartMinimized       = _startMinimized,
            TextWordWrap         = _previewTabs.OfType<TextPreviewControl>().FirstOrDefault()?.WordWrap ?? false,
            HotkeyEnabled        = _hotkeyEnabled,
            HotkeyBinding        = _hotkeyBinding.ToString(),
            FormatNotificationsEnabled = _formatNotificationsEnabled,
            FormatNotificationPatterns = _formatNotificationPatterns,
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
        // IsTrackingHistory already updated by the TwoWay binding before this handler fires.
        if (_vm.IsTrackingHistory)
        {
            ShowHistoryPanel();
            ShowHistoryLoadingOverlay();
            // Refresh the sentinel row from the current _formats so [Live Clipboard]
            // reflects the actual clipboard the moment the panel becomes visible — even
            // before the first WM_CLIPBOARDUPDATE after enabling tracking arrives.
            _vm.UpdateLiveRow(_formats);
            // Defer loading until integrity check and migration complete on the channel.
            _vm.EnqueueDatabaseInitialization(
                isTrackingHistory: true,
                onReady: () => Dispatcher.Invoke(() =>
                {
                    HideHistoryLoadingOverlay();
                    CaptureStartupSession();
                    _vm.LoadHistoryFromDatabase(selectLive: true);
                }));
        }
        else
        {
            _vm.OnHistoryTrackingDisabled();
        }

        UpdateStatusBar();
        SavePreferences();
    }

    private void HistoryListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not MainWindowViewModel.HistoryItem entry)
        {
            _vm.SelectedSession = null;
            return;
        }
        _vm.SelectedSession = entry;

        if (entry.IsLive)
        {
            // [Live Clipboard] sentinel: show the actual clipboard, not a stored snapshot.
            // Only re-read when we're currently displaying snapshot data — otherwise the live
            // view is already on screen (e.g. just after a write, RefreshFormats already ran).
            if (_historySnapshots != null)
            {
                _historySnapshots = null;
                RefreshFormats();
            }
            return;
        }

        List<FormatSnapshot> snapshots;
        try
        {
            snapshots = _history.LoadSessionFormats(entry.SessionId);
        }
        catch
        {
            return;
        }

        var previousFormatName = (FormatListBox.SelectedItem as ClipboardFormatItem)?.Name;
        var previousTab        = ContentTabControl.SelectedItem;

        _historySnapshots = snapshots.ToDictionary(s => s.FormatName, StringComparer.Ordinal);
        InitializePreviewState();
        _formats.Clear();

        var filterTerm = string.IsNullOrWhiteSpace(_vm.SearchText) ? null : _vm.SearchText.Trim();
        foreach (var snap in snapshots)
        {
            var contentSize      = snap.OriginalSize > 0 ? snap.OriginalSize.ToString("N0") : "n/a";
            var contentSizeValue = snap.OriginalSize > 0 ? snap.OriginalSize : -1L;
            _formats.Add(new ClipboardFormatItem(snap.Ordinal, snap.FormatId, snap.FormatName, contentSize, contentSizeValue)
            {
                IsHighlighted = filterTerm != null &&
                                (snap.FormatName.Contains(filterTerm, StringComparison.OrdinalIgnoreCase) ||
                                 (_formatClassifier.GetFormatPillLabel(snap.FormatId, snap.FormatName)
                                      ?.Contains(filterTerm, StringComparison.OrdinalIgnoreCase) ?? false))
            });
        }

        if (previousFormatName != null)
        {
            var matching = _formats.FirstOrDefault(f => f.Name == previousFormatName);
            if (matching != null)
            {
                // Restore the tab before triggering FormatListBox_OnSelectionChanged so it
                // can fall back to Hex only when the restored tab is incompatible with the format.
                if (previousTab != null)
                    ContentTabControl.SelectedItem = previousTab;
                FormatListBox.SelectedItem = matching;
            }
        }
    }

    private void HistoryFilterClearButton_Click(object sender, RoutedEventArgs e)
    {
        HistoryFilterBox.Clear();
        HistoryFilterBox.Focus();
    }

    private void HistoryFilterBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _vm.SearchText = HistoryFilterBox.Text;

        _filterCts?.Cancel();
        _filterCts = new CancellationTokenSource();
        var token = _filterCts.Token;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(200, token);
                await Dispatcher.InvokeAsync(() =>
                {
                    if (!token.IsCancellationRequested)
                    {
                        var filter = string.IsNullOrWhiteSpace(_vm.SearchText) ? null : _vm.SearchText;
                        if (filter == null)
                            // Filter removed: show all history and select the [Live Clipboard] sentinel.
                            _vm.LoadHistoryFromDatabase(null, selectLive: true);
                        else
                            // Filter changed: keep the current selection if still visible (falls
                            // back to the sentinel when the previously selected entry is filtered out).
                            _vm.LoadHistoryFromDatabase(filter, _vm.SelectedSession?.SessionId);
                    }
                });
            }
            catch (OperationCanceledException) { }
        });
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

        _vm.EnqueueHistorySession(snapshots, timestamp);
    }

    /// <summary>
    /// Called once on startup (after the initial history load) to snapshot the current
    /// clipboard and record it as a new session if it differs from the last saved entry.
    /// Runs the DB work through the history channel so it is ordered after
    /// <see cref="IHistoryMaintenance.MigrateSchema"/>.
    /// </summary>
    private void CaptureStartupSession()
    {
        if (_formats.Count == 0)
            return;

        var timestamp = DateTime.Now;
        var snapshots = _clipboardReader.CaptureAllFormats(_formats.ToList());
        if (snapshots.Count == 0)
            return;

        _vm.EnqueueStartupSession(snapshots, timestamp);
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

    // ── History context menu ─────────────────────────────────────────────────

    private void HistoryContextMenu_Opening(object sender, ContextMenuEventArgs e)
    {
        var selection    = HistoryListView.SelectedItem as MainWindowViewModel.HistoryItem;
        var hasSelection = selection != null;
        var isLive       = selection?.IsLive ?? false;

        // Load: only meaningful for a real entry whose data isn't already on the clipboard.
        HistoryMenuLoadIntoClipboard.IsEnabled = hasSelection && !isLive
            && _historySnapshots != null
            && !IsHistorySessionCurrentClipboard(_historySnapshots);

        // Save As: real entry → save snapshot blobs.  Sentinel → save the live clipboard
        // (handled by HistorySaveAs_Click branching).  Disabled on the sentinel when the
        // clipboard is empty since there's nothing to save.
        HistoryMenuSaveAs.IsEnabled = hasSelection && (!isLive || _formats.Count > 0);

        // Delete: only valid for real entries — the sentinel is not persisted.
        HistoryMenuDelete.IsEnabled = hasSelection && !isLive;
    }

    /// <summary>
    /// Returns true when the current clipboard content matches the given session snapshots
    /// (same format names and identical per-format byte content).
    /// </summary>
    private bool IsHistorySessionCurrentClipboard(IReadOnlyDictionary<string, FormatSnapshot> sessionSnapshots)
    {
        var currentFormats = _clipboardReader.EnumerateFormats();
        if (currentFormats.Count != sessionSnapshots.Count)
            return false;

        var currentSnapshots = _clipboardReader.CaptureAllFormats(currentFormats);
        if (currentSnapshots.Count != sessionSnapshots.Count)
            return false;

        foreach (var cur in currentSnapshots)
        {
            if (!sessionSnapshots.TryGetValue(cur.FormatName, out var hist))
                return false;
            if ((cur.Data == null) != (hist.Data == null))
                return false;
            if (cur.Data != null && !cur.Data.AsSpan().SequenceEqual(hist.Data!.AsSpan()))
                return false;
        }
        return true;
    }

    private void HistoryListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var item = FindVisualAncestor<ListViewItem>(e.OriginalSource as DependencyObject);
        if (item?.DataContext is MainWindowViewModel.HistoryItem historyItem)
        {
            _historyDragStartPoint = e.GetPosition(HistoryListView);
            HistoryListView.SelectedItem = historyItem;
        }
        else
        {
            _historyDragStartPoint = null;
        }
    }

    private void HistoryListView_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _historyDragStartPoint == null)
            return;

        var current = e.GetPosition(HistoryListView);
        if (Math.Abs(current.X - _historyDragStartPoint.Value.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(current.Y - _historyDragStartPoint.Value.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        _historyDragStartPoint = null;

        if (HistoryListView.SelectedItem is not MainWindowViewModel.HistoryItem entry)
            return;

        List<FormatSnapshot> snapshots;
        try   { snapshots = _history.LoadSessionFormats(entry.SessionId); }
        catch { return; }

        var dataObj = TryBuildDragDataObject(snapshots);
        if (dataObj == null) return;

        // Suppress ListView selection-drag and release WPF capture before DoDragDrop.
        e.Handled = true;
        Mouse.Capture(null);
        DragDrop.DoDragDrop(HistoryListView, dataObj, DragDropEffects.Copy);
    }

    /// <summary>
    /// Builds a WPF DataObject from the given history snapshots, mapping each recognised
    /// clipboard format to the corresponding WPF DataFormats entry.
    /// Returns null when no supported format has any data.
    /// </summary>
    private static System.Windows.DataObject? TryBuildDragDataObject(List<FormatSnapshot> snapshots)
    {
        var obj = new System.Windows.DataObject();
        bool hasAny = false, hasText = false, hasHtml = false, hasRtf = false, hasFiles = false;
        bool hasDib = false, hasBitmapSource = false, hasPng = false;

        foreach (var snap in snapshots)
        {
            if (snap.Data is not { Length: > 0 }) continue;

            // ── Text (CF_UNICODETEXT preferred; fall back to CF_TEXT / CF_OEMTEXT) ─
            if (!hasText)
            {
                string? text = null;
                if (snap.FormatId == CF_UNICODETEXT)
                    text = Encoding.Unicode.GetString(snap.Data).TrimEnd('\0');
                else if (snap.FormatId == CF_TEXT)
                    text = Encoding.Default.GetString(snap.Data).TrimEnd('\0');
                else if (snap.FormatId == CF_OEMTEXT)
                    text = Encoding.GetEncoding((int)NativeMethods.GetOEMCP()).GetString(snap.Data).TrimEnd('\0');

                if (text is { Length: > 0 })
                {
                    // DataFormats.StringFormat ("System.String") is required to activate
                    // WPF's synonym mapping so that CF_UNICODETEXT and CF_TEXT are both
                    // enumerated in the native OLE IDataObject.
                    obj.SetData(DataFormats.StringFormat, text);
                    obj.SetData(DataFormats.UnicodeText,  text);
                    obj.SetData(DataFormats.Text,         text);
                    hasText = hasAny = true;
                }
            }

            // ── HTML ────────────────────────────────────────────────────────
            if (!hasHtml && IsHtmlFormat(snap.FormatName))
            {
                var html = DecodeHtmlBytes(snap.FormatName, snap.Data);
                if (html.Length > 0)
                { obj.SetData(DataFormats.Html, html); hasHtml = hasAny = true; }
            }

            // ── RTF ─────────────────────────────────────────────────────────
            if (!hasRtf && IsRtfFormat(snap.FormatName))
            {
                var rtf = Encoding.ASCII.GetString(snap.Data).TrimEnd('\0');
                if (rtf.Length > 0)
                { obj.SetData(DataFormats.Rtf, rtf); hasRtf = hasAny = true; }
            }

            // ── File drop (CF_HDROP) ────────────────────────────────────────
            if (!hasFiles && snap.FormatId == CF_HDROP)
            {
                var files = ParseDropFiles(snap.Data);
                if (files.Length > 0)
                { obj.SetData(DataFormats.FileDrop, files); hasFiles = hasAny = true; }
            }

            // ── CF_DIB / CF_DIBV5: raw HGLOBAL bytes as MemoryStream ────────
            // Most native apps (Word, Paint, etc.) accept CF_DIB more reliably than CF_BITMAP.
            if (!hasDib && snap.FormatId is CF_DIB or CF_DIBV5)
            {
                obj.SetData(DataFormats.Dib, new MemoryStream(snap.Data));
                hasDib = hasAny = true;
                // Also produce a frozen BitmapSource for WPF/modern apps (CF_BITMAP).
                if (!hasBitmapSource)
                {
                    var bmp = TryDecodeDibToBitmapSource(snap.Data);
                    if (bmp != null)
                    { obj.SetData(DataFormats.Bitmap, bmp); hasBitmapSource = true; }
                }
            }

            // ── CF_BITMAP (stored as DIB bytes via GetDIBits) ───────────────
            if (!hasBitmapSource && HBitmapFormats.Contains(snap.FormatId))
            {
                var bmp = TryDecodeDibToBitmapSource(snap.Data);
                if (bmp != null)
                {
                    obj.SetData(DataFormats.Bitmap, bmp);
                    hasBitmapSource = hasAny = true;
                    // Also expose as CF_DIB for native apps that prefer it.
                    if (!hasDib)
                    { obj.SetData(DataFormats.Dib, new MemoryStream(snap.Data)); hasDib = true; }
                }
            }

            // ── PNG ─────────────────────────────────────────────────────────
            // Apps like Office and Chrome can consume a "PNG" custom format directly.
            if (!hasPng && (string.Equals(snap.FormatName, "PNG",       StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(snap.FormatName, "image/png", StringComparison.OrdinalIgnoreCase)))
            {
                obj.SetData(snap.FormatName, new MemoryStream(snap.Data));
                hasPng = hasAny = true;
            }
        }

        return hasAny ? obj : null;
    }

    /// <summary>
    /// Parses a CF_HDROP HGLOBAL byte array (DROPFILES header + file list) into
    /// an array of file paths.
    /// </summary>
    private static string[] ParseDropFiles(byte[] data)
    {
        // DROPFILES: pFiles(4) + POINT(8) + fNC(4) + fWide(4) = 20 bytes minimum
        if (data.Length < 20) return [];
        int  pFiles = BitConverter.ToInt32(data, 0);
        bool fWide  = BitConverter.ToInt32(data, 16) != 0;
        if (pFiles < 0 || pFiles >= data.Length) return [];

        var files = new List<string>();
        if (fWide)
        {
            // Null-terminated UTF-16 strings; list ends at double-null
            for (int i = pFiles; i + 1 < data.Length; )
            {
                int start = i;
                while (i + 1 < data.Length && (data[i] | data[i + 1]) != 0) i += 2;
                if (i == start) break;  // double-null terminator
                files.Add(Encoding.Unicode.GetString(data, start, i - start));
                i += 2;
            }
        }
        else
        {
            // Null-terminated ANSI strings; list ends at double-null
            for (int i = pFiles; i < data.Length; )
            {
                int start = i;
                while (i < data.Length && data[i] != 0) i++;
                if (i == start) break;
                files.Add(Encoding.Default.GetString(data, start, i - start));
                i++;
            }
        }
        return [.. files];
    }

    /// <summary>
    /// Decodes a raw DIB byte array (BITMAPINFOHEADER + palette + pixels, as stored for
    /// CF_DIB / CF_DIBV5 / CF_BITMAP) into a <see cref="BitmapSource"/>.
    /// Returns null on any failure.
    /// </summary>
    private static BitmapSource? TryDecodeDibToBitmapSource(byte[] data)
    {
        try
        {
            if (data.Length < 40) return null;  // minimum BITMAPINFOHEADER size
            int biSize     = BitConverter.ToInt32(data,  0);
            int biWidth    = BitConverter.ToInt32(data,  4);
            int biBitCount = BitConverter.ToInt16(data, 14);
            int biCompr    = BitConverter.ToInt32(data, 16);
            int biClrUsed  = BitConverter.ToInt32(data, 32);
            if (biSize < 40 || biWidth <= 0) return null;

            // BI_BITFIELDS (3) stores 3 DWORD colour masks instead of a palette.
            int paletteCount = biBitCount <= 8
                ? (biClrUsed > 0 ? biClrUsed : 1 << biBitCount)
                : biCompr == 3 ? 3 : 0;

            // Prepend BITMAPFILEHEADER (14 bytes) to make a valid .bmp stream.
            int pixelOffset = 14 + biSize + paletteCount * 4;
            var bmpFile     = new byte[14 + data.Length];
            bmpFile[0] = (byte)'B'; bmpFile[1] = (byte)'M';
            BitConverter.GetBytes(bmpFile.Length).CopyTo(bmpFile,  2);  // bfSize
            BitConverter.GetBytes(pixelOffset).CopyTo(bmpFile,    10);  // bfOffBits
            data.CopyTo(bmpFile, 14);

            using var ms  = new MemoryStream(bmpFile);
            var decoder   = BitmapDecoder.Create(ms, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            if (decoder.Frames.Count == 0) return null;
            var frame = decoder.Frames[0];
            frame.Freeze();  // required: DragDrop marshals the BitmapSource on a different thread
            return frame;
        }
        catch { return null; }
    }

    /// <summary>
    /// Decodes raw HTML clipboard bytes to a string, handling UTF-8, UTF-16 LE/BE (with or
    /// without BOM).  "text/html" from Chromium browsers arrives as UTF-16 LE without a BOM.
    /// </summary>
    private static string DecodeHtmlBytes(string formatName, byte[] data)
    {
        // Explicit BOM takes priority over format name.
        if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE)
            return Encoding.Unicode.GetString(data, 2, data.Length - 2).TrimEnd('\0');
        if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF)
            return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2).TrimEnd('\0');
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            return Encoding.UTF8.GetString(data, 3, data.Length - 3).TrimEnd('\0');

        // "HTML Format" (Windows CF_HTML) is always UTF-8 per spec.
        if (string.Equals(formatName, "HTML Format", StringComparison.OrdinalIgnoreCase))
            return Encoding.UTF8.GetString(data).TrimEnd('\0');

        // "text/html" from Chromium is UTF-16 LE without a BOM.
        // Sniff: in UTF-16 LE, the high byte of every ASCII character is 0x00.
        // Sample the first few code-unit pairs; a majority of 0x00 high bytes is conclusive.
        if (data.Length >= 4)
        {
            int limit = Math.Min(data.Length / 2, 8);
            int nullsAtOdd = 0;
            for (int i = 0; i < limit; i++)
                if (data[2 * i + 1] == 0x00) nullsAtOdd++;
            if (nullsAtOdd > limit / 2)
                return Encoding.Unicode.GetString(data).TrimEnd('\0');
        }

        return Encoding.UTF8.GetString(data).TrimEnd('\0');
    }

    private static T? FindVisualAncestor<T>(DependencyObject? obj) where T : DependencyObject
    {
        while (obj != null)
        {
            if (obj is T t) return t;
            obj = VisualTreeHelper.GetParent(obj);
        }
        return null;
    }

    private void HistoryLoadIntoClipboard_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not MainWindowViewModel.HistoryItem entry)
            return;
        if (entry.IsLive)
            return;  // [Live Clipboard] IS the clipboard — nothing to load.

        List<FormatSnapshot> snapshots;
        try
        {
            snapshots = _history.LoadSessionFormats(entry.SessionId);
        }
        catch
        {
            return;
        }

        if (snapshots.Count == 0)
            return;

        var formats = snapshots
            .Select(s => new SavedClipboardFormat(s.Ordinal, s.FormatId, s.FormatName, s.HandleType, s.Data))
            .ToList();

        if (!ExecuteWithClipboard(_hwndSource?.Handle ?? IntPtr.Zero, () =>
        {
            NativeMethods.EmptyClipboard();
            _clipboardWriter.RestoreFormats(formats);
        }))
        {
            ShowWarning(
                "Unable to open the clipboard for writing.\n\nClose any application that may be using the clipboard and try again.",
                "Load Failed");
            return;
        }

        if (!_vm.IsMonitoring)
            RefreshFormats();
    }

    private void HistorySaveAs_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not MainWindowViewModel.HistoryItem entry)
            return;

        if (entry.IsLive)
        {
            // [Live Clipboard]: save the live clipboard, not a stored session.  Reuses the
            // existing File-menu "Save As" path (dialog + SaveToFile over _formats).
            MenuItemSaveAs_Click(sender, e);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title           = "Save Clipboard Database",
            Filter          = "Clipboard Database (*.clipdb)|*.clipdb",
            DefaultExt      = "clipdb",
            OverwritePrompt = true,
        };

        if (dlg.ShowDialog(this) != true)
            return;

        try
        {
            var snapshots = _history.LoadSessionFormats(entry.SessionId);
            if (snapshots.Count == 0)
            {
                ShowInfo("The selected history entry has no formats to save.", "Nothing to Save");
                return;
            }

            var formats = snapshots
                .Select(s => new SavedClipboardFormat(s.Ordinal, s.FormatId, s.FormatName, s.HandleType, s.Data))
                .ToList();
            _clipboardFiles.Save(dlg.FileName, formats);
        }
        catch (Exception ex)
        {
            ShowFileOperationError("save", ex);
        }
    }

    private void HistoryDelete_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListView.SelectedItem is not MainWindowViewModel.HistoryItem entry)
            return;
        if (entry.IsLive)
            return;  // sentinel is not persisted — nothing to delete.

        var confirm = MessageBox.Show(
            this,
            $"Delete the clipboard history entry from {entry.DateText}?",
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes)
            return;

        var filter = string.IsNullOrWhiteSpace(_vm.SearchText) ? null : _vm.SearchText;
        _vm.EnqueueDeleteSession(entry.SessionId, filter);
    }

    private void HistoryClearAll_Click(object sender, RoutedEventArgs e)
    {
        var confirm = MessageBox.Show(
            this,
            "Clear all clipboard history? This cannot be undone.",
            "Confirm Clear All",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes)
            return;

        _vm.EnqueueClearAll();
    }

    private void ShowHistoryPanel()
    {
        LeftPanelGrid.RowDefinitions[0].Height = new GridLength(300);
        LeftPanelGrid.RowDefinitions[1].Height = new GridLength(5);
        LeftPanelGrid.RowDefinitions[2].Height = new GridLength(1, GridUnitType.Star);
    }

    private void HideHistoryPanel()
    {
        LeftPanelGrid.RowDefinitions[0].Height = new GridLength(1, GridUnitType.Star);
        LeftPanelGrid.RowDefinitions[1].Height = new GridLength(0);
        LeftPanelGrid.RowDefinitions[2].Height = new GridLength(0);
    }

    private void ShowHistoryLoadingOverlay()  => HistoryLoadingOverlay.Visibility = Visibility.Visible;
    private void HideHistoryLoadingOverlay()  => HistoryLoadingOverlay.Visibility = Visibility.Collapsed;
}
