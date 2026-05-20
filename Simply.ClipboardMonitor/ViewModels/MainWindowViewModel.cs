using System.Collections.ObjectModel;
using System.Threading.Channels;
using System.Windows.Threading;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Common.Mvvm;
using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;

namespace Simply.ClipboardMonitor.ViewModels;

/// <summary>
/// ViewModel for <see cref="Simply.ClipboardMonitor.MainWindow"/>.
/// Owns all clipboard-monitoring state, history-channel orchestration, and
/// database-lifecycle management so the code-behind retains only view concerns
/// (WndProc, tray icon, file dialogs, column sorting).
/// </summary>
internal sealed class MainWindowViewModel : ObservableObject
{
    // ── Inner data models ────────────────────────────────────────────────────

    /// <summary>
    /// Represents one persisted clipboard session shown in the history panel,
    /// or the synthetic [Live Clipboard] sentinel row when <see cref="IsLive"/> is <c>true</c>.
    /// </summary>
    internal sealed class HistoryItem : ObservableObject
    {
        /// <summary>SessionId reserved for the synthetic [Live Clipboard] sentinel row.</summary>
        public const long LiveSessionId = -1L;

        public required long     SessionId      { get; init; }
        public required DateTime Timestamp      { get; init; }
        public required string   FormatsText    { get; init; }
        public required long     TotalSize      { get; init; }
        public required IReadOnlyList<(uint FormatId, string FormatName)> Formats { get; init; }

        /// <summary>
        /// True for the synthetic [Live Clipboard] sentinel row pinned at the top of the
        /// history list whenever history tracking is enabled.  Sentinel rows are not persisted.
        /// </summary>
        public bool IsLive { get; init; }

        private IReadOnlyList<FormatPill> _pills = [];
        /// <summary>
        /// Pills shown in the Formats column. Mutable on the sentinel row so its pills can
        /// be refreshed on every clipboard change; written once on real entries.
        /// </summary>
        public IReadOnlyList<FormatPill> Pills
        {
            get => _pills;
            set
            {
                if (SetProperty(ref _pills, value))
                    OnPropertyChanged(nameof(ShowEmptyClipboardPlaceholder));
            }
        }

        private string _formatsTooltip = string.Empty;
        public string FormatsTooltip
        {
            get => _formatsTooltip;
            set => SetProperty(ref _formatsTooltip, value);
        }

        private long _liveTotalSize;
        /// <summary>
        /// Total byte-size of the current live clipboard (sum of positive
        /// <see cref="ClipboardFormatItem.ContentSizeValue"/>).  Meaningful only on the
        /// sentinel row; updated on every clipboard change via <see cref="UpdateLiveRow"/>.
        /// </summary>
        public long LiveTotalSize
        {
            get => _liveTotalSize;
            set
            {
                if (SetProperty(ref _liveTotalSize, value))
                    OnPropertyChanged(nameof(SizeText));
            }
        }

        public string DateText => IsLive ? "Live" : Timestamp.ToString("yyyy-MM-dd HH:mm:ss");

        /// <summary>
        /// Sentinel: blank when the clipboard is empty / all sizes unknown; otherwise the
        /// formatted live total.  Real entries: the persisted <see cref="TotalSize"/>.
        /// </summary>
        public string SizeText => IsLive
            ? (LiveTotalSize > 0 ? DisplayHelper.FormatFileSize(LiveTotalSize) : string.Empty)
            : DisplayHelper.FormatFileSize(TotalSize);

        /// <summary>
        /// True only for the sentinel row when the current live clipboard has no formats —
        /// drives the "Empty clipboard" placeholder shown in the Formats column.
        /// </summary>
        public bool ShowEmptyClipboardPlaceholder => IsLive && Pills.Count == 0;
    }

    private record SessionPayload(
        IReadOnlyList<(uint FormatId, string FormatName)> Formats,
        IReadOnlyList<FormatPill>                         Pills,
        string                                            PillsText,
        string                                            Tooltip,
        IReadOnlyList<string?>                            TextContents);

    // ── Services ─────────────────────────────────────────────────────────────

    private readonly IHistoryRepository   _history;
    private readonly IHistoryMaintenance  _historyMaintenance;
    private readonly IFormatClassifier    _formatClassifier;
    private readonly ITextDecodingService _textDecoding;
    private readonly IDialogService       _dialogService;
    private readonly Dispatcher           _dispatcher;

    // ── UI callbacks (for operations that can't be data-bound yet) ───────────
    //
    // These are populated by MainWindow and allow the ViewModel to request
    // purely view-level operations without taking a direct dependency on WPF types.

    /// <summary>true = show the history panel; false = hide it.</summary>
    private readonly Action<bool>          _setHistoryPanelVisible;
    /// <summary>true = show the "Loading…" overlay; false = hide it.</summary>
    private readonly Action<bool>          _setHistoryLoadingOverlayVisible;
    /// <summary>
    /// Clears <c>_historySnapshots</c>, calls <c>InitializePreviewState()</c>,
    /// and empties <c>_formats</c> — used when no history item is selected.
    /// </summary>
    private readonly Action                _clearFormatPanelToEmpty;
    /// <summary>
    /// If history snapshots are currently displayed, clears them and calls
    /// <c>RefreshFormats()</c> to restore the live-clipboard view.
    /// </summary>
    private readonly Action                _clearSnapshotsAndRefreshToLive;
    /// <summary>Sets <c>HistoryListView.SelectedItem</c> and scrolls it into view.</summary>
    private readonly Action<HistoryItem?>  _setSelectedHistoryItem;
    /// <summary>Sets <c>MainWindow.IsEnabled</c> (used while recovery is running).</summary>
    private readonly Action<bool>          _setMainWindowEnabled;
    /// <summary>Shows a modal warning MessageBox.</summary>
    private readonly Action<string, string> _showWarning;
    /// <summary>Persists the current application preferences to disk.</summary>
    private readonly Action                _savePreferences;

    // ── Observable properties ─────────────────────────────────────────────────

    private bool _isMonitoring;
    /// <summary>Whether the application is currently monitoring clipboard changes.</summary>
    public bool IsMonitoring
    {
        get => _isMonitoring;
        set { if (SetProperty(ref _isMonitoring, value)) UpdateStatusBarText(); }
    }

    private bool _isTrackingHistory;
    /// <summary>Whether clipboard history is being recorded to the database.</summary>
    public bool IsTrackingHistory
    {
        get => _isTrackingHistory;
        set { if (SetProperty(ref _isTrackingHistory, value)) UpdateStatusBarText(); }
    }

    private bool _isPaused;
    /// <summary>Whether monitoring is temporarily paused.</summary>
    public bool IsPaused
    {
        get => _isPaused;
        private set => SetProperty(ref _isPaused, value);
    }

    private string _statusBarText = string.Empty;
    /// <summary>Text shown in the main status bar item (left side).</summary>
    public string StatusBarText
    {
        get => _statusBarText;
        private set => SetProperty(ref _statusBarText, value);
    }

    private string _historyCountText = string.Empty;
    /// <summary>Entry-count label shown in the history panel header.</summary>
    public string HistoryCountText
    {
        get => _historyCountText;
        private set => SetProperty(ref _historyCountText, value);
    }

    private HistoryItem? _selectedSession;
    /// <summary>
    /// The history entry that is currently selected in the list.
    /// Set by <c>HistoryListView_SelectionChanged</c> when the user clicks,
    /// or by <see cref="LoadHistoryFromDatabase"/> when it programmatically
    /// selects an entry via the <see cref="_setSelectedHistoryItem"/> callback.
    /// </summary>
    internal HistoryItem? SelectedSession
    {
        get => _selectedSession;
        set => SetProperty(ref _selectedSession, value);
    }

    private string _searchText = string.Empty;
    /// <summary>Current text in the history filter box.</summary>
    internal string SearchText
    {
        get => _searchText;
        set => SetProperty(ref _searchText, value);
    }

    // ── Collections ──────────────────────────────────────────────────────────

    /// <summary>The list of clipboard history sessions bound to <c>HistoryListView</c>.</summary>
    public ObservableCollection<HistoryItem> Sessions { get; } = [];

    // ── [Live Clipboard] sentinel row ────────────────────────────────────────

    private HistoryItem? _liveItem;
    /// <summary>
    /// The synthetic [Live Clipboard] row pinned at the top of <see cref="Sessions"/>
    /// whenever history tracking is enabled.  Lazily created; never persisted.
    /// Its <see cref="HistoryItem.Pills"/> and <see cref="HistoryItem.FormatsTooltip"/>
    /// are refreshed on every WM_CLIPBOARDUPDATE via <see cref="UpdateLivePills"/>.
    /// </summary>
    internal HistoryItem LiveItem => _liveItem ??= new HistoryItem
    {
        SessionId   = HistoryItem.LiveSessionId,
        Timestamp   = DateTime.MinValue,
        FormatsText = string.Empty,
        TotalSize   = 0,
        Formats     = Array.Empty<(uint FormatId, string FormatName)>(),
        IsLive      = true,
    };

    /// <summary>
    /// Recomputes the [Live Clipboard] sentinel's pills, tooltip, and total byte-size
    /// from the current live clipboard formats.  Safe to call on every WM_CLIPBOARDUPDATE —
    /// the row updates regardless of which entry is selected.  Must be called on the UI thread.
    /// </summary>
    internal void UpdateLiveRow(IReadOnlyList<ClipboardFormatItem> formats)
    {
        var tuples = formats.Select(f => (f.FormatId, f.Name)).ToList();
        LiveItem.Pills          = _formatClassifier.ComputePills(tuples);
        LiveItem.FormatsTooltip = _formatClassifier.ComputeTooltip(tuples);
        // Sum only known sizes; -1 (n/a) entries are skipped, matching how each format's
        // OriginalSize is computed during CaptureAllFormats.
        long total = 0;
        foreach (var f in formats)
            if (f.ContentSizeValue > 0)
                total += f.ContentSizeValue;
        LiveItem.LiveTotalSize = total;
    }

    // ── Mutable state ─────────────────────────────────────────────────────────

    /// <summary>Maximum number of history sessions to retain.</summary>
    internal int MaxEntries { get; set; } = 100;
    /// <summary>Maximum total blob storage (in megabytes).</summary>
    internal int MaxSizeMb  { get; set; } = 100;
    /// <summary>True once integrity check and schema migration have completed.</summary>
    internal bool DatabaseReady { get; private set; }

    private DateTime         _pauseUntil;
    private DispatcherTimer? _pauseTimer;

    // ── History channel ───────────────────────────────────────────────────────

    /// <summary>
    /// Serialises all database operations through one background consumer.
    /// Each item is an <see cref="Action"/> so both session writes and
    /// limit-enforcement can share the queue.
    /// </summary>
    internal readonly Channel<Action> HistoryChannel = Channel.CreateUnbounded<Action>();

    // ── Constructor ───────────────────────────────────────────────────────────

    internal MainWindowViewModel(
        IHistoryRepository    history,
        IHistoryMaintenance   historyMaintenance,
        IFormatClassifier     formatClassifier,
        ITextDecodingService  textDecoding,
        IDialogService        dialogService,
        Dispatcher            dispatcher,
        Action<bool>          setHistoryPanelVisible,
        Action<bool>          setHistoryLoadingOverlayVisible,
        Action                clearFormatPanelToEmpty,
        Action                clearSnapshotsAndRefreshToLive,
        Action<HistoryItem?>  setSelectedHistoryItem,
        Action<bool>          setMainWindowEnabled,
        Action<string, string> showWarning,
        Action                savePreferences)
    {
        _history                         = history;
        _historyMaintenance              = historyMaintenance;
        _formatClassifier                = formatClassifier;
        _textDecoding                    = textDecoding;
        _dialogService                   = dialogService;
        _dispatcher                      = dispatcher;
        _setHistoryPanelVisible          = setHistoryPanelVisible;
        _setHistoryLoadingOverlayVisible = setHistoryLoadingOverlayVisible;
        _clearFormatPanelToEmpty         = clearFormatPanelToEmpty;
        _clearSnapshotsAndRefreshToLive  = clearSnapshotsAndRefreshToLive;
        _setSelectedHistoryItem          = setSelectedHistoryItem;
        _setMainWindowEnabled            = setMainWindowEnabled;
        _showWarning                     = showWarning;
        _savePreferences                 = savePreferences;
    }

    // ── Status bar text ───────────────────────────────────────────────────────

    /// <summary>
    /// Recomputes <see cref="StatusBarText"/> from the current monitoring/pause/
    /// tracking state and notifies bindings.  Safe to call from any thread
    /// provided the database file-size call is fast (it is).
    /// </summary>
    internal void UpdateStatusBarText()
    {
        if (_isMonitoring && _isPaused)
        {
            if (_isTrackingHistory)
            {
                var size = _historyMaintenance.GetDatabaseFileSize();
                StatusBarText = $"Paused until {PauseUntilLabel}... · Tracking history ({DisplayHelper.FormatFileSize(size)} storage size)...";
            }
            else
            {
                StatusBarText = $"Paused until {PauseUntilLabel}...";
            }
        }
        else if (_isMonitoring && _isTrackingHistory)
        {
            var size = _historyMaintenance.GetDatabaseFileSize();
            StatusBarText = $"Monitoring... · Tracking history ({DisplayHelper.FormatFileSize(size)} storage size)...";
        }
        else
        {
            StatusBarText = _isMonitoring ? "Monitoring..." : "Press F5 to refresh";
        }
    }

    // ── Pause management ─────────────────────────────────────────────────────

    /// <summary>
    /// Formatted "HH:mm" or "d/M/yy HH:mm" string for the tray "Paused until …" label.
    /// Empty when not paused.
    /// </summary>
    internal string PauseUntilLabel =>
        _isPaused
            ? (_pauseUntil.Date == DateTime.Today
                ? _pauseUntil.ToString("t")
                : _pauseUntil.ToString("g"))
            : string.Empty;

    /// <summary>
    /// Activates (or extends) a monitoring pause for the given number of minutes.
    /// Must be called on the UI thread (creates a <see cref="DispatcherTimer"/>).
    /// </summary>
    internal void ActivatePause(int minutes)
    {
        if (_isPaused)
            _pauseUntil = _pauseUntil.AddMinutes(minutes);
        else
        {
            IsPaused    = true;
            _pauseUntil = DateTime.Now.AddMinutes(minutes);
        }

        _pauseTimer?.Stop();
        var remaining = _pauseUntil - DateTime.Now;
        if (remaining <= TimeSpan.Zero)
        {
            CancelPause();
            return;
        }
        _pauseTimer       = new DispatcherTimer { Interval = remaining };
        _pauseTimer.Tick += (_, _) => CancelPause();
        _pauseTimer.Start();
        UpdateStatusBarText();
    }

    /// <summary>Cancels an active pause and resumes monitoring. Must be called on the UI thread.</summary>
    internal void CancelPause()
    {
        IsPaused = false;
        _pauseTimer?.Stop();
        _pauseTimer = null;
        UpdateStatusBarText();
    }

    // ── History channel ───────────────────────────────────────────────────────

    /// <summary>
    /// Long-running background consumer: drains <see cref="HistoryChannel"/> one
    /// action at a time, guaranteeing all DB operations are serialised in arrival
    /// order with no contention.  Start with <c>Task.Run</c>.
    /// </summary>
    internal async Task ProcessHistoryChannelAsync()
    {
        await foreach (var action in HistoryChannel.Reader.ReadAllAsync())
            action();
    }

    // ── Database initialisation orchestration ─────────────────────────────────

    /// <summary>
    /// Enqueues a database integrity check + schema migration on the history channel.
    /// <paramref name="onReady"/> is invoked (still on the background thread) once
    /// the database is ready; callers typically <c>Dispatcher.Invoke</c> inside it.
    /// </summary>
    internal void EnqueueDatabaseInitialization(bool isTrackingHistory, Action? onReady)
        => HistoryChannel.Writer.TryWrite(() => InitializeDatabaseOnChannel(isTrackingHistory, onReady));

    /// <summary>Called as a channel action (background thread).</summary>
    private void InitializeDatabaseOnChannel(bool isTrackingHistory, Action? onReady)
    {
        // Fast path: integrity check and schema migration already completed this session.
        if (DatabaseReady)
        {
            onReady?.Invoke();
            return;
        }

        var status = _historyMaintenance.CheckIntegrity();

        if (status == DatabaseIntegrityStatus.Corrupted)
            ErrorLogger.Log(new InvalidOperationException(
                "history.db failed PRAGMA integrity_check — " +
                (isTrackingHistory ? "prompting user for recovery action."
                                   : "history tracking is off, skipping recovery.")));

        if (status != DatabaseIntegrityStatus.Corrupted)
        {
            // Absent or Healthy: migrate schema and signal readiness.
            _historyMaintenance.MigrateSchema();
            DatabaseReady = true;
            onReady?.Invoke();
            return;
        }

        if (!isTrackingHistory)
            return; // Silently skip; will re-check when history tracking is enabled.

        // Corrupted + tracking is on: prompt the user.
        HandleCorruptDatabase(onReady);
    }

    /// <summary>Called on the history background thread.</summary>
    private void HandleCorruptDatabase(Action? onReady)
    {
        var choice = _dispatcher.Invoke(
            () => _dialogService.ShowDatabaseCorruption(showRecoverOption: true));
        ErrorLogger.LogInfo($"Corruption dialog: user chose {choice}.");

        switch (choice)
        {
            case CorruptionDialogResult.Recover:
                AttemptRecovery(onReady);
                break;
            case CorruptionDialogResult.DeleteAndReset:
                DeleteAndReset(onReady);
                break;
            default: // Disable
                _dispatcher.Invoke(DisableHistoryTracking);
                DrainHistoryChannel();
                break;
        }
    }

    /// <summary>Called on the history background thread.</summary>
    private void AttemptRecovery(Action? onReady)
    {
        Action? closeWindow = null;
        _dispatcher.Invoke(() =>
        {
            closeWindow = _dialogService.ShowDatabaseRecovering();
            _setMainWindowEnabled(false);
        });

        var result = _historyMaintenance.TryRecover();

        _dispatcher.Invoke(() =>
        {
            closeWindow?.Invoke();
            _setMainWindowEnabled(true);
        });

        if (result.Success)
        {
            ErrorLogger.LogInfo(
                $"Recovery succeeded using {result.Strategy}. " +
                $"Sessions recovered: {result.SessionsRecovered}, lost: {result.SessionsLost}.");

            _historyMaintenance.MigrateSchema();
            DatabaseReady = true;

            if (result.HadUnreadableRows)
            {
                _dispatcher.Invoke(() => _showWarning(
                    "Recovery completed. Some history entries may be incomplete or missing " +
                    "due to unreadable rows in the corrupted database.",
                    "Recovery Warning"));
            }

            onReady?.Invoke();
        }
        else
        {
            ErrorLogger.LogInfo($"Recovery failed (strategy: {result.Strategy}).");

            var choice = _dispatcher.Invoke(
                () => _dialogService.ShowDatabaseCorruption(showRecoverOption: false));
            ErrorLogger.LogInfo($"Post-recovery failure dialog: user chose {choice}.");

            switch (choice)
            {
                case CorruptionDialogResult.DeleteAndReset:
                    DeleteAndReset(onReady);
                    break;
                default: // Disable
                    _dispatcher.Invoke(DisableHistoryTracking);
                    DrainHistoryChannel();
                    break;
            }
        }
    }

    private void DeleteAndReset(Action? onReady)
    {
        _historyMaintenance.DeleteDatabase();
        _historyMaintenance.InitializeFreshDatabase();
        DatabaseReady = true;
        onReady?.Invoke();
    }

    /// <summary>
    /// Drains all pending actions from the history channel.
    /// Called from within a channel action (the consumer is paused), so TryRead is safe.
    /// </summary>
    private void DrainHistoryChannel()
    {
        while (HistoryChannel.Reader.TryRead(out _)) { }
    }

    /// <summary>
    /// Turns history tracking off, hides the history panel and loading overlay,
    /// clears sessions, and restores the live-clipboard view.
    /// Does <em>not</em> save preferences — callers that own the save call should
    /// use this; callers that don't should call <see cref="DisableHistoryTracking"/>.
    /// Must be called on the UI thread.
    /// </summary>
    internal void OnHistoryTrackingDisabled()
    {
        IsTrackingHistory = false;
        _setHistoryLoadingOverlayVisible(false);
        _setHistoryPanelVisible(false);
        Sessions.Clear();
        _clearSnapshotsAndRefreshToLive();
    }

    /// <summary>
    /// Like <see cref="OnHistoryTrackingDisabled"/> but also saves preferences.
    /// Used by the corruption-dialog path where there is no outer save call.
    /// Must be called on the UI thread.
    /// </summary>
    internal void DisableHistoryTracking()
    {
        OnHistoryTrackingDisabled();
        _savePreferences();
    }

    /// <summary>
    /// Called when monitoring is turned off. Cancels any active pause and, if history
    /// was being tracked, disables it (without saving — the caller saves preferences).
    /// Must be called on the UI thread.
    /// </summary>
    internal void OnMonitoringDisabled()
    {
        if (IsPaused)
            CancelPause();
        if (IsTrackingHistory)
            OnHistoryTrackingDisabled();
    }

    /// <summary>
    /// Clears the session list and restores the live-clipboard view.
    /// Called when the Settings dialog clears all history.
    /// Must be called on the UI thread.
    /// </summary>
    internal void OnHistoryCleared()
    {
        Sessions.Clear();
        _clearSnapshotsAndRefreshToLive();
    }

    // ── Session payload builder ───────────────────────────────────────────────

    private SessionPayload BuildSessionPayload(IReadOnlyList<FormatSnapshot> snapshots)
    {
        var formats      = snapshots.Select(s => (s.FormatId, s.FormatName)).ToList();
        var pills        = _formatClassifier.ComputePills(formats);
        var pillsText    = string.Join(" ", pills.Select(p => p.Label));
        var tooltip      = _formatClassifier.ComputeTooltip(formats);
        var textContents = snapshots
            .Select(s =>
            {
                if (s.Data == null) return (string?)null;
                var result = _textDecoding.Decode(s.FormatId, s.FormatName, s.Data);
                return result.Success ? result.Text : null;
            })
            .ToList();
        return new SessionPayload(formats, pills, pillsText, tooltip, textContents);
    }

    // ── History session writing ───────────────────────────────────────────────

    /// <summary>Enqueues a new history session write. Called from the UI thread.</summary>
    internal void EnqueueHistorySession(IReadOnlyList<FormatSnapshot> snapshots, DateTime timestamp)
        => HistoryChannel.Writer.TryWrite(() => WriteHistorySession(snapshots, timestamp));

    /// <summary>Enqueues the startup clipboard snapshot write. Called from the UI thread.</summary>
    internal void EnqueueStartupSession(IReadOnlyList<FormatSnapshot> snapshots, DateTime timestamp)
        => HistoryChannel.Writer.TryWrite(() => WriteStartupSession(snapshots, timestamp));

    /// <summary>
    /// Background-thread handler for the startup clipboard snapshot.
    /// Writes a new session only when the clipboard differs from the last history entry;
    /// always finishes by reloading the history list on the UI thread and selecting the
    /// newest entry (whether newly added or the pre-existing last one).
    /// </summary>
    private void WriteStartupSession(IReadOnlyList<FormatSnapshot> snapshots, DateTime timestamp)
    {
        try
        {
            if (!_history.IsDuplicateOfLastSession(snapshots))
            {
                var payload          = BuildSessionPayload(snapshots);
                var maxDatabaseBytes = (long)MaxSizeMb * 1024L * 1024L;
                _history.AddSession(snapshots, payload.TextContents, payload.PillsText,
                    timestamp, MaxEntries, maxDatabaseBytes);
            }
        }
        catch
        {
            // Silently ignore write failures.
        }
        finally
        {
            // Reload the list and select the [Live Clipboard] sentinel in all cases.
            _dispatcher.InvokeAsync(() => LoadHistoryFromDatabase(null, selectLive: true));
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
            var maxDatabaseBytes = (long)MaxSizeMb * 1024L * 1024L;
            var payload          = BuildSessionPayload(snapshots);

            var (sessionId, trimmed) = _history.AddSession(
                snapshots, payload.TextContents, payload.PillsText,
                timestamp, MaxEntries, maxDatabaseBytes);

            var formatsText = _history.BuildFormatsText(snapshots);
            var totalSize   = snapshots.Sum(s => s.OriginalSize);

            _dispatcher.InvokeAsync(() =>
            {
                var activeFilter = SearchText;
                if (trimmed || !string.IsNullOrEmpty(activeFilter))
                {
                    if (string.IsNullOrEmpty(activeFilter))
                    {
                        // Trimmed with no active filter: reload and select the sentinel.
                        LoadHistoryFromDatabase(null, selectLive: true);
                    }
                    else
                    {
                        // Filter is active: reload filtered list, preserving the current selection
                        // (falls back to the sentinel when the previously selected real entry is
                        // no longer in the filtered set).
                        LoadHistoryFromDatabase(activeFilter, SelectedSession?.SessionId);
                    }
                }
                else
                {
                    var entry = new HistoryItem
                    {
                        SessionId      = sessionId,
                        Timestamp      = timestamp,
                        FormatsText    = formatsText,
                        TotalSize      = totalSize,
                        Formats        = payload.Formats,
                        Pills          = payload.Pills,
                        FormatsTooltip = payload.Tooltip,
                    };
                    // Insert below the [Live Clipboard] sentinel (which sits at index 0).
                    var insertIndex = (Sessions.Count > 0 && Sessions[0].IsLive) ? 1 : 0;
                    Sessions.Insert(insertIndex, entry);
                    // Always snap selection back to the sentinel so the live view is the default home.
                    _setSelectedHistoryItem(LiveItem);
                }
                UpdateStatusBarText();
            });
        }
        catch
        {
            // Background write failure — silently ignored.
        }
    }

    // ── Settings-driven history maintenance ──────────────────────────────────

    /// <summary>
    /// Enqueues a limit-enforcement pass on the history channel (only when the channel
    /// is currently idle).  If rows are removed and history tracking is active, reloads
    /// the list and refreshes the view; always updates the status bar text afterwards.
    /// </summary>
    internal void EnqueueEnforceLimits()
    {
        if (HistoryChannel.Reader.Count != 0) return;

        var maxEntries      = MaxEntries;
        var maxBytes        = (long)MaxSizeMb * 1024L * 1024L;
        var trackingHistory = IsTrackingHistory;
        HistoryChannel.Writer.TryWrite(() =>
        {
            var removed = _historyMaintenance.EnforceLimits(maxEntries, maxBytes);
            if (!removed && !trackingHistory) return;

            _dispatcher.InvokeAsync(() =>
            {
                if (removed && trackingHistory)
                {
                    _clearSnapshotsAndRefreshToLive();
                    LoadHistoryFromDatabase();
                }
                UpdateStatusBarText();
            });
        });
    }

    // ── History mutation ─────────────────────────────────────────────────────

    /// <summary>
    /// Enqueues a single-session delete on the history channel, then reloads the list.
    /// <paramref name="filter"/> is the search text active at the time the user confirmed.
    /// </summary>
    internal void EnqueueDeleteSession(long sessionId, string? filter)
        => HistoryChannel.Writer.TryWrite(() =>
        {
            try { _history.DeleteSession(sessionId); }
            catch { /* ignore write failure */ }

            _dispatcher.InvokeAsync(() =>
            {
                _clearSnapshotsAndRefreshToLive();
                LoadHistoryFromDatabase(filter);
            });
        });

    /// <summary>
    /// Enqueues a clear-all history operation on the history channel, then reloads
    /// the (now-empty) list.
    /// </summary>
    internal void EnqueueClearAll()
        => HistoryChannel.Writer.TryWrite(() =>
        {
            try { _historyMaintenance.ClearHistory(); }
            catch { /* ignore write failure */ }

            _dispatcher.InvokeAsync(() =>
            {
                _clearSnapshotsAndRefreshToLive();
                LoadHistoryFromDatabase();
            });
        });

    // ── History loading ──────────────────────────────────────────────────────

    /// <summary>
    /// Reloads the history list from the database (must be called on the UI thread).
    /// Always pins the <see cref="LiveItem"/> sentinel at index 0.  Selection rules:
    /// <list type="bullet">
    ///   <item><paramref name="preserveSessionId"/> wins when supplied; resolves to
    ///         <see cref="LiveItem"/> for <see cref="HistoryItem.LiveSessionId"/>.</item>
    ///   <item><paramref name="selectLive"/> selects the sentinel.</item>
    ///   <item>Otherwise (or when the requested item no longer exists), falls back to
    ///         the sentinel so the format panel always has a live target.</item>
    /// </list>
    /// </summary>
    internal void LoadHistoryFromDatabase(
        string? filter            = null,
        long?   preserveSessionId = null,
        bool    selectLive        = false)
    {
        try
        {
            var sessions = _history.LoadSessions(filter);
            Sessions.Clear();
            foreach (var session in sessions)
            {
                var pills   = _formatClassifier.ComputePills(session.Formats);
                var tooltip = _formatClassifier.ComputeTooltip(session.Formats);
                Sessions.Add(new HistoryItem
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

            // Pin the [Live Clipboard] sentinel at the top of the list.  Survives filter,
            // clear-all, trim, delete — it's re-inserted on every reload.
            Sessions.Insert(0, LiveItem);

            // Determine which item to (re)select after the reload.
            HistoryItem? toSelect = null;
            if (preserveSessionId.HasValue)
            {
                toSelect = preserveSessionId.Value == HistoryItem.LiveSessionId
                    ? LiveItem
                    : Sessions.FirstOrDefault(h => h.SessionId == preserveSessionId.Value);
            }
            else if (selectLive)
            {
                toSelect = LiveItem;
            }

            // Fall back to the sentinel so the format panel never goes blank.
            toSelect ??= LiveItem;

            _setSelectedHistoryItem(toSelect);  // also scrolls into view in MainWindow

            UpdateHistoryCountText(Sessions.Count - 1, filter);  // exclude sentinel
        }
        catch
        {
            // Silently ignore load failures.
        }
    }

    private void UpdateHistoryCountText(int filteredCount, string? filter)
    {
        int    total = string.IsNullOrEmpty(filter) ? filteredCount : _history.GetSessionCount();
        string n     = filteredCount == 1 ? "1 item" : $"{filteredCount} items";
        HistoryCountText = string.IsNullOrEmpty(filter) || filteredCount == total
            ? n
            : $"{n} ({total} total)";
    }
}
