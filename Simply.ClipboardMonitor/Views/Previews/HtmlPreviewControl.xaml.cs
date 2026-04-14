using Microsoft.Web.WebView2.Core;
using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Services;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class HtmlPreviewControl : UserControl, IPreviewTab
{
    private readonly ITextDecodingService _textDecoding;

    private bool    _webViewReady;
    private bool    _initAttempted;
    private string? _pendingHtml;

    public TabItem TabItem { get; }
    public int Priority => 2;

    public HtmlPreviewControl(ITextDecodingService textDecoding)
    {
        _textDecoding = textDecoding;
        InitializeComponent();
        TabItem = new TabItem { Header = "HTML", Content = this };
    }

    // ── IPreviewTab implementation ───────────────────────────────────────────

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = IsHtmlFormat(name);

        if (!TabItem.IsEnabled)
        {
            SetUnavailable("HTML preview unavailable for this format.");
            return;
        }

        if (bytes == null)
        {
            SetUnavailable("HTML preview requires byte-addressable clipboard data.");
            return;
        }

        try
        {
            // "HTML Format" (CF_HTML) has an ASCII header with byte offsets; strip it first.
            // For all other HTML formats the payload is a bare HTML string whose encoding is
            // detected by the standard text pipeline (UTF-16 LE, UTF-8 BOM, ANSI).
            var html = TryExtractCfHtml(bytes)
                    ?? _textDecoding.Decode(0, name, bytes).Text
                    ?? string.Empty;
            SetHtml(html);
        }
        catch
        {
            SetUnavailable("Could not decode HTML content.");
        }
    }

    public void Reset()
    {
        TabItem.IsEnabled = true;
        SetUnavailable("Select a clipboard format to preview.");
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private void SetHtml(string html)
    {
        _pendingHtml             = null;
        HtmlStatusTextBlock.Text = string.Empty;

        if (_webViewReady)
        {
            HtmlWebView.Visibility = Visibility.Visible;
            HtmlWebView.NavigateToString(html);
        }
        else
        {
            // WebView2 not yet initialised — queue content for when init completes.
            _pendingHtml = html;
        }
    }

    private void SetUnavailable(string message)
    {
        _pendingHtml             = null;
        HtmlStatusTextBlock.Text = message;
        HtmlWebView.Visibility   = Visibility.Collapsed;
        if (_webViewReady)
            HtmlWebView.CoreWebView2.Navigate("about:blank");
    }

    private async void HtmlWebView_Loaded(object sender, RoutedEventArgs e)
    {
        if (_initAttempted) return;
        _initAttempted = true;
        await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Simply.ClipboardMonitor", "WebView2");

            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
            await HtmlWebView.EnsureCoreWebView2Async(env);

            var settings = HtmlWebView.CoreWebView2.Settings;
            settings.IsScriptEnabled               = false;
            settings.AreDefaultContextMenusEnabled = false;
            settings.AreDevToolsEnabled            = false;
            settings.IsStatusBarEnabled            = false;

            // The preview is read-only.  Cancel all user-initiated navigation (link clicks,
            // anchor jumps, form submits).  NavigationStarting fires for programmatic calls too
            // (NavigateToString, Navigate("about:blank")), so we must not cancel those —
            // IsUserInitiated is false for code-driven navigations and true for user gestures.
            HtmlWebView.CoreWebView2.NewWindowRequested += (_, e) => e.Handled = true;
            HtmlWebView.CoreWebView2.NavigationStarting += (_, e) =>
            {
                if (e.IsUserInitiated) e.Cancel = true;
            };

            _webViewReady = true;

            // Apply any HTML that was set while initialisation was in progress.
            if (_pendingHtml is { } html)
            {
                _pendingHtml             = null;
                HtmlWebView.Visibility   = Visibility.Visible;
                HtmlStatusTextBlock.Text = string.Empty;
                HtmlWebView.NavigateToString(html);
            }
        }
        catch (Exception ex)
        {
            ErrorLogger.Log(ex);
            HtmlStatusTextBlock.Text =
                "HTML preview is unavailable: WebView2 Runtime not found. " +
                "Install the Microsoft Edge WebView2 Evergreen Runtime to enable this feature.";
        }
    }

    /// <summary>
    /// Attempts to extract the HTML document portion from a CF_HTML ("HTML Format") payload
    /// by reading the <c>StartHTML</c> and <c>EndHTML</c> byte offsets in its ASCII header.
    /// Returns <see langword="null"/> if the format is not CF_HTML or the header cannot be
    /// parsed, so the caller can fall back to standard encoding-aware text decoding.
    /// </summary>
    private static string? TryExtractCfHtml(byte[] bytes)
    {
        try
        {
            var headerRegion = Encoding.ASCII.GetString(bytes, 0, Math.Min(512, bytes.Length));

            if (!headerRegion.StartsWith("Version:", StringComparison.OrdinalIgnoreCase))
                return null;

            var startMatch = Regex.Match(headerRegion, @"StartHTML:(-?\d+)", RegexOptions.IgnoreCase);
            var endMatch   = Regex.Match(headerRegion, @"EndHTML:(-?\d+)",   RegexOptions.IgnoreCase);

            if (startMatch.Success && endMatch.Success
                && int.TryParse(startMatch.Groups[1].Value, out var startHtml)
                && int.TryParse(endMatch.Groups[1].Value,   out var endHtml)
                && startHtml >= 0 && endHtml > startHtml && endHtml <= bytes.Length)
            {
                return Encoding.UTF8.GetString(bytes, startHtml, endHtml - startHtml);
            }
        }
        catch { }
        return null;
    }
}
