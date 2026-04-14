using Simply.ClipboardMonitor.Models;
using Simply.ClipboardMonitor.Services;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class TextPreviewControl : UserControl, IPreviewTab
{
    private readonly ITextDecodingService _textDecoding;

    // ── Per-selection state ──────────────────────────────────────────────────
    private byte[]?                      _bytes;
    private bool                         _suppressChange;
    private IReadOnlyList<EncodingItem>? _encodingItems;
    private bool                         _manuallyChanged;

    // ── IPreviewTab ──────────────────────────────────────────────────────────
    public TabItem TabItem { get; }
    public int Priority => 1;

    // ── Export helpers (read by the main window to build FormatExportContext) ─
    public Encoding? AutoDetectedEncoding { get; private set; }

    /// <summary>
    /// The encoding explicitly selected by the user in the combobox, or
    /// <see langword="null"/> if the user has not overridden the auto-detected encoding.
    /// </summary>
    public Encoding? ManuallyChangedEncoding =>
        _manuallyChanged && TextEncodingComboBox.SelectedItem is EncodingItem { Encoding: var enc }
            ? enc : null;

    public TextPreviewControl(ITextDecodingService textDecoding)
    {
        _textDecoding = textDecoding;
        InitializeComponent();
        TabItem = new TabItem { Header = "Text", Content = this };
    }

    // ── IPreviewTab implementation ───────────────────────────────────────────

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = _textDecoding.IsTextCompatible(formatId, name);

        if (!TabItem.IsEnabled)
        {
            SetUnavailable("Text preview unavailable for this format.");
            return;
        }

        if (bytes == null)
        {
            SetUnavailable("Text preview requires byte-addressable clipboard data.");
            return;
        }

        _bytes           = bytes;
        AutoDetectedEncoding = null;
        _manuallyChanged = false;

        var result = _textDecoding.Decode(formatId, name, bytes);

        if (!result.Success)
        {
            SetUnavailable(result.FailureMessage ?? "Text preview unavailable for this format.");
            return;
        }

        AutoDetectedEncoding            = result.DetectedEncoding;
        TextContentTextBox.Foreground   = SystemColors.WindowTextBrush;
        TextContentTextBox.Text         = result.Text ?? string.Empty;
        TextStatusTextBlock.Text        = _textDecoding.GetDecodedTextStats(result.Text ?? string.Empty);
        PopulateEncodingComboBox(result.DetectedEncoding);
    }

    public void Reset()
    {
        TabItem.IsEnabled = true;
        SetUnavailable("Text preview unavailable for this format.");
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private void SetUnavailable(string message)
    {
        _bytes               = null;
        AutoDetectedEncoding = null;
        _manuallyChanged     = false;
        TextStatusTextBlock.Text       = message;
        TextContentTextBox.Foreground  = SystemColors.WindowTextBrush;
        TextContentTextBox.Text        = string.Empty;
        TextEncodingComboBox.IsEnabled = false;
    }

    private void PopulateEncodingComboBox(Encoding? preselect)
    {
        _encodingItems ??= _textDecoding.GetAvailableEncodings();

        if (TextEncodingComboBox.ItemsSource == null)
            TextEncodingComboBox.ItemsSource = _encodingItems;

        TextEncodingComboBox.IsEnabled = true;

        _suppressChange = true;
        TextEncodingComboBox.SelectedItem = preselect != null
            ? _encodingItems.FirstOrDefault(e => e.Encoding.CodePage == preselect.CodePage)
            : null;
        _suppressChange = false;
    }

    private void TextEncodingComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressChange)
            return;
        if (TextEncodingComboBox.SelectedItem is not EncodingItem item || _bytes == null)
            return;

        _manuallyChanged = true;

        var result = _textDecoding.DecodeWith(_bytes, item.Encoding);

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
}
