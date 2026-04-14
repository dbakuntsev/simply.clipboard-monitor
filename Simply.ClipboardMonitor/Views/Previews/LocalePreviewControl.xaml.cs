using Simply.ClipboardMonitor.Services;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class LocalePreviewControl : UserControl, IPreviewTab
{
    public TabItem TabItem { get; }
    public int Priority => 1;

    public LocalePreviewControl()
    {
        InitializeComponent();
        TabItem = new TabItem { Header = "Locale", Content = this };
    }

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = formatId == CF_LOCALE;

        if (!TabItem.IsEnabled)
        {
            SetUnavailable("Locale preview unavailable for this format.");
            return;
        }

        if (bytes == null)
        {
            SetUnavailable("Locale preview requires byte-addressable clipboard data.");
            return;
        }

        if (bytes.Length < 4)
        {
            SetUnavailable("CF_LOCALE data is too short to decode.");
            return;
        }

        var lcid    = BitConverter.ToUInt32(bytes, 0);
        var lcidHex = $"0x{lcid:X4}";

        string? tag         = null;
        string? displayName = null;
        try
        {
            var culture = CultureInfo.GetCultureInfo((int)lcid);
            tag         = culture.Name;
            displayName = culture.DisplayName;
        }
        catch { }

        LocaleStatusTextBlock.Text    = string.Empty;
        LocaleLcidTextBlock.Text      = lcidHex;
        LocaleTagTextBlock.Text       = tag ?? string.Empty;
        LocaleNameTextBlock.Text      = displayName ?? string.Empty;
        LocaleTagRow.Visibility       = tag != null ? Visibility.Visible : Visibility.Collapsed;
        LocaleNameRow.Visibility      = displayName != null ? Visibility.Visible : Visibility.Collapsed;
        LocaleContentPanel.Visibility = Visibility.Visible;
    }

    public void Reset()
    {
        TabItem.IsEnabled = true;
        SetUnavailable("Select a clipboard format to preview.");
    }

    private void SetUnavailable(string message)
    {
        LocaleStatusTextBlock.Text    = message;
        LocaleContentPanel.Visibility = Visibility.Collapsed;
    }
}
