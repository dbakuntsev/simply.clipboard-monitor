using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Services;
using System.Windows.Controls;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class HexPreviewControl : UserControl, IPreviewTab
{
    public TabItem TabItem { get; }
    public int Priority => 0;

    public HexPreviewControl()
    {
        InitializeComponent();
        TabItem = new TabItem { Header = "Hex", Content = this };
    }

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = true; // Hex renders any format that has bytes

        if (bytes != null)
        {
            var rows = (bytes.Length + HexRowCollection.BytesPerRow - 1) / HexRowCollection.BytesPerRow;
            HexStatusTextBlock.Text = $"Showing {bytes.Length:N0} bytes ({rows:N0} rows).";
            HexListView.ItemsSource = new HexRowCollection(bytes);
        }
        else
        {
            HexStatusTextBlock.Text = "Hex preview requires byte-addressable clipboard data.";
            HexListView.ItemsSource = null;
        }
    }

    public void Reset()
    {
        TabItem.IsEnabled       = true;
        HexStatusTextBlock.Text = "Select a clipboard format to preview.";
        HexListView.ItemsSource = null;
    }
}
