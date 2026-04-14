using Simply.ClipboardMonitor.Services;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using static Simply.ClipboardMonitor.Common.ClipboardFormatConstants;

namespace Simply.ClipboardMonitor.Views.Previews;

public sealed partial class RtfPreviewControl : UserControl, IPreviewTab
{
    public TabItem TabItem { get; }
    public int Priority => 2;

    public RtfPreviewControl()
    {
        InitializeComponent();
        TabItem = new TabItem { Header = "RTF", Content = this };
    }

    // ── IPreviewTab implementation ───────────────────────────────────────────

    public void Update(uint formatId, string name, byte[]? bytes)
    {
        TabItem.IsEnabled = IsRtfFormat(name);

        if (!TabItem.IsEnabled)
        {
            SetUnavailable("RTF preview unavailable for this format.");
            return;
        }

        if (bytes == null)
        {
            SetUnavailable("RTF preview requires byte-addressable clipboard data.");
            return;
        }

        try
        {
            var doc = new FlowDocument();
            using var stream = new MemoryStream(bytes);
            new TextRange(doc.ContentStart, doc.ContentEnd).Load(stream, DataFormats.Rtf);

            // WPF's RTF importer stores compound GDI face names (e.g. "Aptos SemiBold")
            // verbatim as FontFamily.Source.  WPF's resolver treats them as family names,
            // fails to find them, and falls back to normal weight.  Walk the document and
            // split any compound names into base family + explicit weight / style.
            NormalizeDocumentFonts(doc);

            RtfContentBox.Document   = doc;
            RtfContentBox.Visibility = Visibility.Visible;

            var text      = new TextRange(doc.ContentStart, doc.ContentEnd).Text;
            var charCount = text.Count(c => c != '\r' && c != '\n' && c != '\0');
            RtfStatusTextBlock.Text = $"Showing {charCount:N0} characters ({bytes.Length:N0} bytes).";
        }
        catch
        {
            SetUnavailable("Could not parse RTF content.");
        }
    }

    public void Reset()
    {
        TabItem.IsEnabled = true;
        SetUnavailable("Select a clipboard format to preview.");
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private void SetUnavailable(string message)
    {
        RtfStatusTextBlock.Text  = message;
        RtfContentBox.Visibility = Visibility.Collapsed;
        RtfContentBox.Document   = new FlowDocument();
    }

    // ── RTF font normalization ───────────────────────────────────────────────

    private static readonly Dictionary<string, FontWeight> _weightTokens =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Thin"]       = FontWeights.Thin,
            ["ExtraLight"] = FontWeights.ExtraLight,
            ["UltraLight"] = FontWeights.ExtraLight,
            ["Light"]      = FontWeights.Light,
            ["Medium"]     = FontWeights.Medium,
            ["SemiBold"]   = FontWeights.SemiBold,
            ["DemiBold"]   = FontWeights.SemiBold,
            ["Bold"]       = FontWeights.Bold,
            ["ExtraBold"]  = FontWeights.ExtraBold,
            ["UltraBold"]  = FontWeights.ExtraBold,
            ["Black"]      = FontWeights.Black,
            ["Heavy"]      = FontWeights.Black,
            ["ExtraBlack"] = FontWeights.ExtraBlack,
            ["UltraBlack"] = FontWeights.ExtraBlack,
        };

    private static readonly Dictionary<string, FontStyle> _styleTokens =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Italic"]  = FontStyles.Italic,
            ["Oblique"] = FontStyles.Oblique,
        };

    /// <summary>
    /// Walks every <see cref="TextElement"/> in <paramref name="doc"/> that has a locally-set
    /// <see cref="TextElement.FontFamilyProperty"/> and splits compound GDI face names
    /// (e.g. "Aptos SemiBold", "Segoe UI Bold Italic") into a base family name plus explicit
    /// <see cref="TextElement.FontWeight"/> and <see cref="TextElement.FontStyle"/> values.
    /// </summary>
    private static void NormalizeDocumentFonts(FlowDocument doc)
    {
        var queue = new Queue<TextElement>();
        foreach (Block b in doc.Blocks) queue.Enqueue(b);

        while (queue.Count > 0)
        {
            var elem = queue.Dequeue();

            if (elem.ReadLocalValue(TextElement.FontFamilyProperty) is FontFamily family)
                TryFixCompoundFontName(elem, family.Source);

            switch (elem)
            {
                case Paragraph p:
                    foreach (Inline i in p.Inlines) queue.Enqueue(i);
                    break;
                case Section s:
                    foreach (Block b in s.Blocks) queue.Enqueue(b);
                    break;
                case Span sp:
                    foreach (Inline i in sp.Inlines) queue.Enqueue(i);
                    break;
                case List lst:
                    foreach (ListItem li in lst.ListItems) queue.Enqueue(li);
                    break;
                case ListItem li:
                    foreach (Block b in li.Blocks) queue.Enqueue(b);
                    break;
                case Table tbl:
                    foreach (TableRowGroup rg in tbl.RowGroups)
                        foreach (TableRow row in rg.Rows)
                            foreach (TableCell cell in row.Cells)
                                foreach (Block b in cell.Blocks) queue.Enqueue(b);
                    break;
            }
        }
    }

    /// <summary>
    /// Peels recognized weight and style tokens off the end of <paramref name="familySource"/>
    /// and applies them to <paramref name="element"/> as explicit font properties, then sets the
    /// family to the remaining base name.  Does nothing if no tokens are recognized.
    /// </summary>
    private static void TryFixCompoundFontName(TextElement element, string familySource)
    {
        var parts = familySource.Split(' ');
        if (parts.Length < 2) return;

        FontWeight? weight = null;
        FontStyle?  style  = null;
        int end = parts.Length;

        while (end > 1)
        {
            var token = parts[end - 1];
            if (style == null && _styleTokens.TryGetValue(token, out var s))
            { style = s; end--; }
            else if (weight == null && _weightTokens.TryGetValue(token, out var w))
            { weight = w; end--; }
            else break;
        }

        if (end == parts.Length) return;

        element.FontFamily = new FontFamily(string.Join(" ", parts, 0, end));
        if (weight.HasValue) element.FontWeight = weight.Value;
        if (style.HasValue)  element.FontStyle  = style.Value;
    }
}
