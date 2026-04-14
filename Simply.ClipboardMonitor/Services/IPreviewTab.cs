using System.Windows.Controls;

namespace Simply.ClipboardMonitor.Services;

/// <summary>
/// Represents a preview tab in the main window's content <see cref="TabControl"/>.
/// Implementations are registered in the DI container and injected into the main window,
/// which builds the <see cref="TabControl"/> dynamically at startup.
/// Adding a new preview tab requires no changes to the main window.
/// </summary>
public interface IPreviewTab
{
    /// <summary>
    /// The <see cref="TabItem"/> that wraps this control.
    /// Each implementation creates and owns its <see cref="TabItem"/> in its constructor.
    /// </summary>
    TabItem TabItem { get; }

    /// <summary>
    /// Controls which tab is selected when the currently active tab becomes disabled.
    /// The enabled tab with the lowest <see cref="Priority"/> value is activated.
    /// Suggested values: Hex = 0, Text = 1, Image = 1, Locale = 1, HTML = 2, RTF = 2.
    /// </summary>
    int Priority { get; }

    /// <summary>
    /// Updates the preview for the given clipboard format.
    /// The implementation is responsible for updating
    /// <see cref="TabItem"/>.<see cref="TabItem.IsEnabled"/> to reflect whether it can
    /// handle this format.
    /// When <paramref name="bytes"/> is <see langword="null"/> the format is recognised but
    /// its data is not byte-addressable; the tab should remain enabled (if it handles the
    /// format) and show an appropriate unavailability message.
    /// </summary>
    void Update(uint formatId, string name, byte[]? bytes);

    /// <summary>
    /// Resets the tab to its initial state (no format selected).
    /// Sets <see cref="TabItem"/>.<see cref="TabItem.IsEnabled"/> to
    /// <see langword="true"/> and clears any displayed content.
    /// </summary>
    void Reset();
}
