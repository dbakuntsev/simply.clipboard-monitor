using System.Windows.Media;

namespace Simply.ClipboardMonitor.Models;

/// <summary>A coloured badge representing one format category in the history list.</summary>
public sealed record FormatPill(string Label, SolidColorBrush Background);
