using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Classifies clipboard formats into display categories for the history list.</summary>
public interface IFormatClassifier
{
    /// <summary>
    /// Builds the ordered list of pills that summarise the format categories present
    /// in <paramref name="formats"/>. "OTHER" is included only when no other category applies.
    /// </summary>
    IReadOnlyList<FormatPill> ComputePills(
        IReadOnlyList<(uint FormatId, string FormatName)> formats);

    /// <summary>Builds the multi-line tooltip text for a history row.</summary>
    string ComputeTooltip(
        IReadOnlyList<(uint FormatId, string FormatName)> formats);

    /// <summary>
    /// Returns the pill label (<c>"IMG"</c>, <c>"TXT"</c>, <c>"HTML"</c>,
    /// <c>"RTF"</c>, or <c>"FILE"</c>) for a single format, or
    /// <see langword="null"/> when the format falls into the uncategorised bucket.
    /// </summary>
    string? GetFormatPillLabel(uint formatId, string formatName);
}
