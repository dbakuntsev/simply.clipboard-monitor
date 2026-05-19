using System.Text.RegularExpressions;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Default <see cref="IFormatNotificationMatcher"/>.  Compiles each user-supplied
/// pattern to a case-insensitive whole-string regex with <c>*</c> as wildcard.
/// </summary>
internal sealed class FormatNotificationMatcher : IFormatNotificationMatcher
{
    private static readonly char[] LineSeparators = ['\r', '\n'];

    private bool _enabled;
    private IReadOnlyList<Regex> _patterns = [];

    public bool Enabled => _enabled;

    public bool HasActivePatterns => _enabled && _patterns.Count > 0;

    public void Configure(bool enabled, string? patternsText)
    {
        _enabled  = enabled;
        _patterns = ParsePatterns(patternsText);
    }

    public IReadOnlyList<string> Match(IEnumerable<string> formatNames)
    {
        if (!_enabled || _patterns.Count == 0)
            return [];

        var matched = new List<string>();
        var seen    = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var name in formatNames)
        {
            if (string.IsNullOrEmpty(name) || !seen.Add(name))
                continue;

            foreach (var pattern in _patterns)
            {
                if (pattern.IsMatch(name))
                {
                    matched.Add(name);
                    break;
                }
            }
        }

        return matched;
    }

    private static IReadOnlyList<Regex> ParsePatterns(string? patternsText)
    {
        if (string.IsNullOrWhiteSpace(patternsText))
            return [];

        var regexes = new List<Regex>();
        foreach (var raw in patternsText.Split(LineSeparators, StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = raw.Trim();
            if (trimmed.Length == 0) continue;

            // Escape regex meta-characters, then unescape \* back to .* so * acts as a wildcard.
            var pattern = "^" + Regex.Escape(trimmed).Replace("\\*", ".*") + "$";
            try
            {
                regexes.Add(new Regex(pattern,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
            }
            catch (ArgumentException)
            {
                // Should not happen given the construction above, but guard against future edits.
            }
        }
        return regexes;
    }
}
