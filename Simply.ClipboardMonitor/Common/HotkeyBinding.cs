using System.Windows.Input;

namespace Simply.ClipboardMonitor.Common;

/// <summary>
/// Represents a global hotkey combination (modifier flags + virtual key code).
/// Immutable value type; use <see cref="TryParse"/> to deserialise from preferences
/// and <see cref="ToString"/> to serialise back.
/// </summary>
public readonly struct HotkeyBinding : IEquatable<HotkeyBinding>
{
    // RegisterHotKey fsModifiers flags
    public const uint MOD_ALT      = 0x0001;
    public const uint MOD_CONTROL  = 0x0002;
    public const uint MOD_SHIFT    = 0x0004;
    public const uint MOD_WIN      = 0x0008;
    public const uint MOD_NOREPEAT = 0x4000;

    public uint Modifiers  { get; init; }
    public uint VirtualKey { get; init; }

    /// <summary>The built-in default binding: Alt+Win+V.</summary>
    public static readonly HotkeyBinding Default = new()
    {
        Modifiers  = MOD_ALT | MOD_WIN,
        VirtualKey = (uint)KeyInterop.VirtualKeyFromKey(Key.V),
    };

    // ── Equality ─────────────────────────────────────────────────────────────

    public bool Equals(HotkeyBinding other) =>
        Modifiers == other.Modifiers && VirtualKey == other.VirtualKey;

    public override bool Equals(object? obj) =>
        obj is HotkeyBinding h && Equals(h);

    public override int GetHashCode() =>
        HashCode.Combine(Modifiers, VirtualKey);

    public static bool operator ==(HotkeyBinding left, HotkeyBinding right) => left.Equals(right);
    public static bool operator !=(HotkeyBinding left, HotkeyBinding right) => !left.Equals(right);

    // ── Formatting ───────────────────────────────────────────────────────────

    /// <summary>
    /// Returns a human-readable string such as <c>"Alt+Win+V"</c> or <c>"Ctrl+Shift+F5"</c>.
    /// Order is always: Ctrl, Alt, Shift, Win, Key.
    /// </summary>
    public override string ToString()
    {
        var parts = new List<string>(5);
        if ((Modifiers & MOD_CONTROL) != 0) parts.Add("Ctrl");
        if ((Modifiers & MOD_ALT)     != 0) parts.Add("Alt");
        if ((Modifiers & MOD_SHIFT)   != 0) parts.Add("Shift");
        if ((Modifiers & MOD_WIN)     != 0) parts.Add("Win");
        parts.Add(FormatKey(VirtualKey));
        return string.Join("+", parts);
    }

    /// <summary>
    /// Formats a modifier-flags value as a partial string, e.g. <c>"Ctrl+Alt"</c>.
    /// Returns an empty string when no modifiers are set.
    /// </summary>
    public static string FormatModifiers(uint mods)
    {
        var parts = new List<string>(4);
        if ((mods & MOD_CONTROL) != 0) parts.Add("Ctrl");
        if ((mods & MOD_ALT)     != 0) parts.Add("Alt");
        if ((mods & MOD_SHIFT)   != 0) parts.Add("Shift");
        if ((mods & MOD_WIN)     != 0) parts.Add("Win");
        return string.Join("+", parts);
    }

    // ── Parsing ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Tries to parse a string such as <c>"Alt+Win+V"</c> into a <see cref="HotkeyBinding"/>.
    /// Returns <see langword="false"/> if the string is null/empty, contains an unrecognised
    /// token, has no modifier, or has no non-modifier key.
    /// </summary>
    public static bool TryParse(string? s, out HotkeyBinding binding)
    {
        binding = default;
        if (string.IsNullOrWhiteSpace(s))
            return false;

        uint mods = 0;
        uint vk   = 0;

        foreach (var raw in s.Split('+'))
        {
            switch (raw.Trim())
            {
                case "Alt":   mods |= MOD_ALT;     break;
                case "Ctrl":  mods |= MOD_CONTROL;  break;
                case "Shift": mods |= MOD_SHIFT;    break;
                case "Win":   mods |= MOD_WIN;      break;
                default:
                    if (!TryParseKey(raw.Trim(), out vk))
                        return false;
                    break;
            }
        }

        if (vk == 0 || mods == 0)
            return false;

        binding = new HotkeyBinding { Modifiers = mods, VirtualKey = vk };
        return true;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string FormatKey(uint vk)
    {
        var key = KeyInterop.KeyFromVirtualKey((int)vk);
        // Display digits as "0"–"9" rather than "D0"–"D9".
        if (key is >= Key.D0 and <= Key.D9)
            return ((char)('0' + (key - Key.D0))).ToString();
        return key.ToString();
    }

    private static bool TryParseKey(string token, out uint vk)
    {
        // Direct Key enum name, e.g. "V", "F5", "D1".
        if (Enum.TryParse<Key>(token, out var key) && key != Key.None)
        {
            var code = KeyInterop.VirtualKeyFromKey(key);
            if (code != 0) { vk = (uint)code; return true; }
        }

        // Digit shorthand: "0"–"9" → Key.D0–Key.D9.
        if (token.Length == 1 && token[0] is >= '0' and <= '9')
        {
            key = Key.D0 + (token[0] - '0');
            vk  = (uint)KeyInterop.VirtualKeyFromKey(key);
            return true;
        }

        vk = 0;
        return false;
    }
}
