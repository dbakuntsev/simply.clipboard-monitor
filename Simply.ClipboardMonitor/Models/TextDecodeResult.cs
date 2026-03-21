using System.Text;

namespace Simply.ClipboardMonitor.Models;

/// <summary>Result of a text-decode attempt.</summary>
public sealed record TextDecodeResult(
    string?  Text,
    Encoding? DetectedEncoding,
    bool     Success,
    string?  FailureMessage = null);
