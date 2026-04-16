using Simply.ClipboardMonitor.Common;
using Simply.ClipboardMonitor.Models;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Simply.ClipboardMonitor.Services.Impl;

/// <summary>
/// Resolves the current clipboard owner by mapping the owner HWND to a process,
/// then eagerly reading its full image path (via QueryFullProcessImageName) and
/// command line (via NtQueryInformationProcess + PEB walk, with WOW64 support).
/// </summary>
internal sealed class ClipboardOwnerService : IClipboardOwnerService
{
    public ClipboardOwnerInfo? Resolve()
    {
        // Empty clipboard — nothing to show.
        if (NativeMethods.CountClipboardFormats() == 0)
            return null;

        // No owner HWND (possible if clipboard data was set without an owner window).
        var ownerHwnd = NativeMethods.GetClipboardOwner();
        if (ownerHwnd == IntPtr.Zero)
            return new ClipboardOwnerInfo("Owner: (unknown)", null);

        // Map HWND → PID.
        NativeMethods.GetWindowThreadProcessId(ownerHwnd, out uint pid);
        if (pid == 0)
            return new ClipboardOwnerInfo("Owner: (unknown)", null);

        // Try full access first (needed for command-line reading), then limited (path only).
        var hProcess = NativeMethods.OpenProcess(
            NativeMethods.PROCESS_QUERY_INFORMATION | NativeMethods.PROCESS_VM_READ,
            false, pid);
        bool hasVmRead = hProcess != IntPtr.Zero;

        if (!hasVmRead)
            hProcess = NativeMethods.OpenProcess(
                NativeMethods.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);

        // Cannot open at all.
        if (hProcess == IntPtr.Zero)
            return new ClipboardOwnerInfo(
                $"Owner: unknown process ({pid})",
                BuildTooltip(pid, null, "(access denied)", null, "(access denied)"));

        try
        {
            string? fullPath    = TryGetFullPath(hProcess, out string? pathError);
            string  processName = !string.IsNullOrEmpty(fullPath)
                                ? Path.GetFileName(fullPath)!
                                : $"unknown process ({pid})";

            string? cmdLine      = null;
            string? cmdLineError = hasVmRead ? null : "(access denied)";
            if (hasVmRead)
                cmdLine = TryGetCommandLine(hProcess, out cmdLineError);

            return new ClipboardOwnerInfo(
                $"Owner: {processName} ({pid})",
                BuildTooltip(pid, fullPath, pathError, cmdLine, cmdLineError));
        }
        finally
        {
            NativeMethods.CloseHandle(hProcess);
        }
    }

    // ── Tooltip builder ──────────────────────────────────────────────────────

    private static string BuildTooltip(uint pid,
        string? path,    string? pathError,
        string? cmdLine, string? cmdLineError)
    {
        var sb = new StringBuilder();
        sb.Append($"PID: {pid}");
        sb.Append($"\nPath: {path ?? pathError ?? "(unavailable)"}");
        sb.Append($"\nCommand line: {cmdLine ?? cmdLineError ?? "(unavailable)"}");
        return sb.ToString();
    }

    // ── Path ─────────────────────────────────────────────────────────────────

    private static string? TryGetFullPath(IntPtr hProcess, out string? error)
    {
        error = null;
        var buffer = new char[1024];
        uint size = (uint)buffer.Length;
        if (NativeMethods.QueryFullProcessImageName(hProcess, 0, buffer, ref size))
            return size > 0 ? new string(buffer, 0, (int)size) : null;

        int lastError = Marshal.GetLastWin32Error();
        error = lastError == 5 ? "(access denied)" : $"(error {lastError})";
        return null;
    }

    // ── Command line ─────────────────────────────────────────────────────────

    private static string? TryGetCommandLine(IntPtr hProcess, out string? error)
    {
        error = null;
        try
        {
            NativeMethods.IsWow64Process(hProcess, out bool isWow64);
            return isWow64
                ? ReadCommandLineWow64(hProcess, out error)
                : ReadCommandLineNative(hProcess, out error);
        }
        catch (Exception ex)
        {
            error = $"({ex.Message})";
            return null;
        }
    }

    /// <summary>Reads the command line of a native (same-bitness as the host) target process.</summary>
    private static string? ReadCommandLineNative(IntPtr hProcess, out string? error)
    {
        error = null;

        var pbi = new PROCESS_BASIC_INFORMATION();
        int status = NativeMethods.NtQueryInformationProcessBasic(
            hProcess, 0, ref pbi, Marshal.SizeOf<PROCESS_BASIC_INFORMATION>(), out _);
        if (status != 0) { error = $"(NT error 0x{status:X8})"; return null; }

        // x64 PEB layout: ProcessParameters pointer at PEB+0x20.
        var ppAddress = ReadPtr64(hProcess, pbi.PebBaseAddress + 0x20, out error);
        if (ppAddress == IntPtr.Zero) return null;

        // x64 RTL_USER_PROCESS_PARAMETERS: CommandLine UNICODE_STRING at offset 0x70.
        return ReadUnicodeString64(hProcess, ppAddress + 0x70, out error);
    }

    /// <summary>Reads the command line of a WOW64 (32-bit on 64-bit Windows) target process.</summary>
    private static string? ReadCommandLineWow64(IntPtr hProcess, out string? error)
    {
        error = null;

        // ProcessWow64Information (class 26) returns the address of the 32-bit PEB.
        IntPtr peb32Address = IntPtr.Zero;
        int status = NativeMethods.NtQueryInformationProcessWow64(
            hProcess, 26, ref peb32Address, IntPtr.Size, out _);
        if (status != 0 || peb32Address == IntPtr.Zero)
        {
            error = status != 0 ? $"(NT error 0x{status:X8})" : "(WOW64 PEB unavailable)";
            return null;
        }

        // x86 PEB layout: ProcessParameters pointer at PEB+0x10 (4-byte pointer).
        var ppAddress = ReadPtr32(hProcess, peb32Address + 0x10, out error);
        if (ppAddress == IntPtr.Zero) return null;

        // x86 RTL_USER_PROCESS_PARAMETERS: CommandLine UNICODE_STRING at offset 0x40.
        return ReadUnicodeString32(hProcess, ppAddress + 0x40, out error);
    }

    // ── Memory-reading helpers ───────────────────────────────────────────────

    private static IntPtr ReadPtr64(IntPtr hProcess, IntPtr address, out string? error)
    {
        var buf = new byte[8];
        if (!NativeMethods.ReadProcessMemory(hProcess, address, buf, 8, out int read) || read < 8)
        { error = $"(read error {Marshal.GetLastWin32Error()})"; return IntPtr.Zero; }
        error = null;
        return (IntPtr)BitConverter.ToInt64(buf, 0);
    }

    private static IntPtr ReadPtr32(IntPtr hProcess, IntPtr address, out string? error)
    {
        var buf = new byte[4];
        if (!NativeMethods.ReadProcessMemory(hProcess, address, buf, 4, out int read) || read < 4)
        { error = $"(read error {Marshal.GetLastWin32Error()})"; return IntPtr.Zero; }
        error = null;
        return (IntPtr)BitConverter.ToUInt32(buf, 0);
    }

    /// <summary>
    /// Reads a UNICODE_STRING from a 64-bit process.
    /// Layout: [Length:2][MaxLength:2][Padding:4][Buffer:8] = 16 bytes.
    /// </summary>
    private static string? ReadUnicodeString64(IntPtr hProcess, IntPtr address, out string? error)
    {
        var buf = new byte[16];
        if (!NativeMethods.ReadProcessMemory(hProcess, address, buf, 16, out int read) || read < 16)
        { error = $"(read error {Marshal.GetLastWin32Error()})"; return null; }

        ushort byteLength = BitConverter.ToUInt16(buf, 0);
        long   bufPtr     = BitConverter.ToInt64(buf, 8);
        return ReadUnicodeChars(hProcess, (IntPtr)bufPtr, byteLength, out error);
    }

    /// <summary>
    /// Reads a UNICODE_STRING from a 32-bit (WOW64) process.
    /// Layout: [Length:2][MaxLength:2][Buffer:4] = 8 bytes.
    /// </summary>
    private static string? ReadUnicodeString32(IntPtr hProcess, IntPtr address, out string? error)
    {
        var buf = new byte[8];
        if (!NativeMethods.ReadProcessMemory(hProcess, address, buf, 8, out int read) || read < 8)
        { error = $"(read error {Marshal.GetLastWin32Error()})"; return null; }

        ushort byteLength = BitConverter.ToUInt16(buf, 0);
        uint   bufPtr     = BitConverter.ToUInt32(buf, 4);
        return ReadUnicodeChars(hProcess, (IntPtr)bufPtr, byteLength, out error);
    }

    private static string? ReadUnicodeChars(IntPtr hProcess, IntPtr bufPtr,
        ushort byteLength, out string? error)
    {
        error = null;
        if (byteLength == 0 || bufPtr == IntPtr.Zero)
            return string.Empty;

        var charBytes = new byte[byteLength];
        if (!NativeMethods.ReadProcessMemory(hProcess, bufPtr, charBytes, byteLength, out int read)
            || read < byteLength)
        { error = $"(read error {Marshal.GetLastWin32Error()})"; return null; }

        return Encoding.Unicode.GetString(charBytes);
    }
}
