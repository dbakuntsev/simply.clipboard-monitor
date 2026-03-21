using Simply.ClipboardMonitor.Models;

namespace Simply.ClipboardMonitor.Services;

/// <summary>Saves and loads clipboard snapshots to/from .clipdb files.</summary>
public interface IClipboardFileRepository
{
    /// <summary>
    /// Creates (or overwrites) the file at <paramref name="path"/> and writes
    /// every format in <paramref name="formats"/> into it.
    /// </summary>
    void Save(string path, IReadOnlyList<SavedClipboardFormat> formats);

    /// <summary>
    /// Opens the file at <paramref name="path"/> and returns all stored clipboard formats.
    /// </summary>
    List<SavedClipboardFormat> Load(string path);
}
