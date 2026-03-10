using System.Collections;
using System.Text;

namespace Simply.ClipboardMonitor;

internal sealed class HexRowCollection(byte[] data) : IList
{
    internal const int BytesPerRow = 16;

    private readonly Dictionary<int, HexRow> _cache = [];

    public int Count => (data.Length + BytesPerRow - 1) / BytesPerRow;
    public bool IsReadOnly => true;
    public bool IsFixedSize => true;
    public bool IsSynchronized => false;
    public object SyncRoot => this;

    public object? this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (_cache.TryGetValue(index, out var cached))
            {
                return cached;
            }

            var offset = index * BytesPerRow;
            var rowLength = Math.Min(BytesPerRow, data.Length - offset);
            var hexBuilder = new StringBuilder();
            var asciiBuilder = new StringBuilder();

            for (var i = 0; i < BytesPerRow; i++)
            {
                if (i < rowLength)
                {
                    var b = data[offset + i];
                    hexBuilder.Append(b.ToString("X2"));
                    asciiBuilder.Append(b >= 32 && b <= 126 ? (char)b : '.');
                }
                else
                {
                    hexBuilder.Append("  ");
                }

                if (i != BytesPerRow - 1)
                {
                    hexBuilder.Append(' ');
                }
            }

            var row = new HexRow(offset.ToString("X8"), hexBuilder.ToString(), asciiBuilder.ToString());
            _cache[index] = row;
            return row;
        }
        set => throw new NotSupportedException();
    }

    public int Add(object? value) => throw new NotSupportedException();
    public void Clear() => throw new NotSupportedException();
    public bool Contains(object? value) => false;
    public int IndexOf(object? value) => -1;
    public void Insert(int index, object? value) => throw new NotSupportedException();
    public void Remove(object? value) => throw new NotSupportedException();
    public void RemoveAt(int index) => throw new NotSupportedException();
    public void CopyTo(Array array, int index) => throw new NotSupportedException();

    public IEnumerator GetEnumerator()
    {
        for (var i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }
}

internal sealed class HexRow(string offset, string hex, string ascii)
{
    public string Offset { get; } = offset;
    public string Hex { get; } = hex;
    public string Ascii { get; } = ascii;
}
