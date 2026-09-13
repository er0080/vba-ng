namespace VbaNg.Compiler.Syntax;

/// <summary>Source text with a line map, so offsets convert to the 1-based line and column diagnostics use.</summary>
public sealed class SourceText
{
    private readonly int[] lineStarts;

    public SourceText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        Text = text;
        lineStarts = ComputeLineStarts(text);
    }

    public string Text { get; }

    public int Length => Text.Length;

    public int LineCount => lineStarts.Length;

    public char this[int index] => Text[index];

    /// <summary>1-based line and column of an offset. Offsets past the end map to the last position.</summary>
    public (int Line, int Column) GetLinePosition(int offset)
    {
        offset = Math.Clamp(offset, 0, Text.Length);
        var index = Array.BinarySearch(lineStarts, offset);
        if (index < 0)
        {
            index = ~index - 1;
        }

        return (index + 1, offset - lineStarts[index] + 1);
    }

    /// <summary>
    /// Line terminators per MS-VBAL 3.2.1: CR LF, CR, LF, LS (U+2028), PS (U+2029).
    /// Returns the terminator length at <paramref name="index"/>, or 0.
    /// </summary>
    public static int LineTerminatorLength(string text, int index)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (index >= text.Length)
        {
            return 0;
        }

        return text[index] switch
        {
            '\r' => index + 1 < text.Length && text[index + 1] == '\n' ? 2 : 1,
            '\n' or '\x2028' or '\x2029' => 1,
            _ => 0,
        };
    }

    private static int[] ComputeLineStarts(string text)
    {
        var starts = new List<int> { 0 };
        var i = 0;
        while (i < text.Length)
        {
            var length = LineTerminatorLength(text, i);
            if (length > 0)
            {
                i += length;
                starts.Add(i);
            }
            else
            {
                i++;
            }
        }

        return [.. starts];
    }
}
