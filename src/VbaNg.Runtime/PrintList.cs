using System.Globalization;
using System.Text;

using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>
/// Builds the text of a Print output list (MS-VBAL 5.4.5.6 output-list), as Debug.Print and
/// Print # write it: numbers with a leading sign position and a trailing space, strings as they
/// are, a comma moving to the next 14-column print zone, Spc(n) and Tab(n) as in VBA.
/// Pending a Print # golden (ROADMAP.md backlog), these rules follow the documentation.
/// </summary>
public sealed class PrintList
{
    private const int ZoneWidth = 14;

    private readonly StringBuilder text = new();
    private readonly List<string> segments = [];

    /// <summary>The pieces of the list in order: each item, zone advance, Spc, and Tab as text. Print # wraps lines between them under a Width.</summary>
    internal IReadOnlyList<string> Segments => segments;

    public PrintList Item(in Variant value)
    {
        Append(Text(value));
        return this;
    }

    /// <summary>A comma: advance to the start of the next print zone.</summary>
    public PrintList Zone()
    {
        var column = text.Length % ZoneWidth;
        Append(new string(' ', ZoneWidth - column));
        return this;
    }

    public PrintList Spc(int count)
    {
        if (count > 0)
        {
            Append(new string(' ', count));
        }

        return this;
    }

    /// <summary>Tab with no argument moves to the next print zone.</summary>
    public PrintList Tab() => Zone();

    /// <summary>Tab(column): pads to the 1-based column, or starts a new line when the column is already passed.</summary>
    public PrintList Tab(int column)
    {
        if (column < 1)
        {
            column = 1;
        }

        Append(text.Length >= column
            ? "\n" + new string(' ', column - 1)
            : new string(' ', column - 1 - text.Length));
        return this;
    }

    public override string ToString() => text.ToString();

    private void Append(string piece)
    {
        text.Append(piece);
        segments.Add(piece);
    }

    private static string Text(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Empty:
                return string.Empty;
            case VarType.Null:
                return "Null";
            case VarType.String:
                return value.AsString();
            case VarType.Boolean:
                return value.AsBoolean() ? "True" : "False";
            case VarType.Date:
                return DateText.Format(value.AsDate(), Coerce.Culture);
            case VarType.Error:
                return "Error " + value.AsError().Number.ToString(CultureInfo.InvariantCulture);
            case VarType.Object:
            case VarType.Array:
            case VarType.UserDefinedType:
                throw VbaErrors.TypeMismatch();
            default:
                // Numbers print in Str form (a space or a minus first) followed by a space.
                return Coerce.ToString(Conversion.Str(value)) + " ";
        }
    }
}
