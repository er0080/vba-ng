namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Recognizes the text between the "#" delimiters of a date literal (MS-VBAL 3.3.3 date-or-time).
/// Only the shape is checked here; the value is computed by the runtime library in a later
/// milestone, with goldens.
/// </summary>
public static class DateLiteralSyntax
{
    private static readonly string[] MonthNames =
    [
        "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december",
        "jan", "feb", "mar", "apr", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
    ];

    /// <summary>date-or-time = (date-value 1*WSC time-value) / date-value / time-value.</summary>
    public static bool IsDateOrTime(ReadOnlySpan<char> content)
    {
        var parts = Split(content);
        if (parts.Count == 0)
        {
            return false;
        }

        // Try: date only, time only, then date followed by time at every split point.
        if (IsDateValue(parts, 0, parts.Count) || IsTimeValue(parts, 0, parts.Count))
        {
            return true;
        }

        for (var split = 1; split < parts.Count; split++)
        {
            if (parts[split].SeparatorBefore == Separator.Whitespace
                && IsDateValue(parts, 0, split)
                && IsTimeValue(parts, split, parts.Count))
            {
                return true;
            }
        }

        return false;
    }

    private enum Separator
    {
        None,
        Whitespace,
        Slash,
        Hyphen,
        Comma,
        Colon,
        Period,
    }

    private enum PartKind
    {
        Number,
        Month,
        AmPm,
    }

    private readonly record struct Part(PartKind Kind, Separator SeparatorBefore);

    private static List<Part> Split(ReadOnlySpan<char> content)
    {
        var parts = new List<Part>();
        var i = 0;
        var separator = Separator.None;
        while (i < content.Length)
        {
            var c = content[i];
            if (c is ' ' or '\t')
            {
                if (separator == Separator.None && parts.Count > 0)
                {
                    separator = Separator.Whitespace;
                }

                i++;
                continue;
            }

            var punctuation = c switch
            {
                '/' => Separator.Slash,
                '-' => Separator.Hyphen,
                ',' => Separator.Comma,
                ':' => Separator.Colon,
                '.' => Separator.Period,
                _ => Separator.None,
            };

            if (punctuation != Separator.None)
            {
                if (parts.Count == 0 || separator is not (Separator.None or Separator.Whitespace))
                {
                    return [];
                }

                separator = punctuation;
                i++;
                continue;
            }

            var start = i;
            if (char.IsAsciiDigit(c))
            {
                while (i < content.Length && char.IsAsciiDigit(content[i]))
                {
                    i++;
                }

                parts.Add(new Part(PartKind.Number, separator));
            }
            else if (char.IsAsciiLetter(c))
            {
                while (i < content.Length && char.IsAsciiLetter(content[i]))
                {
                    i++;
                }

                var word = content[start..i];
                if (word.Equals("am", StringComparison.OrdinalIgnoreCase) || word.Equals("pm", StringComparison.OrdinalIgnoreCase)
                    || word.Equals("a", StringComparison.OrdinalIgnoreCase) || word.Equals("p", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add(new Part(PartKind.AmPm, separator));
                }
                else if (IsMonthName(word))
                {
                    parts.Add(new Part(PartKind.Month, separator));
                }
                else
                {
                    return [];
                }
            }
            else
            {
                return [];
            }

            separator = Separator.None;
        }

        return separator is Separator.None or Separator.Whitespace ? parts : [];
    }

    private static bool IsMonthName(ReadOnlySpan<char> word)
    {
        foreach (var name in MonthNames)
        {
            if (word.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    // date-value = left-date-value date-separator middle-date-value [date-separator right-date-value];
    // each value is a decimal-literal or a month-name, at most one of them a month-name.
    private static bool IsDateValue(List<Part> parts, int start, int end)
    {
        var count = end - start;
        if (count is not (2 or 3))
        {
            return false;
        }

        var months = 0;
        for (var i = start; i < end; i++)
        {
            if (parts[i].Kind == PartKind.AmPm)
            {
                return false;
            }

            if (parts[i].Kind == PartKind.Month)
            {
                months++;
            }

            if (i > start && parts[i].SeparatorBefore is not (Separator.Whitespace or Separator.Slash or Separator.Hyphen or Separator.Comma))
            {
                return false;
            }
        }

        return months <= 1;
    }

    // time-value = (hour-value ampm) / (hour-value time-separator minute-value [time-separator second-value] [ampm]).
    private static bool IsTimeValue(List<Part> parts, int start, int end)
    {
        var count = end - start;
        if (count < 2 || parts[start].Kind != PartKind.Number)
        {
            return false;
        }

        if (count == 2 && parts[start + 1].Kind == PartKind.AmPm)
        {
            return true;
        }

        var i = start + 1;
        var fields = 1;
        while (i < end && parts[i].Kind == PartKind.Number && parts[i].SeparatorBefore is Separator.Colon or Separator.Period)
        {
            fields++;
            i++;
        }

        if (fields is not (2 or 3))
        {
            return false;
        }

        if (i < end && parts[i].Kind == PartKind.AmPm)
        {
            i++;
        }

        return i == end;
    }
}
