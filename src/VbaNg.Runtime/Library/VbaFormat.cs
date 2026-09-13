using System.Globalization;
using System.Text;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The Format function and its named formats, and the FormatNumber family (MS-VBAL 6.1.2,
/// Strings module). A user-defined format is numeric when it has digit placeholders and no date code before them, otherwise a
/// date format when it has date codes, otherwise a string format. Numbers round half away from
/// zero on their 15-digit decimal form; the details are pinned by the Format goldens.
/// </summary>
public static class VbaFormat
{
    private static readonly string[] NumberCodes = ["0", "#"];

    /// <summary>Format(Expression, [Format], [FirstDayOfWeek], [FirstWeekOfYear]).</summary>
    public static string Format(in Variant value, in Variant format, in Variant firstDayOfWeek, in Variant firstWeekOfYear)
    {
        var pattern = format.IsMissing ? string.Empty : Coerce.ToString(format);
        _ = firstDayOfWeek;
        _ = firstWeekOfYear;
        if (pattern.Length == 0)
        {
            return General(value);
        }

        switch (pattern.ToUpperInvariant())
        {
            case "GENERAL NUMBER":
                return IsNullOrEmptyValue(value) ? string.Empty : NumberText.Format(Coerce.ToDouble(NumberOf(value)), NumberText.DoubleDigits, Separator);
            case "CURRENCY":
                return IsNullOrEmptyValue(value) ? string.Empty : CurrencyText(Coerce.ToDecimal(NumberOf(value)), Culture.NumberFormat.CurrencyDecimalDigits, true, Culture.NumberFormat.CurrencyNegativePattern == 0, true);
            case "FIXED":
                return IsNullOrEmptyValue(value) ? string.Empty : Numeric(value, "0.00");
            case "STANDARD":
                return IsNullOrEmptyValue(value) ? string.Empty : Numeric(value, "#,##0.00");
            case "PERCENT":
                return IsNullOrEmptyValue(value) ? string.Empty : Numeric(value, "0.00%");
            case "SCIENTIFIC":
                return IsNullOrEmptyValue(value) ? string.Empty : Numeric(value, "0.00E+00");
            case "YES/NO":
                return IsNullOrEmptyValue(value) ? string.Empty : Coerce.ToDouble(NumberOf(value)) == 0 ? "No" : "Yes";
            case "TRUE/FALSE":
                return IsNullOrEmptyValue(value) ? string.Empty : Coerce.ToDouble(NumberOf(value)) == 0 ? "False" : "True";
            case "ON/OFF":
                return IsNullOrEmptyValue(value) ? string.Empty : Coerce.ToDouble(NumberOf(value)) == 0 ? "Off" : "On";
            case "GENERAL DATE":
                return value.IsNull ? string.Empty : DateText.Format(DateOf(value), Culture);
            case "LONG DATE":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString(Culture.DateTimeFormat.LongDatePattern, Culture);
            case "MEDIUM DATE":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString("dd-MMM-yy", Culture);
            case "SHORT DATE":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString(Culture.DateTimeFormat.ShortDatePattern, Culture);
            case "LONG TIME":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString(Culture.DateTimeFormat.LongTimePattern, Culture);
            case "MEDIUM TIME":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString("hh:mm tt", Culture);
            case "SHORT TIME":
                return value.IsNull ? string.Empty : DateOf(value).ToDateTime().ToString("HH:mm", Culture);
        }

        if (HasNumberCodes(pattern) && !DateCodeFirst(pattern))
        {
            return Numeric(value, pattern);
        }

        if (HasDateCodes(pattern))
        {
            return value.IsNull ? string.Empty : DateFormat(DateOf(value), pattern);
        }

        return StringFormat(value, pattern);
    }

    /// <summary>FormatNumber(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit], [UseParensForNegativeNumbers], [GroupDigits]).</summary>
    public static VbaString FormatNumber(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group) => VbaString.Temporary(FormatNumberText(value, digits, leadingDigit, parentheses, group));

    private static string FormatNumberText(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group)
    {
        if (value.IsNull)
        {
            return string.Empty;
        }

        var number = Coerce.ToDecimal(NumberOf(value));
        var places = TriState(digits, Culture.NumberFormat.NumberDecimalDigits, allowCount: true);
        var leading = TriState(leadingDigit, 1, allowCount: false) != 0;
        var parens = TriState(parentheses, 0, allowCount: false) != 0;
        var grouping = TriState(group, 1, allowCount: false) != 0;
        return NumberWithOptions(number, places, leading, parens, grouping, string.Empty, string.Empty);
    }

    public static VbaString FormatPercent(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group) => VbaString.Temporary(FormatPercentText(value, digits, leadingDigit, parentheses, group));

    private static string FormatPercentText(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group)
    {
        if (value.IsNull)
        {
            return string.Empty;
        }

        var number = Coerce.ToDecimal(NumberOf(value)) * 100;
        var places = TriState(digits, Culture.NumberFormat.NumberDecimalDigits, allowCount: true);
        var leading = TriState(leadingDigit, 1, allowCount: false) != 0;
        var parens = TriState(parentheses, 0, allowCount: false) != 0;
        var grouping = TriState(group, 1, allowCount: false) != 0;
        return NumberWithOptions(number, places, leading, parens, grouping, string.Empty, "%");
    }

    public static VbaString FormatCurrency(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group) => VbaString.Temporary(FormatCurrencyText(value, digits, leadingDigit, parentheses, group));

    private static string FormatCurrencyText(in Variant value, in Variant digits, in Variant leadingDigit, in Variant parentheses, in Variant group)
    {
        if (value.IsNull)
        {
            return string.Empty;
        }

        var number = Coerce.ToDecimal(NumberOf(value));
        var places = TriState(digits, Culture.NumberFormat.CurrencyDecimalDigits, allowCount: true);
        var leading = TriState(leadingDigit, 1, allowCount: false) != 0;
        var parens = TriState(parentheses, Culture.NumberFormat.CurrencyNegativePattern == 0 ? 1 : 0, allowCount: false) != 0;
        var grouping = TriState(group, 1, allowCount: false) != 0;
        return CurrencyText(number, places, leading, parens, grouping);
    }

    /// <summary>FormatDateTime(Date, [NamedFormat = vbGeneralDate]): 0 general, 1 long date, 2 short date, 3 long time, 4 short time; others raise 5.</summary>
    public static VbaString FormatDateTime(in Variant value, in Variant namedFormat) => VbaString.Temporary(FormatDateTimeText(value, namedFormat));

    private static string FormatDateTimeText(in Variant value, in Variant namedFormat)
    {
        var kind = namedFormat.IsMissing ? 0 : Coerce.ToInt32(namedFormat);
        if (kind is < 0 or > 4)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (value.IsNull)
        {
            return string.Empty;
        }

        var date = DateOf(value);
        return kind switch
        {
            1 => date.ToDateTime().ToString(Culture.DateTimeFormat.LongDatePattern, Culture),
            2 => date.ToDateTime().ToString(Culture.DateTimeFormat.ShortDatePattern, Culture),
            3 => date.ToDateTime().ToString(Culture.DateTimeFormat.LongTimePattern, Culture),
            4 => date.ToDateTime().ToString("HH:mm", Culture),
            _ => DateText.Format(date, Culture),
        };
    }

    /// <summary>MonthName(Month, [Abbreviate]): 1 to 12, else error 5.</summary>
    public static VbaString MonthName(in Variant month, in Variant abbreviate) => VbaString.Temporary(MonthNameText(month, abbreviate));

    private static string MonthNameText(in Variant month, in Variant abbreviate)
    {
        var index = Coerce.ToInt32(month);
        if (index is < 1 or > 12)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var short_ = !abbreviate.IsMissing && Coerce.ToBoolean(abbreviate);
        return short_ ? Culture.DateTimeFormat.GetAbbreviatedMonthName(index) : Culture.DateTimeFormat.GetMonthName(index);
    }

    /// <summary>WeekdayName(Weekday, [Abbreviate], [FirstDayOfWeek]): 1 to 7 counted from the first day of the week, else error 5.</summary>
    public static VbaString WeekdayName(in Variant weekday, in Variant abbreviate, in Variant firstDayOfWeek) => VbaString.Temporary(WeekdayNameText(weekday, abbreviate, firstDayOfWeek));

    private static string WeekdayNameText(in Variant weekday, in Variant abbreviate, in Variant firstDayOfWeek)
    {
        var index = Coerce.ToInt32(weekday);
        var first = firstDayOfWeek.IsMissing ? 0 : Coerce.ToInt32(firstDayOfWeek);
        if (index is < 1 or > 7 || first is < 0 or > 7)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var firstDay = first == 0 ? Culture.DateTimeFormat.FirstDayOfWeek : (DayOfWeek)(first - 1);
        var day = (DayOfWeek)(((int)firstDay + index - 1) % 7);
        var short_ = !abbreviate.IsMissing && Coerce.ToBoolean(abbreviate);
        return short_ ? Culture.DateTimeFormat.GetAbbreviatedDayName(day) : Culture.DateTimeFormat.GetDayName(day);
    }

    private static CultureInfo Culture => Coerce.Culture;

    private static string Separator => Culture.NumberFormat.NumberDecimalSeparator;

    private static bool IsNullOrEmptyValue(in Variant value) => value.IsNull || value.IsEmpty;

    /// <summary>Format with no pattern: strings as they are, Booleans as words, numbers like CStr, dates as General Date.</summary>
    private static string General(in Variant value) => value.Type switch
    {
        VarType.Null or VarType.Empty => string.Empty,
        VarType.String => value.AsString(),
        VarType.Boolean => value.AsBoolean() ? "True" : "False",
        VarType.Date => DateText.Format(value.AsDate(), Culture),
        _ => Coerce.ToString(value),
    };

    /// <summary>A String argument to a numeric or date format is read as a number when it is one.</summary>
    private static Variant NumberOf(in Variant value)
    {
        if (value.IsString)
        {
            return NumberText.TryParse(value.AsVbaString().Chars, Culture, out _, out var number, out _) ? Variant.FromDouble(number) : value;
        }

        return value;
    }

    private static VbaDate DateOf(in Variant value) => value.IsDate ? value.AsDate() : Coerce.ToDate(value);

    private static bool HasNumberCodes(string pattern)
    {
        var bare = StripLiterals(pattern);
        return NumberCodes.Any(code => bare.Contains(code, StringComparison.Ordinal));
    }

    /// <summary>A date code comes before the first digit placeholder, which makes the pattern a date format, its zeros text (Format golden: a Date or 37623.5 with "yyyy-mm-ddTHH:mm:ss.000Z" is 2003-01-02T..., "0.00 yyyy" formats 1.5 as "1.50 yyyy").</summary>
    private static bool DateCodeFirst(string pattern)
    {
        foreach (var c in StripLiterals(pattern))
        {
            switch (char.ToUpperInvariant(c))
            {
                case '0' or '#':
                    return false;
                case 'Y' or 'M' or 'D' or 'H' or 'N' or 'S' or 'Q' or 'W' or 'T' or 'C':
                    return true;
            }
        }

        return false;
    }

    private static bool HasDateCodes(string pattern)
    {
        var bare = StripLiterals(pattern).ToUpperInvariant();
        return bare.Any(c => c is 'Y' or 'M' or 'D' or 'H' or 'N' or 'S' or 'Q' or 'W' or 'T' or 'C');
    }

    private static string StripLiterals(string pattern)
    {
        var bare = new StringBuilder();
        for (var i = 0; i < pattern.Length; i++)
        {
            switch (pattern[i])
            {
                case '\\':
                    i++;
                    break;
                case '"':
                    var close = pattern.IndexOf('"', i + 1);
                    i = close < 0 ? pattern.Length : close;
                    break;
                default:
                    bare.Append(pattern[i]);
                    break;
            }
        }

        return bare.ToString();
    }

    private static string[] Sections(string pattern)
    {
        var sections = new List<string>();
        var current = new StringBuilder();
        for (var i = 0; i < pattern.Length; i++)
        {
            var c = pattern[i];
            if (c == '\\' && i + 1 < pattern.Length)
            {
                current.Append(c).Append(pattern[++i]);
            }
            else if (c == '"')
            {
                var close = pattern.IndexOf('"', i + 1);
                var end = close < 0 ? pattern.Length - 1 : close;
                current.Append(pattern, i, end - i + 1);
                i = end;
            }
            else if (c == ';')
            {
                sections.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        sections.Add(current.ToString());
        return sections.ToArray();
    }

    private static string Numeric(in Variant value, string pattern)
    {
        var sections = Sections(pattern);
        if (value.IsNull || value.IsEmpty)
        {
            return sections.Length >= 4 ? Literal(sections[3]) : string.Empty;
        }

        var argument = NumberOf(value);
        if (argument.IsString)
        {
            return argument.AsString();
        }

        var number = Coerce.ToDecimal(argument);
        var negative = number < 0;
        var section = 0;
        var showSign = true;
        if (negative && sections.Length >= 2)
        {
            section = 1;
            showSign = false;
        }
        else if (number == 0 && sections.Length >= 3)
        {
            section = 2;
        }

        var text = NumericSection(Math.Abs(number), sections[section]);
        return negative && showSign && text.Length > 0 ? "-" + text : text;
    }

    private static string Literal(string section)
    {
        var text = new StringBuilder();
        for (var i = 0; i < section.Length; i++)
        {
            switch (section[i])
            {
                case '\\' when i + 1 < section.Length:
                    text.Append(section[++i]);
                    break;
                case '"':
                    var close = section.IndexOf('"', i + 1);
                    var end = close < 0 ? section.Length : close;
                    text.Append(section, i + 1, end - i - 1);
                    i = end;
                    break;
                default:
                    text.Append(section[i]);
                    break;
            }
        }

        return text.ToString();
    }

    private readonly record struct NumericLayout(
        List<(bool Digit, bool Required, string? Text, bool Fraction)> Items,
        int IntegerPlaceholders,
        int FractionPlaceholders,
        bool Grouping,
        int ScaleCommas,
        int PercentSigns,
        bool HasDecimal,
        char Exponent,
        bool ExponentSign,
        int ExponentDigits);

    private static NumericLayout Parse(string section)
    {
        var items = new List<(bool Digit, bool Required, string? Text, bool Fraction)>();
        var integerPlaceholders = 0;
        var fractionPlaceholders = 0;
        var grouping = false;
        var scaleCommas = 0;
        var percent = 0;
        var inFraction = false;
        var sawDigit = false;
        var exponent = '\0';
        var exponentSign = false;
        var exponentDigits = 0;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            switch (c)
            {
                case '0':
                case '#':
                    if (exponent != '\0')
                    {
                        exponentDigits++;
                        break;
                    }

                    // A comma right before this digit was grouping, not scaling.
                    if (scaleCommas > 0 && !inFraction)
                    {
                        grouping = true;
                        scaleCommas = 0;
                    }

                    items.Add((true, c == '0', null, inFraction));
                    if (inFraction)
                    {
                        fractionPlaceholders++;
                    }
                    else
                    {
                        integerPlaceholders++;
                    }

                    sawDigit = true;
                    break;
                case '.':
                    if (!inFraction)
                    {
                        inFraction = true;
                        items.Add((false, false, ".", false));
                    }

                    break;
                case ',':
                    // Between digits a comma groups thousands; after the last digit each comma divides by 1000.
                    if (sawDigit && !inFraction)
                    {
                        scaleCommas++;
                    }
                    else
                    {
                        items.Add((false, false, ",", inFraction));
                    }

                    break;
                case '%':
                    percent++;
                    items.Add((false, false, "%", inFraction));
                    break;
                case 'E':
                case 'e':
                    if (i + 1 < section.Length && section[i + 1] is '+' or '-' && exponent == '\0')
                    {
                        exponent = c;
                        exponentSign = section[i + 1] == '+';
                        i++;
                        break;
                    }

                    items.Add((false, false, c.ToString(), inFraction));
                    break;
                case '\\':
                    if (i + 1 < section.Length)
                    {
                        items.Add((false, false, section[++i].ToString(), inFraction));
                    }

                    break;
                case '"':
                    {
                        var close = section.IndexOf('"', i + 1);
                        var end = close < 0 ? section.Length : close;
                        items.Add((false, false, section[(i + 1)..end], inFraction));
                        i = end;
                        break;
                    }

                default:
                    items.Add((false, false, c.ToString(), inFraction));
                    break;
            }
        }

        // Commas that were followed by more digits count as grouping only; trailing ones scale by 1000.
        return new NumericLayout(items, integerPlaceholders, fractionPlaceholders, grouping && integerPlaceholders > 0, scaleCommas, percent, inFraction, exponent, exponentSign, exponentDigits);
    }

    private static string NumericSection(decimal magnitude, string section)
    {
        var layout = Parse(section);
        var number = magnitude;
        for (var i = 0; i < layout.PercentSigns; i++)
        {
            number *= 100;
        }

        for (var i = 0; i < layout.ScaleCommas; i++)
        {
            number /= 1000;
        }

        var exponentValue = 0;
        if (layout.Exponent != '\0')
        {
            // Scale so the integer part fills the integer placeholders.
            if (number != 0)
            {
                var digitsBefore = number >= 1 ? (int)Math.Floor(Math.Log10((double)number)) + 1 : (int)Math.Floor(Math.Log10((double)number)) + 1;
                exponentValue = digitsBefore - Math.Max(layout.IntegerPlaceholders, 1);
                number /= Power10(exponentValue);
            }
        }

        number = decimal.Round(number, Math.Min(layout.FractionPlaceholders, 28), MidpointRounding.AwayFromZero);
        if (layout.Exponent != '\0' && number >= Power10(Math.Max(layout.IntegerPlaceholders, 1)))
        {
            // Rounding carried into a new digit: renormalize.
            number /= 10;
            exponentValue++;
        }

        var text = number.ToString("F" + layout.FractionPlaceholders.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        var point = text.IndexOf('.', StringComparison.Ordinal);
        var integerDigits = point < 0 ? text : text[..point];
        var fractionDigits = point < 0 ? string.Empty : text[(point + 1)..];
        if (integerDigits == "0" && layout.IntegerPlaceholders > 0 && !layout.Items.Any(i => i.Digit && !i.Fraction && i.Required) && number < 1)
        {
            integerDigits = string.Empty;
        }
        else if (integerDigits == "0" && layout.IntegerPlaceholders == 0)
        {
            integerDigits = string.Empty;
        }

        // Fraction: a # placeholder drops a trailing zero.
        var requiredFraction = 0;
        var seen = 0;
        foreach (var item in layout.Items.Where(i => i.Digit && i.Fraction))
        {
            seen++;
            if (item.Required)
            {
                requiredFraction = seen;
            }
        }

        fractionDigits = fractionDigits.TrimEnd('0');
        if (fractionDigits.Length < requiredFraction)
        {
            fractionDigits = fractionDigits.PadRight(requiredFraction, '0');
        }

        // Integer part: the grouped digit text is dealt to the placeholders from the right, one
        // character group per placeholder, and the leftmost placeholder takes whatever remains.
        var integerText = layout.Grouping ? Group(integerDigits) : integerDigits;
        var pieces = new string?[layout.Items.Count];
        var position = integerText.Length;
        var firstIntegerSlot = -1;
        for (var index = layout.Items.Count - 1; index >= 0; index--)
        {
            var item = layout.Items[index];
            if (!item.Digit || item.Fraction)
            {
                continue;
            }

            firstIntegerSlot = index;
            if (position > 0)
            {
                var start = position - 1;
                if (!char.IsAsciiDigit(integerText[start]) && start > 0)
                {
                    // Keep a group separator with the digit before it.
                    start--;
                }

                pieces[index] = integerText[start..position];
                position = start;
            }
            else
            {
                pieces[index] = item.Required ? "0" : string.Empty;
            }
        }

        if (firstIntegerSlot >= 0 && position > 0)
        {
            pieces[firstIntegerSlot] = integerText[..position] + pieces[firstIntegerSlot];
        }

        var result = new StringBuilder();
        var fractionIndex = 0;
        for (var index = 0; index < layout.Items.Count; index++)
        {
            var item = layout.Items[index];
            if (!item.Digit)
            {
                if (item.Text == ".")
                {
                    if (fractionDigits.Length > 0 || layout.Items.Any(i => i.Digit && i.Fraction && i.Required))
                    {
                        result.Append(Separator);
                    }
                }
                else
                {
                    result.Append(item.Text);
                }

                continue;
            }

            if (item.Fraction)
            {
                if (fractionIndex < fractionDigits.Length)
                {
                    result.Append(fractionDigits[fractionIndex++]);
                }
                else if (item.Required)
                {
                    result.Append('0');
                }

                continue;
            }

            result.Append(pieces[index]);
        }

        if (layout.Exponent != '\0')
        {
            result.Append(layout.Exponent);
            if (exponentValue < 0)
            {
                result.Append('-');
            }
            else if (layout.ExponentSign)
            {
                result.Append('+');
            }

            result.Append(Math.Abs(exponentValue).ToString(new string('0', Math.Max(layout.ExponentDigits, 1)), CultureInfo.InvariantCulture));
        }

        return result.ToString();
    }

    private static decimal Power10(int exponent)
    {
        decimal result = 1;
        for (var i = 0; i < Math.Abs(exponent); i++)
        {
            result *= 10;
        }

        return exponent < 0 ? 1 / result : result;
    }

    private static string NumberWithOptions(decimal number, int places, bool leadingDigit, bool parentheses, bool grouping, string prefix, string suffix)
    {
        var negative = number < 0;
        var rounded = decimal.Round(Math.Abs(number), places, MidpointRounding.AwayFromZero);
        var text = rounded.ToString("F" + places.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        var point = text.IndexOf('.', StringComparison.Ordinal);
        var integerPart = point < 0 ? text : text[..point];
        var fraction = point < 0 ? string.Empty : text[(point + 1)..];
        if (grouping)
        {
            integerPart = Group(integerPart);
        }

        if (!leadingDigit && integerPart == "0")
        {
            integerPart = string.Empty;
        }

        var body = prefix + integerPart + (places > 0 ? Separator + fraction : string.Empty) + suffix;
        if (!negative || rounded == 0)
        {
            return body;
        }

        return parentheses ? "(" + body + ")" : Culture.NumberFormat.NegativeSign + body;
    }

    private static string CurrencyText(decimal number, int places, bool leadingDigit, bool parentheses, bool grouping) =>
        NumberWithOptions(number, places, leadingDigit, parentheses, grouping, Culture.NumberFormat.CurrencySymbol, string.Empty);

    private static string Group(string digits)
    {
        if (digits.Length <= 3)
        {
            return digits;
        }

        var grouped = new StringBuilder();
        for (var i = 0; i < digits.Length; i++)
        {
            if (i > 0 && (digits.Length - i) % 3 == 0)
            {
                grouped.Append(Culture.NumberFormat.NumberGroupSeparator);
            }

            grouped.Append(digits[i]);
        }

        return grouped.ToString();
    }

    /// <summary>A tri-state argument (-1 True, 0 False, -2 default); for digit counts any non-negative value is a count.</summary>
    private static int TriState(in Variant argument, int defaultValue, bool allowCount)
    {
        if (argument.IsMissing)
        {
            return defaultValue;
        }

        var value = Coerce.ToInt32(argument);
        if (allowCount)
        {
            return value < 0 ? defaultValue : value;
        }

        return value switch
        {
            -2 => defaultValue,
            0 => 0,
            _ => 1,
        };
    }

    private static string DateFormat(VbaDate date, string pattern)
    {
        var value = VbaDateTime.ToDateTime(date.Serial);
        var upper = StripLiterals(pattern).ToUpperInvariant();
        var twelveHour = upper.Contains("AM/PM", StringComparison.Ordinal) || upper.Contains("A/P", StringComparison.Ordinal) || upper.Contains("AMPM", StringComparison.Ordinal);
        var result = new StringBuilder();
        var lastWasHour = false;
        for (var i = 0; i < pattern.Length;)
        {
            var c = pattern[i];
            var rest = pattern.AsSpan(i);
            if (c == '\\' && i + 1 < pattern.Length)
            {
                result.Append(pattern[i + 1]);
                i += 2;
                continue;
            }

            if (c == '"')
            {
                var close = pattern.IndexOf('"', i + 1);
                var end = close < 0 ? pattern.Length : close;
                result.Append(pattern, i + 1, end - i - 1);
                i = end + 1;
                continue;
            }

            if (c == '[')
            {
                // A bracketed code produces nothing in VBA's Format ("[h]:nn" gives ":04", Format golden).
                var close = pattern.IndexOf(']', i + 1);
                i = close < 0 ? pattern.Length : close + 1;
                continue;
            }

            if (StartsWith(rest, "AM/PM"))
            {
                result.Append(value.Hour < 12 ? "AM" : "PM");
                i += 5;
                continue;
            }

            if (StartsWith(rest, "am/pm"))
            {
                result.Append(value.Hour < 12 ? "am" : "pm");
                i += 5;
                continue;
            }

            if (StartsWith(rest, "A/P"))
            {
                result.Append(value.Hour < 12 ? "A" : "P");
                i += 3;
                continue;
            }

            if (StartsWith(rest, "a/p"))
            {
                result.Append(value.Hour < 12 ? "a" : "p");
                i += 3;
                continue;
            }

            if (StartsWith(rest, "AMPM", ignoreCase: true))
            {
                result.Append(value.Hour < 12 ? Culture.DateTimeFormat.AMDesignator : Culture.DateTimeFormat.PMDesignator);
                i += 4;
                continue;
            }

            var run = 1;
            var code = char.ToUpperInvariant(c);
            while (i + run < pattern.Length && char.ToUpperInvariant(pattern[i + run]) == code)
            {
                run++;
            }

            switch (code)
            {
                case 'Y':
                    if (run == 1)
                    {
                        result.Append(value.DayOfYear.ToString(CultureInfo.InvariantCulture));
                    }
                    else if (run <= 2)
                    {
                        result.Append((value.Year % 100).ToString("00", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        result.Append(value.Year.ToString("0000", CultureInfo.InvariantCulture));
                    }

                    lastWasHour = false;
                    break;
                case 'Q':
                    result.Append(((value.Month - 1) / 3 + 1).ToString(CultureInfo.InvariantCulture));
                    lastWasHour = false;
                    break;
                case 'M':
                    if (lastWasHour && run <= 2)
                    {
                        result.Append(value.Minute.ToString(run == 2 ? "00" : "0", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        result.Append(run switch
                        {
                            1 => value.Month.ToString(CultureInfo.InvariantCulture),
                            2 => value.Month.ToString("00", CultureInfo.InvariantCulture),
                            3 => Culture.DateTimeFormat.GetAbbreviatedMonthName(value.Month),
                            _ => Culture.DateTimeFormat.GetMonthName(value.Month),
                        });
                    }

                    lastWasHour = false;
                    break;
                case 'D':
                    result.Append(run switch
                    {
                        1 => value.Day.ToString(CultureInfo.InvariantCulture),
                        2 => value.Day.ToString("00", CultureInfo.InvariantCulture),
                        3 => Culture.DateTimeFormat.GetAbbreviatedDayName(value.DayOfWeek),
                        4 => Culture.DateTimeFormat.GetDayName(value.DayOfWeek),
                        5 => value.ToString(Culture.DateTimeFormat.ShortDatePattern, Culture),
                        _ => value.ToString(Culture.DateTimeFormat.LongDatePattern, Culture),
                    });
                    lastWasHour = false;
                    break;
                case 'W':
                    result.Append(run == 1
                        ? ((int)value.DayOfWeek + 1).ToString(CultureInfo.InvariantCulture)
                        : WeekOfYear(value).ToString(CultureInfo.InvariantCulture));
                    lastWasHour = false;
                    break;
                case 'H':
                    {
                        var hour = twelveHour ? (value.Hour % 12 == 0 ? 12 : value.Hour % 12) : value.Hour;
                        result.Append(hour.ToString(run >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                        lastWasHour = true;
                        break;
                    }

                case 'N':
                    result.Append(value.Minute.ToString(run >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                    lastWasHour = false;
                    break;
                case 'S':
                    result.Append(value.Second.ToString(run >= 2 ? "00" : "0", CultureInfo.InvariantCulture));
                    lastWasHour = false;
                    break;
                case 'T' when run >= 5:
                    result.Append(value.ToString(Culture.DateTimeFormat.LongTimePattern, Culture));
                    lastWasHour = false;
                    break;
                case 'C':
                    result.Append(DateText.Format(date, Culture));
                    lastWasHour = false;
                    break;
                case '/':
                    result.Append(Culture.DateTimeFormat.DateSeparator);
                    break;
                case ':':
                    result.Append(Culture.DateTimeFormat.TimeSeparator);
                    break;
                default:
                    result.Append(c, run);
                    break;
            }

            i += run;
        }

        return result.ToString();
    }

    private static bool StartsWith(ReadOnlySpan<char> text, string prefix, bool ignoreCase = false) =>
        text.StartsWith(prefix, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    private static int WeekOfYear(DateTime value)
    {
        var january1 = new DateTime(value.Year, 1, 1);
        var weekStart = january1.AddDays(-(int)january1.DayOfWeek);
        return (int)((value.Date - weekStart).TotalDays / 7) + 1;
    }

    private static string StringFormat(in Variant value, string pattern)
    {
        var sections = Sections(pattern);
        var text = General(value);
        var section = text.Length == 0 && sections.Length >= 2 ? sections[1] : sections[0];
        var leftToRight = false;
        var upper = false;
        var lower = false;
        var placeholders = new List<(char Kind, string? Literal)>();
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            switch (c)
            {
                case '!':
                    leftToRight = true;
                    break;
                case '<':
                    lower = true;
                    break;
                case '>':
                    upper = true;
                    break;
                case '@':
                case '&':
                    placeholders.Add((c, null));
                    break;
                case '\\' when i + 1 < section.Length:
                    placeholders.Add(('\0', section[++i].ToString()));
                    break;
                case '"':
                    {
                        var close = section.IndexOf('"', i + 1);
                        var end = close < 0 ? section.Length : close;
                        placeholders.Add(('\0', section[(i + 1)..end]));
                        i = end;
                        break;
                    }

                default:
                    placeholders.Add(('\0', c.ToString()));
                    break;
            }
        }

        if (upper)
        {
            text = Strings.MapCase(text, upper: true);
        }
        else if (lower)
        {
            text = Strings.MapCase(text, upper: false);
        }

        var slots = placeholders.Count(p => p.Kind != '\0');
        if (slots == 0)
        {
            return string.Concat(placeholders.Select(p => p.Literal)) + text;
        }

        // Assign characters to slots: with more characters than slots the last slot takes the rest;
        // with fewer, the slots fill from the right unless "!" fills from the left.
        var assigned = new string[slots];
        if (text.Length >= slots)
        {
            for (var s = 0; s < slots - 1; s++)
            {
                assigned[s] = text[s].ToString();
            }

            assigned[slots - 1] = text[(slots - 1)..];
        }
        else
        {
            var offset = leftToRight ? 0 : slots - text.Length;
            for (var s = 0; s < slots; s++)
            {
                var index = s - offset;
                assigned[s] = index >= 0 && index < text.Length ? text[index].ToString() : string.Empty;
            }
        }

        var result = new StringBuilder();
        var slot = 0;
        foreach (var (kind, literal) in placeholders)
        {
            if (kind == '\0')
            {
                result.Append(literal);
                continue;
            }

            var piece = assigned[slot++];
            result.Append(piece.Length == 0 && kind == '@' ? " " : piece);
        }

        return result.ToString();
    }
}
