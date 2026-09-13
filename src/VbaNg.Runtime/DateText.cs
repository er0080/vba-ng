using System.Globalization;

namespace VbaNg.Runtime;

/// <summary>
/// Date text under the host's regional settings (MS-VBAL 5.5.1.2.4 Let-coercion to and from
/// String, Date rows; 5.5.1.2.3 Let-coercion to and from Date). Two-digit years follow the
/// regional two-digit-year window (1950 to 2049 on this Windows), and a day-first date is
/// accepted when month-first does not fit, as Excel does (DateTime and Conversion goldens).
/// </summary>
public static class DateText
{

    private static readonly DateTime Epoch = new(1899, 12, 30);

    private static readonly string[] DayFirstPatterns =
    [
        "d/M/yyyy", "d/M/yy", "d-M-yyyy", "d-M-yy", "d/M/yyyy H:mm:ss", "d/M/yyyy h:mm:ss tt", "d/M/yyyy H:mm", "d/M/yyyy h:mm tt",
    ];

    /// <summary>
    /// Interprets text as a date/time, a time, or a date, in that order of precedence, then as a
    /// number that is a serial. Returns false when none applies (the caller raises error 13).
    /// </summary>
    public static bool TryParse(string text, CultureInfo culture, out VbaDate date, bool allowNumber = true)
    {
        ArgumentNullException.ThrowIfNull(text);
        return TryParse(text.AsSpan(), culture, out date, allowNumber);
    }

    /// <summary>The same over a span, so a String's BSTR is read in place (ARCHITECTURE.md D20).</summary>
    public static bool TryParse(ReadOnlySpan<char> text, CultureInfo culture, out VbaDate date, bool allowNumber = true)
    {
        ArgumentNullException.ThrowIfNull(culture);
        date = VbaDate.Zero;
        var trimmed = text.Trim();
        if (trimmed.Length == 0)
        {
            return false;
        }

        var regional = culture;
        const DateTimeStyles styles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;
        var isNumber = NumberText.TryParse(trimmed, culture, out _, out var serial, out _);
        if (!isNumber && trimmed.IndexOf('.') < 0)
        {
            if (DateTime.TryParse(trimmed, regional, styles, out var parsed)
                || DateTime.TryParseExact(trimmed, DayFirstPatterns, regional, styles, out parsed))
            {
                return TryFromDateTime(parsed, out date);
            }
        }

        if (isNumber && allowNumber && !double.IsInfinity(serial))
        {
            try
            {
                date = VbaDate.FromSerial(serial);
                return true;
            }
            catch (VbaException)
            {
                return false;
            }
        }

        return false;
    }

    /// <summary>
    /// Short Date, Long Time, or both separated by a space: a whole day prints as a date only and a
    /// serial within 30 December 1899 prints as a time only (MS-VBAL 5.5.1.2.4, Date to String).
    /// </summary>
    public static string Format(VbaDate date, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);
        var value = date.ToDateTime();
        var formats = culture.DateTimeFormat;
        if (value.Date == Epoch)
        {
            return value.ToString(formats.LongTimePattern, culture);
        }

        var text = value.ToString(formats.ShortDatePattern, culture);
        return value.TimeOfDay == TimeSpan.Zero ? text : text + " " + value.ToString(formats.LongTimePattern, culture);
    }

    private static bool TryFromDateTime(DateTime parsed, out VbaDate date)
    {
        date = VbaDate.Zero;
        var timeOnly = parsed.Date == DateTime.MinValue.Date;
        var value = timeOnly ? Epoch + parsed.TimeOfDay : parsed;
        if (value.Year < 100)
        {
            return false;
        }

        date = VbaDate.FromDateTime(value);
        return true;
    }
}
