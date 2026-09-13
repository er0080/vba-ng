using System.Globalization;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The DateTime module of the VBA standard library (MS-VBAL 6.1.2, DateTime module). Dates are
/// serials (<see cref="VbaDate"/>); the calendar functions read the serial through the regional
/// Gregorian calendar, rounding the time of day to the nearest second as VBA does. Argument
/// checks and error numbers follow the DateTime goldens.
/// </summary>
public static class VbaDateTime
{
    private const int FirstDayDefault = 1; // vbSunday
    private static readonly DateTime Epoch = new(1899, 12, 30);

    public static VbaDate Now() => VbaDate.FromDateTime(DateTime.Now);

    public static VbaDate Today() => VbaDate.FromDateTime(DateTime.Today);

    public static VbaDate TimeOfDay() => VbaDate.FromDateTime(Epoch + TruncateToSecond(DateTime.Now.TimeOfDay));

    /// <summary>Timer: seconds since midnight as a Single.</summary>
    public static float Timer() => (float)DateTime.Now.TimeOfDay.TotalSeconds;

    /// <summary>
    /// DateSerial(Year, Month, Day): months and days outside their ranges roll over, two-digit
    /// years fall in 1930 to 2029, a year above 9999 raises 5, and Integer arguments overflow at
    /// 32767 (error 6).
    /// </summary>
    public static VbaDate DateSerial(in Variant year, in Variant month, in Variant day)
    {
        int y = Coerce.ToInt16(year);
        int m = Coerce.ToInt16(month);
        int d = Coerce.ToInt16(day);
        if (y is >= 0 and <= 99)
        {
            // Two-digit years follow the regional window, the same one date strings use (DateTime golden: 30 is 2030).
            y = Coerce.Culture.Calendar.ToFourDigitYear(y);
        }

        // Month first, from January of the year, then the day offset from the first of that month.
        var totalMonths = y * 12 + (m - 1);
        var calendarYear = Math.DivRem(totalMonths, 12, out var monthIndex);
        if (monthIndex < 0)
        {
            monthIndex += 12;
            calendarYear--;
        }

        if (calendarYear is < 1 or > 9999)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var firstOfMonth = new DateTime(calendarYear, monthIndex + 1, 1);
        var days = (firstOfMonth - Epoch).TotalDays + (d - 1);
        return FromSerialOrInvalid(days);
    }

    /// <summary>TimeSerial(Hour, Minute, Second): out-of-range parts roll over, and Integer arguments overflow at 32767.</summary>
    public static VbaDate TimeSerial(in Variant hour, in Variant minute, in Variant second)
    {
        int h = Coerce.ToInt16(hour);
        int n = Coerce.ToInt16(minute);
        int s = Coerce.ToInt16(second);
        var seconds = (h * 3600L) + (n * 60L) + s;
        return FromSerialOrInvalid(seconds / 86400.0);
    }

    /// <summary>DateAdd(Interval, Number, Date): the number truncates; month arithmetic clamps to the month's end; a result out of range raises 5.</summary>
    public static Variant DateAdd(in Variant interval, in Variant number, in Variant date)
    {
        var unit = Interval(interval);
        var count = Math.Truncate(Coerce.ToDouble(number));
        if (date.IsNull)
        {
            return Variant.Null;
        }

        var serial = Coerce.ToDate(date).Serial;
        var value = ToDateTime(serial);
        switch (unit)
        {
            case 'y':
                return Result(AddMonthsClamped(value, checked((int)count * 12)));
            case 'q':
                return Result(AddMonthsClamped(value, checked((int)count * 3)));
            case 'm':
                return Result(AddMonthsClamped(value, checked((int)count)));
            case 'd':
            case 'w':
            case 'j':
                return FromSerialOrInvalidVariant(serial + count);
            case 'k':
                return FromSerialOrInvalidVariant(serial + (count * 7));
            case 'h':
                return FromSerialOrInvalidVariant(serial + (count / 24));
            case 'n':
                return FromSerialOrInvalidVariant(serial + (count / 1440));
            default:
                return FromSerialOrInvalidVariant(serial + (count / 86400));
        }

        static Variant Result(DateTime value) => Variant.FromDate(VbaDate.FromDateTime(value));
    }

    /// <summary>DateDiff(Interval, Date1, Date2, [FirstDayOfWeek], [FirstWeekOfYear]): counts interval boundaries crossed.</summary>
    public static Variant DateDiff(in Variant interval, in Variant date1, in Variant date2, in Variant firstDayOfWeek, in Variant firstWeekOfYear)
    {
        var unit = Interval(interval);
        var firstDay = FirstDayOfWeek(firstDayOfWeek);
        _ = FirstWeekOfYear(firstWeekOfYear);
        if (date1.IsNull || date2.IsNull)
        {
            return Variant.Null;
        }

        var from = Coerce.ToDate(date1).Serial;
        var to = Coerce.ToDate(date2).Serial;
        var a = ToDateTime(from);
        var b = ToDateTime(to);
        long result = unit switch
        {
            'y' => b.Year - a.Year,
            'q' => ((b.Year - a.Year) * 4) + ((b.Month - 1) / 3) - ((a.Month - 1) / 3),
            'm' => ((b.Year - a.Year) * 12) + b.Month - a.Month,
            'd' or 'j' => WholeDays(to) - WholeDays(from),
            'w' => (long)Math.Truncate((WholeDays(to) - WholeDays(from)) / 7.0),
            'k' => (long)Math.Round((WeekStart(b.Date, firstDay) - WeekStart(a.Date, firstDay)).TotalDays / 7),
            // Hours, minutes, and seconds count the boundaries crossed, like days do: 23:59:59 to 00:00:00 is one hour.
            'h' => Boundaries(to, 3600) - Boundaries(from, 3600),
            'n' => Boundaries(to, 60) - Boundaries(from, 60),
            _ => Boundaries(to, 1) - Boundaries(from, 1),
        };
        // The result is a Variant: a Long when it fits, a Double beyond that (DateTime golden).
        return result < int.MinValue || result > int.MaxValue ? Variant.FromDouble(result) : Variant.FromInt32((int)result);
    }

    /// <summary>DatePart(Interval, Date, [FirstDayOfWeek], [FirstWeekOfYear]).</summary>
    public static Variant DatePart(in Variant interval, in Variant date, in Variant firstDayOfWeek, in Variant firstWeekOfYear)
    {
        var unit = Interval(interval);
        var firstDay = FirstDayOfWeek(firstDayOfWeek);
        var firstWeek = FirstWeekOfYear(firstWeekOfYear);
        if (date.IsNull)
        {
            return Variant.Null;
        }

        var value = ToDateTime(Coerce.ToDate(date).Serial);
        int result = unit switch
        {
            'y' => value.Year,
            'q' => ((value.Month - 1) / 3) + 1,
            'm' => value.Month,
            'j' => value.DayOfYear,
            'd' => value.Day,
            'w' => WeekdayNumber(value, firstDay),
            'k' => WeekNumber(value, firstDay, firstWeek),
            'h' => value.Hour,
            'n' => value.Minute,
            _ => value.Second,
        };
        return Variant.FromInt16((short)result);
    }

    public static Variant Year(in Variant date) => Part(date, static d => d.Year);

    public static Variant Month(in Variant date) => Part(date, static d => d.Month);

    public static Variant Day(in Variant date) => Part(date, static d => d.Day);

    public static Variant Hour(in Variant date) => Part(date, static d => d.Hour);

    public static Variant Minute(in Variant date) => Part(date, static d => d.Minute);

    public static Variant Second(in Variant date) => Part(date, static d => d.Second);

    /// <summary>Weekday(Date, [FirstDayOfWeek]): 1 to 7 counted from the first day of the week; 0 (vbUseSystemDayOfWeek) is Sunday.</summary>
    public static Variant Weekday(in Variant date, in Variant firstDayOfWeek)
    {
        var firstDay = FirstDayOfWeek(firstDayOfWeek);
        if (date.IsNull)
        {
            return Variant.Null;
        }

        return Variant.FromInt16((short)WeekdayNumber(ToDateTime(Coerce.ToDate(date).Serial), firstDay));
    }

    /// <summary>DateValue(String): the date part of a date string or Date value; numbers are not accepted (error 13).</summary>
    public static VbaDate DateValue(in Variant text)
    {
        var date = text.IsDate ? text.AsDate() : ParseDateText(text);
        return VbaDate.FromSerial(Math.Truncate(date.Serial));
    }

    /// <summary>TimeValue(String): the time part of a date string or Date value; "24:00:00" and a bare hour are not times (error 13).</summary>
    public static VbaDate TimeValue(in Variant text)
    {
        VbaDate date;
        if (text.IsDate)
        {
            date = text.AsDate();
        }
        else
        {
            var value = RequireString(text);
            if (!value.Contains(':', StringComparison.Ordinal) && !value.Contains("AM", StringComparison.OrdinalIgnoreCase) && !value.Contains("PM", StringComparison.OrdinalIgnoreCase))
            {
                throw VbaErrors.TypeMismatch();
            }

            date = ParseDateText(text);
        }

        // Seconds since midnight over a day, so the result is the same serial a time-only string gives.
        return VbaDate.FromSerial(ToDateTime(date.Serial).TimeOfDay.TotalSeconds / 86400.0);
    }

    private static VbaDate ParseDateText(in Variant text)
    {
        var value = RequireString(text);
        if (!DateText.TryParse(value, Coerce.Culture, out var date, allowNumber: false))
        {
            throw VbaErrors.TypeMismatch();
        }

        return date;
    }

    /// <summary>The calendar value of a serial with the time rounded to the nearest second, as the date functions read it.</summary>
    public static DateTime ToDateTime(double serial)
    {
        if (serial <= VbaDate.MinSerial - 1 || serial >= VbaDate.MaxSerial)
        {
            throw VbaErrors.TypeMismatch();
        }

        var days = Math.Truncate(serial);
        var seconds = Math.Round(Math.Abs(serial - days) * 86400);
        return Epoch.AddDays(days).AddSeconds(seconds);
    }

    private static string RequireString(in Variant text)
    {
        if (text.IsNull)
        {
            throw VbaErrors.InvalidUseOfNull();
        }

        return text.IsString ? text.AsString() : throw VbaErrors.TypeMismatch();
    }

    private static Variant Part(in Variant date, Func<DateTime, int> part)
    {
        if (date.IsNull)
        {
            return Variant.Null;
        }

        var serial = date.IsString ? Coerce.ToDate(date).Serial : Coerce.ToDouble(date);
        return Variant.FromInt16((short)part(ToDateTime(serial)));
    }

    /// <summary>The interval codes: yyyy, q, m, y (day of year), d, w (weekday), ww (week), h, n, s; anything else raises 5.</summary>
    private static char Interval(in Variant interval)
    {
        var text = Coerce.ToString(interval).Trim().ToUpperInvariant();
        return text switch
        {
            "YYYY" => 'y',
            "Q" => 'q',
            "M" => 'm',
            "Y" => 'j',
            "D" => 'd',
            "W" => 'w',
            "WW" => 'k',
            "H" => 'h',
            "N" => 'n',
            "S" => 's',
            _ => throw VbaErrors.InvalidProcedureCall(),
        };
    }

    private static DayOfWeek FirstDayOfWeek(in Variant argument)
    {
        var value = argument.IsMissing ? FirstDayDefault : Coerce.ToInt32(argument);
        if (value is < 0 or > 7)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        // 0 is vbUseSystemDayOfWeek: the regional first day, which is Sunday under en-US.
        return value == 0 ? Coerce.Culture.DateTimeFormat.FirstDayOfWeek : (DayOfWeek)(value - 1);
    }

    private static int FirstWeekOfYear(in Variant argument)
    {
        var value = argument.IsMissing ? 1 : Coerce.ToInt32(argument);
        if (value is < 0 or > 3)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return value == 0 ? 1 : value;
    }

    private static int WeekdayNumber(DateTime value, DayOfWeek firstDay) => (((int)value.DayOfWeek - (int)firstDay + 7) % 7) + 1;

    private static DateTime WeekStart(DateTime date, DayOfWeek firstDay) => date.AddDays(-(((int)date.DayOfWeek - (int)firstDay + 7) % 7));

    /// <summary>The week number under the first-week rule: 1 = the week of January 1, 2 = the first week with four days in the year, 3 = the first full week.</summary>
    private static int WeekNumber(DateTime value, DayOfWeek firstDay, int rule)
    {
        var date = value.Date;
        var start = WeekOneStart(date.Year, firstDay, rule);
        if (date < start)
        {
            start = WeekOneStart(date.Year - 1, firstDay, rule);
        }
        else if (date.Year < 9999 && date >= WeekOneStart(date.Year + 1, firstDay, rule))
        {
            return 1;
        }

        return (int)((date - start).TotalDays / 7) + 1;
    }

    private static DateTime WeekOneStart(int year, DayOfWeek firstDay, int rule)
    {
        var january1 = new DateTime(year, 1, 1);
        var weekStart = WeekStart(january1, firstDay);
        var daysInYear = 7 - (int)(january1 - weekStart).TotalDays;
        return rule switch
        {
            2 when daysInYear < 4 => weekStart.AddDays(7),
            3 when daysInYear < 7 => weekStart.AddDays(7),
            _ => weekStart,
        };
    }

    private static DateTime AddMonthsClamped(DateTime value, int months)
    {
        var totalMonths = (value.Year * 12) + (value.Month - 1) + months;
        var year = Math.DivRem(totalMonths, 12, out var monthIndex);
        if (monthIndex < 0)
        {
            monthIndex += 12;
            year--;
        }

        if (year is < 1 or > 9999)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var day = Math.Min(value.Day, DateTime.DaysInMonth(year, monthIndex + 1));
        return new DateTime(year, monthIndex + 1, day) + value.TimeOfDay;
    }

    private static long WholeDays(double serial) => (long)Math.Truncate(serial);

    /// <summary>The number of whole units of the given length in seconds since the epoch, with the time rounded to seconds first.</summary>
    private static long Boundaries(double serial, int unitSeconds)
    {
        var days = Math.Truncate(serial);
        var seconds = (long)Math.Round(Math.Abs(serial - days) * 86400);
        return (((long)days * 86400) + seconds) / unitSeconds;
    }

    private static TimeSpan TruncateToSecond(TimeSpan time) => TimeSpan.FromSeconds(Math.Truncate(time.TotalSeconds));

    private static VbaDate FromSerialOrInvalid(double serial)
    {
        try
        {
            return VbaDate.FromSerial(serial);
        }
        catch (VbaException)
        {
            throw VbaErrors.InvalidProcedureCall();
        }
    }

    private static Variant FromSerialOrInvalidVariant(double serial) => Variant.FromDate(FromSerialOrInvalid(serial));

    internal static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);
}
