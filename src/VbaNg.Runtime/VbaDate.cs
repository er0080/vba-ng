using System.Diagnostics;

namespace VbaNg.Runtime;

/// <summary>
/// The VBA Date type: a serial number of days since 30 December 1899, with the time of day in
/// the fractional part (MS-VBAL 5.5.1.2.3 Let-coercion to and from Date). For negative serials
/// the fraction still counts forward from midnight, so -1.5 is 29 December 1899 at noon.
/// </summary>
[DebuggerDisplay("{DebugView,nq}")]
public readonly struct VbaDate : IEquatable<VbaDate>, IComparable<VbaDate>
{
    /// <summary>1 January 100.</summary>
    public const double MinSerial = -657434;

    /// <summary>The last instant of 31 December 9999.</summary>
    public const double MaxSerial = 2958466;

    private static readonly DateTime Epoch = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);

    private VbaDate(double serial) => Serial = serial;

    public static readonly VbaDate Zero;

    public double Serial { get; }

    /// <summary>Serials outside the Date range raise error 6 (MS-VBAL 5.5.1.2.3).</summary>
    public static VbaDate FromSerial(double serial)
    {
        if (double.IsNaN(serial) || serial <= MinSerial - 1 || serial >= MaxSerial)
        {
            throw VbaErrors.Overflow();
        }

        return new VbaDate(serial);
    }

    public static VbaDate FromDateTime(DateTime value)
    {
        var days = (value.Date - Epoch).TotalDays;
        var time = value.TimeOfDay.TotalDays;
        return new VbaDate(days < 0 ? days - time : days + time);
    }

    /// <summary>
    /// The calendar date and time this serial denotes, the time rounded to the nearest second as
    /// VBA reads a Date for text and for its parts: 2:30:00 PM on 11 September 2026 is a serial
    /// a fraction of a microsecond short of the second, and prints as 2:30:00 PM (Conversion and
    /// Format goldens; Second in DateTime).
    /// </summary>
    public DateTime ToDateTime()
    {
        var days = Math.Truncate(Serial);
        var seconds = Math.Round(Math.Abs(Serial - days) * 86400);
        return Epoch.AddDays(days).AddSeconds(seconds);
    }

    public bool Equals(VbaDate other) => Serial.Equals(other.Serial);

    public override bool Equals(object? obj) => obj is VbaDate other && Equals(other);

    public override int GetHashCode() => Serial.GetHashCode();

    public int CompareTo(VbaDate other) => Serial.CompareTo(other.Serial);

    public override string ToString() => ToDateTime().ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>What the locals window shows (ARCHITECTURE.md section 9): the Date as VBA's locals window writes it, its text between #s.</summary>
    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    internal string DebugView => "#" + DateText.Format(this, Coerce.Culture) + "#";

    public static bool operator ==(VbaDate left, VbaDate right) => left.Equals(right);

    public static bool operator !=(VbaDate left, VbaDate right) => !left.Equals(right);

    public static bool operator <(VbaDate left, VbaDate right) => left.Serial < right.Serial;

    public static bool operator >(VbaDate left, VbaDate right) => left.Serial > right.Serial;

    public static bool operator <=(VbaDate left, VbaDate right) => left.Serial <= right.Serial;

    public static bool operator >=(VbaDate left, VbaDate right) => left.Serial >= right.Serial;
}
