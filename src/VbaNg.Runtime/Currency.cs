using System.Globalization;

namespace VbaNg.Runtime;

/// <summary>
/// The VBA Currency type: a 64-bit integer scaled by 10,000 (MS-VBAL 2.1). Arithmetic that leaves
/// the range raises error 6, and results with more than four decimals round half to even
/// (MS-VBAL 5.5.1.2.1.1 Banker's rounding).
/// </summary>
public readonly struct Currency : IEquatable<Currency>, IComparable<Currency>
{
    public const long ScaleFactor = 10_000;

    public static readonly Currency Zero;
    public static readonly Currency MaxValue = new(long.MaxValue);
    public static readonly Currency MinValue = new(long.MinValue);

    private Currency(long scaled) => Scaled = scaled;

    /// <summary>The value multiplied by 10,000.</summary>
    public long Scaled { get; }

    public static Currency FromScaled(long scaled) => new(scaled);

    /// <summary>Exact when the value has at most four decimals; otherwise rounds half to even.</summary>
    public static Currency FromDecimal(decimal value)
    {
        var scaled = decimal.Round(value * ScaleFactor, 0, MidpointRounding.ToEven);
        if (scaled < long.MinValue || scaled > long.MaxValue)
        {
            throw VbaErrors.Overflow();
        }

        return new Currency((long)scaled);
    }

    /// <summary>
    /// Scales in floating point and rounds half to even, as OLE Automation's VarCyFromR8 does, so
    /// values such as 1.00005 (slightly above the midpoint in binary) round up.
    /// </summary>
    public static Currency FromDouble(double value)
    {
        var scaled = Math.Round(value * ScaleFactor, MidpointRounding.ToEven);
        if (double.IsNaN(scaled) || scaled < long.MinValue || scaled >= 9223372036854775808.0)
        {
            throw VbaErrors.Overflow();
        }

        return new Currency((long)scaled);
    }

    public static Currency FromInt64(long value)
    {
        if (value > long.MaxValue / ScaleFactor || value < long.MinValue / ScaleFactor)
        {
            throw VbaErrors.Overflow();
        }

        return new Currency(value * ScaleFactor);
    }

    /// <summary>The value as a Decimal with scale 4, which is how CDec of a Currency reads (Conversion golden).</summary>
    public decimal ToDecimal()
    {
        var magnitude = Scaled < 0 ? (ulong)(-(Scaled + 1)) + 1 : (ulong)Scaled;
        return new decimal((int)(uint)magnitude, (int)(uint)(magnitude >> 32), 0, Scaled < 0, 4);
    }

    public double ToDouble() => Scaled / (double)ScaleFactor;

    /// <summary>The whole part, truncated toward zero.</summary>
    public long Truncate() => Scaled / ScaleFactor;

    public static Currency Add(Currency left, Currency right)
    {
        try
        {
            return new Currency(checked(left.Scaled + right.Scaled));
        }
        catch (OverflowException)
        {
            throw VbaErrors.Overflow();
        }
    }

    public static Currency Subtract(Currency left, Currency right)
    {
        try
        {
            return new Currency(checked(left.Scaled - right.Scaled));
        }
        catch (OverflowException)
        {
            throw VbaErrors.Overflow();
        }
    }

    public static Currency Multiply(Currency left, Currency right) => FromDecimal(left.ToDecimal() * right.ToDecimal());

    /// <summary>Division by zero raises error 11; the quotient rounds to four decimals.</summary>
    public static Currency Divide(Currency left, Currency right)
    {
        if (right.Scaled == 0)
        {
            throw VbaErrors.DivisionByZero();
        }

        return FromDecimal(left.ToDecimal() / right.ToDecimal());
    }

    public static Currency Negate(Currency value)
    {
        if (value.Scaled == long.MinValue)
        {
            throw VbaErrors.Overflow();
        }

        return new Currency(-value.Scaled);
    }

    public bool Equals(Currency other) => Scaled == other.Scaled;

    public override bool Equals(object? obj) => obj is Currency other && Equals(other);

    public override int GetHashCode() => Scaled.GetHashCode();

    public int CompareTo(Currency other) => Scaled.CompareTo(other.Scaled);

    /// <summary>Invariant text with trailing zeros removed, for example "1.5" or "-0.0001".</summary>
    public override string ToString() => DecimalText.Normalize(ToDecimal()).ToString(CultureInfo.InvariantCulture);

    public static bool operator ==(Currency left, Currency right) => left.Equals(right);

    public static bool operator !=(Currency left, Currency right) => !left.Equals(right);

    public static bool operator <(Currency left, Currency right) => left.Scaled < right.Scaled;

    public static bool operator >(Currency left, Currency right) => left.Scaled > right.Scaled;

    public static bool operator <=(Currency left, Currency right) => left.Scaled <= right.Scaled;

    public static bool operator >=(Currency left, Currency right) => left.Scaled >= right.Scaled;
}

/// <summary>Helpers for decimal text the way VBA prints fixed-point values.</summary>
public static class DecimalText
{
    /// <summary>Drops trailing zeros from the scale: 1.5000 becomes 1.5, 2.0 becomes 2.</summary>
    public static decimal Normalize(decimal value) => value / 1.000000000000000000000000000000000m;
}
