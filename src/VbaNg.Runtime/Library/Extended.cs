using System.Numerics;

namespace VbaNg.Runtime.Library;

/// <summary>
/// A number in the x87 extended format: a 64-bit significand, rounded to nearest even after every
/// operation. VBA's Financial functions evaluate in it (Financial golden: NPV, Rate, and IRR match
/// VBA bit for bit only under it), so they compute in this software version of it. Only what they
/// use exists; the exponent range is left unbounded, and a result that is not finite (a division
/// by zero) carries on as the Double it becomes.
/// </summary>
internal readonly struct Extended
{
    private const ulong Half = 1UL << 63;

    // A finite value is significand * 2^exponent, bit 63 of the significand set, or zero with a zero significand.
    private readonly ulong significand;
    private readonly int exponent;
    private readonly bool negative;

    // Infinity or NaN, carried as its Double.
    private readonly bool special;
    private readonly double specialValue;

    private Extended(ulong significand, int exponent, bool negative)
    {
        this.significand = significand;
        this.exponent = exponent;
        this.negative = negative;
    }

    private Extended(double specialValue)
    {
        special = true;
        this.specialValue = specialValue;
    }

    public static Extended One => new(Half, -63, false);

    public bool IsNegative => special ? specialValue < 0 : negative && significand != 0;

    public bool IsZero => !special && significand == 0;

    public static implicit operator Extended(double value)
    {
        if (!double.IsFinite(value))
        {
            return new Extended(value);
        }

        var bits = BitConverter.DoubleToInt64Bits(value);
        var fraction = (ulong)bits & 0xF_FFFF_FFFF_FFFF;
        var biased = (int)((bits >> 52) & 0x7FF);
        if (biased == 0 && fraction == 0)
        {
            return new Extended(0, 0, bits < 0);
        }

        var exponent = biased == 0 ? -1074 : biased - 1075;
        if (biased != 0)
        {
            fraction |= 1UL << 52;
        }

        var shift = BitOperations.LeadingZeroCount(fraction);
        return new Extended(fraction << shift, exponent - shift, bits < 0);
    }

    /// <summary>The Double this value stores as (FST): rounded to nearest even once more, to 53 bits.</summary>
    public double ToDouble()
    {
        if (special)
        {
            return specialValue;
        }

        if (significand == 0)
        {
            return negative ? -0.0 : 0.0;
        }

        var top = exponent + 63;
        if (top > 1023)
        {
            return negative ? double.NegativeInfinity : double.PositiveInfinity;
        }

        // 11 bits go in a normal Double; more below its smallest normal exponent.
        var drop = top >= -1022 ? 11 : 11 + (-1022 - top);
        if (drop > 64)
        {
            return negative ? -0.0 : 0.0;
        }

        ulong kept;
        bool up;
        if (drop == 64)
        {
            kept = 0;
            up = significand > Half;
        }
        else
        {
            kept = significand >> drop;
            var rest = significand & ((1UL << drop) - 1);
            var half = 1UL << (drop - 1);
            up = rest > half || (rest == half && (kept & 1) != 0);
        }

        var magnitude = Math.ScaleB(up ? kept + 1 : kept, exponent + drop);
        return negative ? -magnitude : magnitude;
    }

    public static Extended operator -(Extended value) =>
        value.special ? new Extended(-value.specialValue) : new Extended(value.significand, value.exponent, !value.negative);

    public static Extended operator +(Extended left, Extended right) => Add(left, right, subtract: false);

    public static Extended operator -(Extended left, Extended right) => Add(left, right, subtract: true);

    public static Extended operator *(Extended left, Extended right)
    {
        if (left.special || right.special)
        {
            return new Extended(left.ToDouble() * right.ToDouble());
        }

        var negative = left.negative ^ right.negative;
        if (left.significand == 0 || right.significand == 0)
        {
            return new Extended(0, 0, negative);
        }

        return Round((UInt128)left.significand * right.significand, left.exponent + right.exponent, sticky: false, negative);
    }

    public static Extended operator /(Extended left, Extended right)
    {
        if (left.special || right.special || right.significand == 0)
        {
            return new Extended(left.ToDouble() / right.ToDouble());
        }

        var negative = left.negative ^ right.negative;
        if (left.significand == 0)
        {
            return new Extended(0, 0, negative);
        }

        var (quotient, remainder) = UInt128.DivRem((UInt128)left.significand << 64, right.significand);
        var exponent = left.exponent - right.exponent - 64;
        if (quotient >> 64 == 0)
        {
            // Below 2^64 the quotient needs one more bit before it can round.
            remainder <<= 1;
            quotient <<= 1;
            exponent--;
            if (remainder >= right.significand)
            {
                remainder -= right.significand;
                quotient |= 1;
            }
        }

        return Round(quotient, exponent, remainder != 0, negative);
    }

    private static Extended Add(Extended left, Extended right, bool subtract)
    {
        if (left.special || right.special)
        {
            return new Extended(subtract ? left.ToDouble() - right.ToDouble() : left.ToDouble() + right.ToDouble());
        }

        var rightNegative = right.negative ^ subtract;
        if (right.significand == 0)
        {
            return left.significand == 0 ? new Extended(0, 0, left.negative && rightNegative) : left;
        }

        if (left.significand == 0)
        {
            return new Extended(right.significand, right.exponent, rightNegative);
        }

        // The larger magnitude first; normalized significands order by exponent, then by significand.
        var (large, largeNegative, small, smallNegative) = right.exponent > left.exponent || (right.exponent == left.exponent && right.significand > left.significand)
            ? (right, rightNegative, left, left.negative)
            : (left, left.negative, right, rightNegative);
        var distance = large.exponent - small.exponent;
        var wideLarge = (UInt128)large.significand << 63;
        var wideSmall = (UInt128)small.significand << 63;
        var sticky = false;
        if (distance >= 127)
        {
            sticky = true;
            wideSmall = 0;
        }
        else if (distance > 0)
        {
            sticky = (wideSmall & ((UInt128.One << distance) - 1)) != 0;
            wideSmall >>= distance;
        }

        UInt128 sum;
        if (largeNegative == smallNegative)
        {
            sum = wideLarge + wideSmall;
        }
        else
        {
            // The bits shifted out made the smaller operand larger than what remains of it: the difference lies just below.
            sum = wideLarge - wideSmall - (sticky ? UInt128.One : UInt128.Zero);
            if (sum == 0)
            {
                return new Extended(0, 0, false);
            }
        }

        return Round(sum, large.exponent - 63, sticky, largeNegative);
    }

    // value = wide * 2^exponent, plus something below its last bit when sticky; rounded to 64 bits, nearest even.
    private static Extended Round(UInt128 wide, int exponent, bool sticky, bool negative)
    {
        var shift = (int)UInt128.LeadingZeroCount(wide);
        wide <<= shift;
        exponent -= shift;
        var significand = (ulong)(wide >> 64);
        var rest = (ulong)wide;
        if (rest > Half || (rest == Half && (sticky || (significand & 1) != 0)))
        {
            significand++;
            if (significand == 0)
            {
                significand = Half;
                exponent++;
            }
        }

        return new Extended(significand, exponent + 64, negative);
    }
}
