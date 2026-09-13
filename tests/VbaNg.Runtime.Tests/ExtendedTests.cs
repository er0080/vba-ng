using System.Numerics;

using VbaNg.Runtime.Library;

using Xunit;

using Assert = Xunit.Assert;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The x87 extended format the Financial functions compute in (ROADMAP.md D-F): every operation
/// rounds its exact result to a 64-bit significand, nearest even, and a store to Double rounds
/// once more. The oracle computes each result exactly, as a fraction of BigIntegers, and rounds
/// it once; the Financial golden backs the format itself (R2).
/// </summary>
public sealed class ExtendedTests
{
    [Fact]
    public void Arithmetic_RoundsTheExactResultToSixtyFourBits()
    {
        var random = new Random(12345);
        double Next() => Math.ScaleB((random.NextDouble() * 2) - 1, random.Next(-70, 70));
        var misses = new List<string>();
        void Check(string operation, Extended actual, Fraction expected)
        {
            if (Pin(actual) != Pin(expected) && misses.Count < 5)
            {
                misses.Add($"{operation}: {Pin(actual)} instead of {Pin(expected)}");
            }
        }

        for (var i = 0; i < 20000; i++)
        {
            double a = Next(), b = Next(), c = Next(), d = Next();
            if (a == 0 || b == 0 || c == 0 || d == 0)
            {
                continue;
            }

            // Operands with a full 64-bit significand, built alike on both sides; every fourth pair nearly cancels.
            Extended left = (Extended)a / c, right = (Extended)b * d;
            Fraction exactLeft = (Fraction.Of(a) / Fraction.Of(c)).Round(64), exactRight = (Fraction.Of(b) * Fraction.Of(d)).Round(64);
            if (i % 4 == 1)
            {
                right = -left + b;
                exactRight = (-exactLeft + Fraction.Of(b)).Round(64);
            }

            Check("+", left + right, (exactLeft + exactRight).Round(64));
            Check("-", left - right, (exactLeft - exactRight).Round(64));
            Check("*", left * right, (exactLeft * exactRight).Round(64));
            if (!exactRight.Numerator.IsZero)
            {
                Check("/", left / right, (exactLeft / exactRight).Round(64));
            }
        }

        Assert.Empty(misses);
    }

    [Fact]
    public void ToDouble_RoundsASecondTime()
    {
        // 1 + 2^-53 + 2^-70 rounds to 1 + 2^-53 in 64 bits, a tie that goes to 1 in 53; rounded once it is 1 + 2^-52.
        var value = (Extended)1.0 + Math.ScaleB(1, -53) + Math.ScaleB(1, -70);

        Assert.Equal(1.0, value.ToDouble());
    }

    [Fact]
    public void Intermediate_PassesTheRangeOfADouble() =>
        Assert.Equal(1e308, ((Extended)1e308 * 10 / 10).ToDouble());

    [Fact]
    public void Division_ByZero_ContinuesAsADouble()
    {
        Assert.Equal(double.PositiveInfinity, ((Extended)1 / 0.0).ToDouble());
        Assert.True(double.IsNaN(((Extended)0.0 / 0.0).ToDouble()));
    }

    // A 64-bit value as its Double and the exact residual below it.
    private static (double, double) Pin(Extended value)
    {
        var rounded = value.ToDouble();
        return (rounded, (value - rounded).ToDouble());
    }

    private static (double, double) Pin(Fraction value)
    {
        var rounded = value.ToDouble();
        return (rounded, (value - Fraction.Of(rounded)).ToDouble());
    }

    private readonly record struct Fraction(BigInteger Numerator, BigInteger Denominator)
    {
        public static Fraction Of(double value)
        {
            var bits = BitConverter.DoubleToInt64Bits(value);
            var exponent = (int)((bits >> 52) & 0x7FF);
            var significand = bits & 0xF_FFFF_FFFF_FFFF;
            if (exponent == 0)
            {
                exponent = 1;
            }
            else
            {
                significand |= 1L << 52;
            }

            exponent -= 1075;
            BigInteger numerator = bits < 0 ? -significand : significand;
            return exponent >= 0 ? new(numerator << exponent, 1) : new(numerator, BigInteger.One << -exponent);
        }

        public static Fraction operator +(Fraction left, Fraction right) =>
            new((left.Numerator * right.Denominator) + (right.Numerator * left.Denominator), left.Denominator * right.Denominator);

        public static Fraction operator -(Fraction left, Fraction right) => left + -right;

        public static Fraction operator -(Fraction value) => new(-value.Numerator, value.Denominator);

        public static Fraction operator *(Fraction left, Fraction right) => new(left.Numerator * right.Numerator, left.Denominator * right.Denominator);

        public static Fraction operator /(Fraction left, Fraction right) =>
            right.Numerator.Sign < 0
                ? new(-left.Numerator * right.Denominator, left.Denominator * -right.Numerator)
                : new(left.Numerator * right.Denominator, left.Denominator * right.Numerator);

        // Rounded to the given number of significand bits, nearest even.
        public Fraction Round(int bits)
        {
            if (Numerator.IsZero)
            {
                return this;
            }

            var magnitude = BigInteger.Abs(Numerator);
            var shift = bits + 2 - (int)(magnitude.GetBitLength() - Denominator.GetBitLength());
            var (quotient, remainder) = shift >= 0
                ? BigInteger.DivRem(magnitude << shift, Denominator)
                : BigInteger.DivRem(magnitude, Denominator << -shift);
            var extra = (int)quotient.GetBitLength() - bits;
            var kept = quotient >> extra;
            var dropped = quotient - (kept << extra);
            var half = BigInteger.One << (extra - 1);
            if (dropped > half || (dropped == half && (!remainder.IsZero || !kept.IsEven)))
            {
                kept += 1;
            }

            if (Numerator.Sign < 0)
            {
                kept = -kept;
            }

            var exponent = extra - shift;
            return exponent >= 0 ? new(kept << exponent, 1) : new(kept, BigInteger.One << -exponent);
        }

        public double ToDouble()
        {
            var rounded = Round(53);
            return (double)rounded.Numerator / (double)rounded.Denominator;
        }
    }
}
