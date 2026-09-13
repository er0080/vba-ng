using System.Runtime.CompilerServices;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The Variant's own layout (ARCHITECTURE.md section 5, D20; ROADMAP.md M7 C2): 24 bytes, as the
/// VARIANT is, with a Decimal inline in the DECIMAL layout, so a Decimal goes in and comes out
/// without allocating (CLAUDE.md R20). What VBA does with Decimals is the goldens' business.
/// </summary>
public sealed class VariantTests
{
    /// <summary>The locals window shows a Date as VBA's writes it, between #s and rounded to the second (ARCHITECTURE.md section 9).</summary>
    [Fact]
    public void DebugView_ShowsADateAsVbaWritesIt()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;
        System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");
        try
        {
            Assert.Equal("#9/11/2026 2:30:00 PM# (Date)", Variant.FromDate(VbaDate.FromSerial(46276.604166666664)).ToString());
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = previous;
        }
    }

    public static TheoryData<decimal> Decimals() =>
    [
        0m,
        1.5m,
        -1.5m,
        1.500m,
        -0.0000000000000000000000000001m,
        79228162514264337593543950335m,
        -79228162514264337593543950335m,
        7.9228162514264337593543950335m,
        4294967296m,
        18446744073709551616m,
    ];

    [Fact]
    public void Variant_IsAVariantsSize() => Assert.Equal(24, Unsafe.SizeOf<Variant>());

    [Theory]
    [MemberData(nameof(Decimals))]
    public void Decimal_RoundTripsWithItsScale(decimal value)
    {
        var variant = Variant.FromDecimal(value);

        Assert.Equal(VarType.Decimal, variant.Type);
        Assert.Equal(value, variant.AsDecimal());
        Assert.Equal(value.Scale, variant.AsDecimal().Scale);
        Assert.Equal(decimal.IsNegative(value), decimal.IsNegative(variant.AsDecimal()));
    }

    [Fact]
    public void Decimal_EqualsTheSameValueAtAnotherScale()
    {
        Assert.Equal(Variant.FromDecimal(1.5m), Variant.FromDecimal(1.500m));
        Assert.Equal(Variant.FromDecimal(1.5m).GetHashCode(), Variant.FromDecimal(1.500m).GetHashCode());
        Assert.NotEqual(Variant.FromDecimal(1.5m), Variant.FromDecimal(-1.5m));
    }

    [Fact]
    public void Decimal_AllocatesNothing()
    {
        var total = 0m;
        total += Variant.FromDecimal(2.25m).AsDecimal();
        var before = GC.GetAllocatedBytesForCurrentThread();
        for (var i = 0; i < 1000; i++)
        {
            total += Variant.FromDecimal(i + 0.25m).AsDecimal();
        }

        Assert.Equal(before, GC.GetAllocatedBytesForCurrentThread());
        Assert.NotEqual(0m, total);
    }
}
