using System.Globalization;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// Typed operands through C# arithmetic (ROADMAP.md D-J, M7 E2): generated code calls the typed
/// helpers when both operands of +, -, or * have declared numeric types, and steps a typed For
/// counter natively, so each helper must mean what the Variant operators mean for the same typed
/// operands. The Variant operators are the ones the Operators, Errors, and ControlFlow goldens
/// back (R2); these tests hold the helpers to them over every pair of boundary values, the result
/// or the error number, and for a Double the error left pending for the statement's check.
/// </summary>
public sealed class TypedArithmeticTests
{
    private static readonly byte[] Bytes = [0, 1, 2, 15, 16, 17, 127, 128, 254, 255];

    private static readonly short[] Integers = [short.MinValue, short.MinValue + 1, -182, -181, -1, 0, 1, 181, 182, 255, 256, short.MaxValue - 1, short.MaxValue];

    private static readonly int[] Longs = [int.MinValue, int.MinValue + 1, -46341, -46340, -1, 0, 1, 46340, 46341, 65536, int.MaxValue - 1, int.MaxValue];

    private static readonly long[] LongLongs = [long.MinValue, long.MinValue + 1, -3037000500, -3037000499, -1, 0, 1, 3037000499, 3037000500, 4294967296, long.MaxValue - 1, long.MaxValue];

    private static readonly double[] Doubles = [0, -0.0, 1, -1, 0.5, 1e-308, 1e308, -1e308, double.MaxValue, -double.MaxValue, double.Epsilon, double.PositiveInfinity, double.NegativeInfinity];

    [Fact]
    public void Byte_AgreesWithTheVariantOperators() => Agree(
        Bytes,
        Variant.FromByte,
        v => v.AsByte(),
        ("+", Operators.AddByte, (l, r) => Operators.Add(l, r)),
        ("-", Operators.SubtractByte, (l, r) => Operators.Subtract(l, r)),
        ("*", Operators.MultiplyByte, (l, r) => Operators.Multiply(l, r)));

    [Fact]
    public void Integer_AgreesWithTheVariantOperators() => Agree(
        Integers,
        Variant.FromInt16,
        v => v.AsInt16(),
        ("+", Operators.AddInt16, (l, r) => Operators.Add(l, r)),
        ("-", Operators.SubtractInt16, (l, r) => Operators.Subtract(l, r)),
        ("*", Operators.MultiplyInt16, (l, r) => Operators.Multiply(l, r)));

    [Fact]
    public void Long_AgreesWithTheVariantOperators() => Agree(
        Longs,
        Variant.FromInt32,
        v => v.AsInt32(),
        ("+", Operators.AddInt32, (l, r) => Operators.Add(l, r)),
        ("-", Operators.SubtractInt32, (l, r) => Operators.Subtract(l, r)),
        ("*", Operators.MultiplyInt32, (l, r) => Operators.Multiply(l, r)));

    [Fact]
    public void LongLong_AgreesWithTheVariantOperators() => Agree(
        LongLongs,
        Variant.FromInt64,
        v => v.AsInt64(),
        ("+", Operators.AddInt64, (l, r) => Operators.Add(l, r)),
        ("-", Operators.SubtractInt64, (l, r) => Operators.Subtract(l, r)),
        ("*", Operators.MultiplyInt64, (l, r) => Operators.Multiply(l, r)));

    /// <summary>A typed Double operation stores its IEEE result and leaves 6 pending (docs/vba-quirks.md, "Runtime, floating point"), as the Variant operators do under the Floating flag.</summary>
    [Fact]
    public void Double_AgreesWithTheVariantOperators_TheOverflowPending() => Agree(
        Doubles,
        Variant.FromDouble,
        v => v.AsDouble(),
        ("+", Operators.AddDouble, (l, r) => Operators.Add(l, r, DeclaredTypes.Floating)),
        ("-", Operators.SubtractDouble, (l, r) => Operators.Subtract(l, r, DeclaredTypes.Floating)),
        ("*", Operators.MultiplyDouble, (l, r) => Operators.Multiply(l, r, DeclaredTypes.Floating)),
        ("/", Operators.DivideDouble, (l, r) => Operators.Divide(l, r, DeclaredTypes.Floating)));

    /// <summary>The Next of a typed For counter is the Variant loop's Operators.Add with neither operand a Variant: an overflow raises at once, a Double's included.</summary>
    [Fact]
    public void ForNext_AgreesWithTheVariantLoopsIncrement()
    {
        Agree(Integers, Variant.FromInt16, v => v.AsInt16(), ("Next", ForLoop.Next, (l, r) => Operators.Add(l, r)));
        Agree(Longs, Variant.FromInt32, v => v.AsInt32(), ("Next", ForLoop.Next, (l, r) => Operators.Add(l, r)));
        Agree(Doubles, Variant.FromDouble, v => v.AsDouble(), ("Next", ForLoop.Next, (l, r) => Operators.Add(l, r)));
    }

    /// <summary>The typed loop test is the Variant loop's: a step of zero or more counts up to the limit, a negative step down to it.</summary>
    [Fact]
    public void ForContinues_AgreesWithTheVariantLoopsTest()
    {
        int[] values = [int.MinValue, -2, -1, 0, 1, 2, int.MaxValue];
        foreach (var counter in values)
        {
            foreach (var limit in values)
            {
                foreach (var step in values)
                {
                    var boxed = ForLoop.Continues(Variant.FromInt32(counter), Variant.FromInt32(limit), Variant.FromInt32(step));
                    Assert.True(boxed == ForLoop.Continues(counter, limit, step), $"Continues({counter}, {limit}, {step})");
                    Assert.True(boxed == ForLoop.Continues((double)counter, limit, step), $"Continues({counter}.0, {limit}, {step})");
                }
            }
        }
    }

    private static void Agree<T>(T[] values, Func<T, Variant> box, Func<Variant, T> unbox, params (string Name, Func<T, T, T> Typed, Func<Variant, Variant, Variant> Boxed)[] operations)
    {
        Operators.CheckFloating();
        var disagreements = new List<string>();
        foreach (var (name, typed, boxed) in operations)
        {
            foreach (var left in values)
            {
                foreach (var right in values)
                {
                    var expected = Outcome(() => Format(unbox(boxed(box(left), box(right)))));
                    var actual = Outcome(() => Format(typed(left, right)));
                    if (expected != actual)
                    {
                        disagreements.Add($"{Format(left)} {name} {Format(right)}: typed {actual}, Variant {expected}");
                    }
                }
            }
        }

        Assert.Empty(disagreements);
    }

    /// <summary>The value, or the error number raised, and any floating-point error left pending after the value.</summary>
    private static string Outcome(Func<string> compute)
    {
        string value;
        try
        {
            value = compute();
        }
        catch (VbaException e)
        {
            return "error " + e.Number.ToString(CultureInfo.InvariantCulture);
        }

        try
        {
            Operators.CheckFloating();
            return value;
        }
        catch (VbaException e)
        {
            return value + " then error " + e.Number.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static string Format<T>(T value) => value is double d ? BitConverter.DoubleToInt64Bits(d).ToString("X16", CultureInfo.InvariantCulture) : Convert.ToString(value, CultureInfo.InvariantCulture)!;
}
