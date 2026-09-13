namespace VbaNg.Runtime.Library;

/// <summary>
/// The Math module of the VBA standard library (MS-VBAL 6.1.2, Math module). Abs, Sgn, and
/// Round keep the operand's numeric type; the transcendental functions return Double and pass
/// Null through.
/// </summary>
public static class VbaMath
{
    /// <summary>Abs keeps the type of its operand; the Integer and Long minimums overflow; strings become Double.</summary>
    public static Variant Abs(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Empty:
                return Variant.FromInt16(0);
            case VarType.Byte:
                return value;
            case VarType.Boolean:
            case VarType.Integer:
                {
                    var n = Coerce.ToInt16(value);
                    return n == short.MinValue ? throw VbaErrors.Overflow() : Variant.FromInt16(Math.Abs(n));
                }

            case VarType.Long:
                {
                    var n = value.AsInt32();
                    return n == int.MinValue ? throw VbaErrors.Overflow() : Variant.FromInt32(Math.Abs(n));
                }

            case VarType.LongLong:
                {
                    var n = value.AsInt64();
                    return n == long.MinValue ? throw VbaErrors.Overflow() : Variant.FromInt64(Math.Abs(n));
                }

            case VarType.Single:
                return Variant.FromSingle(Math.Abs(value.AsSingle()));
            case VarType.Double:
                return Variant.FromDouble(Math.Abs(value.AsDouble()));
            case VarType.Currency:
                {
                    var c = value.AsCurrency();
                    return Variant.FromCurrency(c.Scaled < 0 ? Currency.Negate(c) : c);
                }

            case VarType.Decimal:
                return Variant.FromDecimal(Math.Abs(value.AsDecimal()));
            case VarType.Date:
                return Variant.FromDate(VbaDate.FromSerial(Math.Abs(value.AsDate().Serial)));
            default:
                return Variant.FromDouble(Math.Abs(Coerce.ToDouble(value)));
        }
    }

    /// <summary>Sgn: -1, 0, or 1 as an Integer; Null raises 94 (Math golden).</summary>
    public static Variant Sgn(in Variant value)
    {
        if (value.IsNull)
        {
            throw VbaErrors.InvalidUseOfNull();
        }

        return value.Type switch
        {
            VarType.Currency => Variant.FromInt16((short)Math.Sign(value.AsCurrency().Scaled)),
            VarType.Decimal => Variant.FromInt16((short)Math.Sign(value.AsDecimal())),
            _ => Variant.FromInt16((short)Math.Sign(Coerce.ToDouble(value))),
        };
    }

    /// <summary>Sqr: a negative operand raises 5; Null raises 94, since the parameter is a Double (Math golden).</summary>
    public static double Sqr(in Variant value)
    {
        var x = Coerce.ToDouble(value);
        return x < 0 ? throw VbaErrors.InvalidProcedureCall() : Math.Sqrt(x);
    }

    public static double Exp(in Variant value)
    {
        var result = Math.Exp(Coerce.ToDouble(value));
        return double.IsInfinity(result) ? throw VbaErrors.Overflow() : result;
    }

    /// <summary>Log: zero or a negative operand raises 5.</summary>
    public static double Log(in Variant value)
    {
        var x = Coerce.ToDouble(value);
        return x <= 0 ? throw VbaErrors.InvalidProcedureCall() : Math.Log(x);
    }

    public static double Sin(in Variant value) => Math.Sin(Coerce.ToDouble(value));

    public static double Cos(in Variant value) => Math.Cos(Coerce.ToDouble(value));

    /// <summary>Tan as the quotient of the sine and cosine, each rounded to Double: Excel's Tan(1) is one ulp above the correctly rounded tangent and equal to Sin(1) / Cos(1) (Math golden).</summary>
    public static double Tan(in Variant value)
    {
        var angle = Coerce.ToDouble(value);
        return Math.Sin(angle) / Math.Cos(angle);
    }

    public static double Atn(in Variant value) => Math.Atan(Coerce.ToDouble(value));

    /// <summary>Round(Number, [NumDigitsAfterDecimal]): half to even, keeping the operand's numeric type; Null gives Null; digits below 0 or above 22 raise 5.</summary>
    public static Variant Round(in Variant value, in Variant digits)
    {
        var places = digits.IsMissing ? 0 : Coerce.ToInt32(digits);
        if (places is < 0 or > 22)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        switch (value.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Empty:
                return Variant.FromInt16(0);
            case VarType.Byte:
            case VarType.Integer:
            case VarType.Long:
            case VarType.LongLong:
                return value;
            case VarType.Boolean:
                return Variant.FromInt16(Coerce.ToInt16(value));
            case VarType.Currency:
                return Variant.FromCurrency(Currency.FromDecimal(decimal.Round(value.AsCurrency().ToDecimal(), Math.Min(places, 28), MidpointRounding.ToEven)));
            case VarType.Decimal:
                return Variant.FromDecimal(decimal.Round(value.AsDecimal(), Math.Min(places, 28), MidpointRounding.ToEven));
            case VarType.Single:
                return Variant.FromSingle((float)RoundDouble(value.AsSingle(), places));
            case VarType.Date:
                return Variant.FromDate(VbaDate.FromSerial(RoundDouble(value.AsDate().Serial, places)));
            default:
                return Variant.FromDouble(RoundDouble(Coerce.ToDouble(value), places));
        }
    }

    private static double RoundDouble(double value, int places)
    {
        if (places > 15)
        {
            return value;
        }

        return Math.Round(value, places, MidpointRounding.ToEven);
    }
}

/// <summary>
/// VBA's Rnd generator: a 24-bit linear congruential generator with the multiplier 0x43FD43FD
/// and increment 0xC39EC3, seeded to 0x50000 at start. Rnd(negative) reseeds from the
/// argument's Single bits, Rnd(0) repeats the last value, and Randomize mixes a seed into the
/// high bits. The Math goldens pin every value.
/// </summary>
public sealed class RandomGenerator
{
    private const uint Multiplier = 0x43FD43FD;
    private const uint Increment = 0xC39EC3;
    private const uint Modulus = 0x1000000;

    private uint state = 0x50000;

    /// <summary>The generator behind Rnd and Randomize in compiled code; one per process, as VBA keeps one per session. Replaceable so tests start from a known state.</summary>
    public static RandomGenerator Shared { get; set; } = new();

    public float Next() => Advance();

    /// <summary>Rnd(Number): negative reseeds and returns the first value of the new sequence, 0 repeats the last value.</summary>
    public float Next(in Variant number)
    {
        var n = Coerce.ToSingle(number);
        if (n < 0)
        {
            var bits = (uint)BitConverter.SingleToInt32Bits(n);
            state = (bits + (bits >> 24)) & (Modulus - 1);
            return Advance();
        }

        if (n == 0)
        {
            return state / (float)Modulus;
        }

        return Advance();
    }

    /// <summary>Randomize [Number]: with a number, its Double bits seed the generator; without one, the timer does.</summary>
    public void Randomize(in Variant number)
    {
        double seed = number.IsMissing ? Environment.TickCount64 / 1000.0 : Coerce.ToDouble(number);
        var bits = (ulong)BitConverter.DoubleToInt64Bits(seed);
        var mixed = (uint)(bits >> 32) ^ (uint)bits;
        var high = (mixed & 0xFFFF) ^ (mixed >> 16);
        state = (state & 0xFF) | ((high & 0xFFFF) << 8);
    }

    private float Advance()
    {
        state = (state * Multiplier + Increment) & (Modulus - 1);
        return state / (float)Modulus;
    }
}
