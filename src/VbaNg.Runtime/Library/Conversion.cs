using System.Globalization;

namespace VbaNg.Runtime.Library;

/// <summary>The Conversion module of the VBA standard library (MS-VBAL 6.1.2.3).</summary>
public static class Conversion
{
    public static bool CBool(in Variant value) => Coerce.ToBoolean(value);

    /// <summary>An Error value converts by its number, unlike let-coercion (Conversion goldens).</summary>
    public static byte CByte(in Variant value) => Coerce.ToByte(ErrorAsNumber(value));

    public static Currency CCur(in Variant value) => Coerce.ToCurrency(ErrorAsNumber(value));

    public static VbaDate CDate(in Variant value) => Coerce.ToDate(value);

    public static double CDbl(in Variant value) => Coerce.ToDouble(ErrorAsNumber(value));

    /// <summary>MS-VBAL 6.1.2.3.1.6 CDec.</summary>
    public static decimal CDec(in Variant value) => Coerce.ToDecimal(ErrorAsNumber(value));

    public static short CInt(in Variant value) => Coerce.ToInt16(ErrorAsNumber(value));

    public static int CLng(in Variant value) => Coerce.ToInt32(ErrorAsNumber(value));

    public static long CLngLng(in Variant value) => Coerce.ToInt64(ErrorAsNumber(value));

    public static long CLngPtr(in Variant value) => Coerce.ToInt64(ErrorAsNumber(value));

    public static float CSng(in Variant value) => Coerce.ToSingle(ErrorAsNumber(value));

    /// <summary>CStr: a String's own BSTR without a copy, any other value's text as a temporary of the statement (D20); Null raises 94.</summary>
    public static VbaString CStr(in Variant value) => value.IsNull ? throw VbaErrors.InvalidUseOfNull() : Coerce.ToText(value);

    public static Variant CVar(in Variant value) => value;

    /// <summary>MS-VBAL 6.1.2.3.1.14 CVErr: numbers 0 to 65535, else error 5.</summary>
    public static ErrorValue CVErr(in Variant value) => ErrorValue.FromNumber(Coerce.ToInt64(value));

    /// <summary>Error$(number): the message for a number, or "" for 0; Null raises 94 (Information golden).</summary>
    public static Variant Error(in Variant number) => Variant.FromString(VbaErrors.ErrorText(Coerce.ToInt64(number)));

    /// <summary>Fix: truncates toward zero and keeps the operand's numeric type; strings give Double, Null gives Null.</summary>
    public static Variant Fix(in Variant value) => Truncate(value, toward: 0);

    /// <summary>Int: rounds toward negative infinity and keeps the operand's numeric type.</summary>
    public static Variant Int(in Variant value) => Truncate(value, toward: -1);

    /// <summary>Hex$: the two's complement digits of the value's own width; Doubles beyond Long use 64 bits.</summary>
    public static Variant Hex(in Variant value) => Radix(value, 16);

    public static Variant Oct(in Variant value) => Radix(value, 8);

    /// <summary>
    /// Str: number text with "." as the decimal separator and a leading space for non-negative
    /// numbers; Null gives Null, an Error value gives "Error n", other values go through String coercion.
    /// </summary>
    public static Variant Str(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Error:
                return Variant.FromString(value.AsError().ToString());
            case VarType.Date:
            case VarType.Boolean:
                return Variant.FromString(Coerce.ToString(value));
            case VarType.Integer:
            case VarType.Long:
            case VarType.LongLong:
            case VarType.Byte:
            case VarType.Single:
            case VarType.Double:
            case VarType.Currency:
            case VarType.Decimal:
                return Variant.FromString(NumericText(value));
            default:
                return Variant.FromString(NumericText(Variant.FromDouble(Coerce.ToDouble(value))));
        }

        static string NumericText(in Variant number)
        {
            var text = number.Type switch
            {
                VarType.Single => NumberText.FormatSingle(number.AsSingle()),
                VarType.Double => NumberText.FormatDouble(number.AsDouble()),
                VarType.Currency => number.AsCurrency().ToString(),
                VarType.Decimal => DecimalText.Normalize(number.AsDecimal()).ToString(CultureInfo.InvariantCulture),
                _ => Coerce.ToInt64(number).ToString(CultureInfo.InvariantCulture),
            };

            // Str omits the zero before the decimal point: 0.5 prints as " .5" (Conversion golden).
            if (text.StartsWith("0.", StringComparison.Ordinal))
            {
                text = text[1..];
            }
            else if (text.StartsWith("-0.", StringComparison.Ordinal))
            {
                text = "-" + text[2..];
            }

            return text.StartsWith('-') ? text : " " + text;
        }
    }

    /// <summary>
    /// Val: reads the longest numeric prefix, ignoring blanks, tabs, and line breaks anywhere,
    /// with &amp;H and &amp;O prefixes and E or D exponents; anything unreadable gives 0.
    /// </summary>
    public static double Val(in Variant value)
    {
        if (value.IsNull)
        {
            throw VbaErrors.InvalidUseOfNull();
        }

        var text = Coerce.ToString(value);
        var compact = string.Concat(text.Where(static c => c is not (' ' or '\t' or '\n' or '\r' or '\f' or '\v')));
        if (compact.Length >= 2 && compact[0] == '&' && compact[1] is 'H' or 'h' or 'O' or 'o')
        {
            var radix = compact[1] is 'H' or 'h' ? 16 : 8;
            var end = 2;
            long magnitude = 0;
            while (end < compact.Length)
            {
                var digit = compact[end] switch
                {
                    >= '0' and <= '9' => compact[end] - '0',
                    >= 'A' and <= 'F' => compact[end] - 'A' + 10,
                    >= 'a' and <= 'f' => compact[end] - 'a' + 10,
                    _ => -1,
                };
                if (digit < 0 || digit >= radix)
                {
                    break;
                }

                magnitude = unchecked(magnitude * radix + digit);
                end++;
            }

            var suffix = end < compact.Length ? compact[end] : '\0';
            return suffix switch
            {
                '&' => (int)(uint)magnitude,
                '^' => magnitude,
                _ when (ulong)magnitude <= 0xFFFF => (short)(ushort)magnitude,
                _ when (ulong)magnitude <= 0xFFFFFFFF => (int)(uint)magnitude,
                _ => magnitude,
            };
        }

        var i = 0;
        var negative = false;
        if (i < compact.Length && compact[i] is '+' or '-')
        {
            negative = compact[i] == '-';
            i++;
        }

        var start = i;
        while (i < compact.Length && char.IsAsciiDigit(compact[i]))
        {
            i++;
        }

        if (i < compact.Length && compact[i] == '.')
        {
            i++;
            while (i < compact.Length && char.IsAsciiDigit(compact[i]))
            {
                i++;
            }
        }

        var mantissa = compact[start..i];
        if (mantissa.Length == 0 || mantissa == ".")
        {
            return 0;
        }

        var exponent = 0;
        if (i < compact.Length && compact[i] is 'e' or 'E' or 'd' or 'D')
        {
            var j = i + 1;
            var exponentNegative = false;
            if (j < compact.Length && compact[j] is '+' or '-')
            {
                exponentNegative = compact[j] == '-';
                j++;
            }

            var digitsStart = j;
            while (j < compact.Length && char.IsAsciiDigit(compact[j]))
            {
                j++;
            }

            if (j > digitsStart && j - digitsStart <= 9)
            {
                exponent = int.Parse(compact.AsSpan(digitsStart, j - digitsStart), NumberStyles.None, CultureInfo.InvariantCulture);
                if (exponentNegative)
                {
                    exponent = -exponent;
                }
            }
        }

        var result = double.Parse(mantissa + "E" + exponent.ToString(CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
        return negative ? -result : result;
    }

    private static Variant ErrorAsNumber(in Variant value) => value.IsError ? Variant.FromInt32(value.AsError().Number) : value;

    private static Variant Truncate(in Variant value, int toward)
    {
        switch (value.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Empty:
                return Variant.FromInt16(0);
            case VarType.Integer:
            case VarType.Long:
            case VarType.LongLong:
            case VarType.Byte:
                return value;
            case VarType.Boolean:
                return Variant.FromInt16(Coerce.ToInt16(value));
            case VarType.Single:
                {
                    var d = value.AsSingle();
                    return Variant.FromSingle(toward < 0 ? MathF.Floor(d) : MathF.Truncate(d));
                }

            case VarType.Double:
                return Variant.FromDouble(toward < 0 ? Math.Floor(value.AsDouble()) : Math.Truncate(value.AsDouble()));
            case VarType.Date:
                {
                    var serial = value.AsDate().Serial;
                    return Variant.FromDate(VbaDate.FromSerial(toward < 0 ? Math.Floor(serial) : Math.Truncate(serial)));
                }

            case VarType.Currency:
                {
                    var d = value.AsCurrency().ToDecimal();
                    return Variant.FromCurrency(Currency.FromDecimal(toward < 0 ? decimal.Floor(d) : decimal.Truncate(d)));
                }

            case VarType.Decimal:
                {
                    var d = value.AsDecimal();
                    return Variant.FromDecimal(toward < 0 ? decimal.Floor(d) : decimal.Truncate(d));
                }

            default:
                {
                    var d = Coerce.ToDouble(value);
                    return Variant.FromDouble(toward < 0 ? Math.Floor(d) : Math.Truncate(d));
                }
        }
    }

    private static Variant Radix(in Variant value, int radix)
    {
        if (value.IsNull)
        {
            return Variant.Null;
        }

        long bits;
        int width;
        switch (value.Type)
        {
            case VarType.Empty:
                bits = 0;
                width = 16;
                break;
            case VarType.Integer:
                bits = value.AsInt16();
                width = 16;
                break;
            case VarType.Byte:
                bits = value.AsByte();
                width = 8;
                break;
            case VarType.Boolean:
                bits = value.AsBoolean() ? -1 : 0;
                width = 16;
                break;
            case VarType.Long:
                bits = value.AsInt32();
                width = 32;
                break;
            case VarType.LongLong:
                bits = value.AsInt64();
                width = 64;
                break;
            case VarType.String:
                {
                    // A string reads as a number and takes the narrowest integer width that holds it.
                    bits = Coerce.ToInt64(value);
                    width = bits is >= short.MinValue and <= short.MaxValue ? 16 : bits is >= int.MinValue and <= int.MaxValue ? 32 : 64;
                    break;
                }

            default:
                // Floating-point and fixed-point values round and print with 64 bits: Hex(-1#) is sixteen Fs (Conversion golden).
                bits = Coerce.ToInt64(value);
                width = 64;
                break;
        }

        var unsigned = width == 64 ? (ulong)bits : (ulong)bits & ((1UL << width) - 1);
        var text = radix == 16 ? unsigned.ToString("X", CultureInfo.InvariantCulture) : Convert.ToString((long)unsigned, 8);
        if (radix == 8 && bits < 0 && width < 64)
        {
            text = Convert.ToString((long)unsigned, 8);
        }

        return Variant.FromString(text);
    }
}
