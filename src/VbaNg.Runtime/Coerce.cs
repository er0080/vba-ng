using System.Globalization;

namespace VbaNg.Runtime;

/// <summary>
/// Let-coercion (MS-VBAL 5.5.1): what happens when a value meets a destination of another
/// declared type, as in <c>Dim i As Integer: i = v</c>, argument passing, and operator operands.
/// Null into a non-Variant raises error 94, an Error value raises 13, a string that is not a
/// number raises 13, and a value outside the destination's range raises 6.
/// </summary>
public static class Coerce
{
    /// <summary>The regional settings used for string conversions: the current culture.</summary>
    public static CultureInfo Culture => CultureInfo.CurrentCulture;

    /// <summary>MS-VBAL 5.5.1.2.1 Let-coercion between numeric types, destination Integer.</summary>
    public static short ToInt16(in Variant value) => (short)ToIntegral(value, short.MinValue, short.MaxValue);

    public static int ToInt32(in Variant value) => (int)ToIntegral(value, int.MinValue, int.MaxValue);

    public static long ToInt64(in Variant value) => ToIntegral(value, long.MinValue, long.MaxValue);

    /// <summary>Byte from Boolean True is 255 (MS-VBAL 5.5.1.2.2 Let-coercion to and from Boolean).</summary>
    public static byte ToByte(in Variant value) =>
        value.Type == VarType.Boolean ? (value.AsBoolean() ? (byte)255 : (byte)0) : (byte)ToIntegral(value, byte.MinValue, byte.MaxValue);

    public static float ToSingle(in Variant value)
    {
        if (value.IsObject)
        {
            return ToSingle(ObjectValue(value));
        }

        if (value.Type == VarType.Single)
        {
            return value.AsSingle();
        }

        var wide = ToDouble(value);
        if (double.IsNaN(wide) || Math.Abs(wide) > float.MaxValue)
        {
            throw VbaErrors.Overflow();
        }

        return (float)wide;
    }

    /// <summary>MS-VBAL 5.5.1.2.1, destination Double; strings parse under the regional settings (5.5.1.2.4).</summary>
    public static double ToDouble(in Variant value)
    {
        if (value.IsObject)
        {
            return ToDouble(ObjectValue(value));
        }

        switch (value.Type)
        {
            case VarType.Double:
                return value.AsDouble();
            case VarType.Single:
                return value.AsSingle();
            case VarType.Integer:
                return value.AsInt16();
            case VarType.Long:
                return value.AsInt32();
            case VarType.LongLong:
                return value.AsInt64();
            case VarType.Byte:
                return value.AsByte();
            case VarType.Boolean:
                return value.BooleanBits;
            case VarType.Currency:
                return value.AsCurrency().ToDouble();
            case VarType.Decimal:
                return (double)value.AsDecimal();
            case VarType.Date:
                return value.AsDate().Serial;
            case VarType.Empty:
                return 0;
            case VarType.String:
                if (!NumberText.TryParse(value.AsVbaString().Chars, Culture, out _, out var parsed, out _))
                {
                    throw VbaErrors.TypeMismatch();
                }

                if (double.IsInfinity(parsed))
                {
                    throw VbaErrors.Overflow();
                }

                return parsed;
            default:
                throw NotConvertible(value);
        }
    }

    /// <summary>Currency from a Double scales in floating point and rounds half to even (MS-VBAL 5.5.1.2.1).</summary>
    public static Currency ToCurrency(in Variant value)
    {
        if (value.IsObject)
        {
            return ToCurrency(ObjectValue(value));
        }

        switch (value.Type)
        {
            case VarType.Currency:
                return value.AsCurrency();
            case VarType.Integer:
                return Currency.FromInt64(value.AsInt16());
            case VarType.Long:
                return Currency.FromInt64(value.AsInt32());
            case VarType.LongLong:
                return Currency.FromInt64(value.AsInt64());
            case VarType.Byte:
                return Currency.FromInt64(value.AsByte());
            case VarType.Boolean:
                return Currency.FromInt64(value.BooleanBits);
            case VarType.Decimal:
                return Currency.FromDecimal(value.AsDecimal());
            case VarType.Empty:
                return Currency.Zero;
            case VarType.String:
                if (!NumberText.TryParse(value.AsVbaString().Chars, Culture, out var exact, out var wide, out var isExact))
                {
                    throw VbaErrors.TypeMismatch();
                }

                return isExact ? Currency.FromDecimal(exact) : Currency.FromDouble(wide);
            case VarType.Single or VarType.Double or VarType.Date:
                return Currency.FromDouble(ToDouble(value));
            default:
                throw NotConvertible(value);
        }
    }

    /// <summary>Decimal from a Double keeps 15 significant digits, as OLE's VarDecFromR8 does (Conversion goldens).</summary>
    public static decimal ToDecimal(in Variant value)
    {
        if (value.IsObject)
        {
            return ToDecimal(ObjectValue(value));
        }

        try
        {
            switch (value.Type)
            {
                case VarType.Decimal:
                    return value.AsDecimal();
                case VarType.Currency:
                    return value.AsCurrency().ToDecimal();
                case VarType.Integer:
                    return value.AsInt16();
                case VarType.Long:
                    return value.AsInt32();
                case VarType.LongLong:
                    return value.AsInt64();
                case VarType.Byte:
                    return value.AsByte();
                case VarType.Boolean:
                    return value.BooleanBits;
                case VarType.Empty:
                    return 0;
                case VarType.String:
                    if (!NumberText.TryParse(value.AsVbaString().Chars, Culture, out var exact, out _, out var isExact))
                    {
                        throw VbaErrors.TypeMismatch();
                    }

                    if (!isExact)
                    {
                        throw VbaErrors.Overflow();
                    }

                    // Trailing zeros in the text do not survive, CDec("1.500") has scale 1, but a zero keeps the text's scale and sign: CDec("-0.0") is a negative zero with one place (Conversion golden).
                    return exact == 0 ? exact : DecimalText.Normalize(exact);
                case VarType.Single or VarType.Double or VarType.Date:
                    return new decimal(ToDouble(value));
                default:
                    throw NotConvertible(value);
            }
        }
        catch (OverflowException)
        {
            throw VbaErrors.Overflow();
        }
    }

    /// <summary>MS-VBAL 5.5.1.2.3 Let-coercion to and from Date: numbers are serials, strings parse as dates first.</summary>
    public static VbaDate ToDate(in Variant value)
    {
        if (value.IsObject)
        {
            return ToDate(ObjectValue(value));
        }

        switch (value.Type)
        {
            case VarType.Date:
                return value.AsDate();
            case VarType.Empty:
                return VbaDate.Zero;
            case VarType.String:
                if (!DateText.TryParse(value.AsVbaString().Chars, Culture, out var date))
                {
                    throw VbaErrors.TypeMismatch();
                }

                return date;
            case VarType.Null or VarType.Error or VarType.Object or VarType.Array:
                throw NotConvertible(value);
            default:
                return VbaDate.FromSerial(ToDouble(value));
        }
    }

    /// <summary>MS-VBAL 5.5.1.2.2 Let-coercion to and from Boolean; strings "True"/"False" and "#TRUE#"/"#FALSE#" (5.5.1.2.4).</summary>
    public static bool ToBoolean(in Variant value)
    {
        if (value.IsObject)
        {
            return ToBoolean(ObjectValue(value));
        }

        switch (value.Type)
        {
            case VarType.Boolean:
                return value.AsBoolean();
            case VarType.Empty:
                return false;
            case VarType.String:
                var text = value.AsVbaString().Chars;
                if (text.Equals("True", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (text.Equals("False", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (text.SequenceEqual("#TRUE#"))
                {
                    return true;
                }

                if (text.SequenceEqual("#FALSE#"))
                {
                    return false;
                }

                return ToDouble(value) != 0;
            case VarType.Decimal:
                return value.AsDecimal() != 0;
            case VarType.Currency:
                return value.AsCurrency().Scaled != 0;
            case VarType.Null or VarType.Error or VarType.Object or VarType.Array:
                throw NotConvertible(value);
            default:
                return ToDouble(value) != 0;
        }
    }

    /// <summary>MS-VBAL 5.5.1.2.4 Let-coercion to String: 15 or 7 significant digits, regional decimal separator, Short Date and Long Time.</summary>
    /// <summary>
    /// The text of a value as a view over a BSTR (D20): a String's own BSTR with no copy, any
    /// other value's string form (MS-VBAL 5.6.9.3) as a temporary of the current statement.
    /// </summary>
    public static VbaString ToText(in Variant value) => value.Type switch
    {
        VarType.String => value.AsVbaString(),
        VarType.Integer => IntegerText(value.AsInt16()),
        VarType.Long => IntegerText(value.AsInt32()),
        VarType.LongLong => IntegerText(value.AsInt64()),
        VarType.Byte => IntegerText(value.AsByte()),

        // A Byte array is its bytes as a String, every byte of them; any other array is a type mismatch (Strings golden).
        _ when value.IsArray => value.AsArray().ToText(),
        _ => Variant.FromString(ToString(value)).AsVbaString(),
    };

    /// <summary>An integer's digits straight into a BSTR the statement owns, with no .NET string on the way (R20).</summary>
    private static VbaString IntegerText(long number)
    {
        Span<char> digits = stackalloc char[20];
        number.TryFormat(digits, out var written, provider: CultureInfo.InvariantCulture);
        return Variant.FromVbaString(VbaString.Alloc(digits[..written])).AsVbaString();
    }

    public static string ToString(in Variant value)
    {
        if (value.IsObject)
        {
            return ToString(ObjectValue(value));
        }

        var culture = Culture;
        return value.Type switch
        {
            VarType.String => value.AsString(),
            VarType.Empty => string.Empty,
            VarType.Integer => value.AsInt16().ToString(CultureInfo.InvariantCulture),
            VarType.Long => value.AsInt32().ToString(CultureInfo.InvariantCulture),
            VarType.LongLong => value.AsInt64().ToString(CultureInfo.InvariantCulture),
            VarType.Byte => value.AsByte().ToString(CultureInfo.InvariantCulture),
            VarType.Boolean => value.AsBoolean() ? "True" : "False",
            VarType.Single => NumberText.Format(value.AsSingle(), NumberText.SingleDigits, culture.NumberFormat.NumberDecimalSeparator),
            VarType.Double => NumberText.Format(value.AsDouble(), NumberText.DoubleDigits, culture.NumberFormat.NumberDecimalSeparator),
            VarType.Currency => FixedPointText(value.AsCurrency().ToDecimal(), culture),
            VarType.Decimal => FixedPointText(value.AsDecimal(), culture),
            VarType.Date => DateText.Format(value.AsDate(), culture),
            VarType.Error => value.AsError().ToString(),
            _ => throw NotConvertible(value),
        };
    }

    /// <summary>Coerces to the given declared type; Variant returns the value unchanged.</summary>
    public static Variant ToType(in Variant value, VarType destination) => destination switch
    {
        VarType.Variant => value,
        VarType.Integer => Variant.FromInt16(ToInt16(value)),
        VarType.Long => Variant.FromInt32(ToInt32(value)),
        VarType.LongLong => Variant.FromInt64(ToInt64(value)),
        VarType.Byte => Variant.FromByte(ToByte(value)),
        VarType.Single => Variant.FromSingle(ToSingle(value)),
        VarType.Double => Variant.FromDouble(ToDouble(value)),
        VarType.Currency => Variant.FromCurrency(ToCurrency(value)),
        VarType.Decimal => Variant.FromDecimal(ToDecimal(value)),
        VarType.Date => Variant.FromDate(ToDate(value)),
        VarType.Boolean => Variant.FromBoolean(ToBoolean(value)),
        VarType.String => Variant.FromString(ToString(value)),
        VarType.Object => value.IsObject ? value : throw VbaErrors.TypeMismatch(),
        _ => throw new ArgumentOutOfRangeException(nameof(destination), destination, "Not a declared type."),
    };

    /// <summary>Coerces a value stored into an array element of the given element type.</summary>
    public static Variant ToElementType(in Variant value, VarType elementType) => elementType switch
    {
        VarType.Variant => value,
        VarType.Object => value.IsObject ? value : throw new VbaException(VbaErrors.ObjectRequired),
        VarType.UserDefinedType => throw VbaErrors.TypeMismatch(),
        _ => ToType(value, elementType),
    };

    /// <summary>An early-bound object reference that must not be Nothing: error 91 otherwise (MS-VBAL 5.6.12).</summary>
    public static T Require<T>(T? value)
        where T : class =>
        value ?? throw VbaErrors.ObjectVariableNotSet();

    /// <summary>Let-assignment of any value (MS-VBAL 5.6.9.3): an object stands for its default member's value. The store that follows copies an array (<see cref="ObjectRefs.Own(in Variant)"/>).</summary>
    public static Variant LetValue(in Variant value) => value.IsObject ? ObjectValue(value) : value;

    /// <summary>
    /// The value an object reference stands for when a value is needed (MS-VBAL 5.6.9.3): its default
    /// member invoked with no arguments, followed until a value comes out. Nothing raises 91; an
    /// object without a default member raises 438, as VBA reports it (Objects golden).
    /// </summary>
    internal static Variant ObjectValue(in Variant value)
    {
        var current = value;
        for (var depth = 0; depth < 8 && current.IsObject; depth++)
        {
            current = current.AsObject() switch
            {
                null => throw VbaErrors.ObjectVariableNotSet(),
                IDispatchObject dispatch => dispatch.Invoke(LateBound.DefaultMember, InvokeKind.PropertyGet | InvokeKind.Method, []),
                // Collection's default member is Item, which needs an index (Objects golden); Err's is Number.
                Library.Collection => throw new VbaException(VbaErrors.WrongNumberOfArguments),
                Library.ErrObject err => Variant.FromInt32(err.Number),
                _ => throw new VbaException(VbaErrors.ObjectDoesNotSupportMember),
            };
        }

        return current.IsObject ? throw new VbaException(VbaErrors.ObjectDoesNotSupportMember) : current;
    }

    /// <summary>A condition (MS-VBAL 5.4.2.8 If, 5.4.2.2 Do): Null counts as False; anything else is let-coerced to Boolean.</summary>
    public static bool ToCondition(in Variant value) => !value.IsNull && ToBoolean(value);

    /// <summary>An object reference of a given class from a Variant or Object: Nothing gives null, a non-object raises 424, another class raises 13.</summary>
    public static T? ToObject<T>(in Variant value)
        where T : class
    {
        if (!value.IsObject)
        {
            throw new VbaException(VbaErrors.ObjectRequired);
        }

        return value.AsObject() switch
        {
            null => null,
            T typed => typed,
            _ => throw VbaErrors.TypeMismatch(),
        };
    }

    /// <summary>Let-coercion to a declared array type (MS-VBAL 5.5.1.2.5): the value must be an array of the element type; a String assigned to a Byte array becomes its bytes.</summary>
    public static VbaArray ToArray(in Variant value, VarType elementType)
    {
        if (value.IsString && elementType == VarType.Byte)
        {
            return ObjectRefs.Owned(VbaArray.FromBytes(value.AsVbaString().Bytes));
        }

        if (!value.IsArray)
        {
            throw VbaErrors.TypeMismatch();
        }

        var array = value.AsArray();
        if (elementType != VarType.Variant && array.ElementType != elementType)
        {
            throw VbaErrors.TypeMismatch();
        }

        // A view of the array; the store that follows takes the copy (MS-VBAL 5.5.1.2.5).
        return array;
    }

    /// <summary>Banker's rounding (MS-VBAL 5.5.1.2.1.1): half rounds to the even neighbor.</summary>
    public static double RoundHalfEven(double value) => Math.Round(value, MidpointRounding.ToEven);

    private static long ToIntegral(in Variant value, long min, long max)
    {
        if (value.IsObject)
        {
            return ToIntegral(ObjectValue(value), min, max);
        }

        long result;
        switch (value.Type)
        {
            case VarType.Integer:
                result = value.AsInt16();
                break;
            case VarType.Long:
                result = value.AsInt32();
                break;
            case VarType.LongLong:
                result = value.AsInt64();
                break;
            case VarType.Byte:
                result = value.AsByte();
                break;
            case VarType.Boolean:
                result = value.BooleanBits;
                break;
            case VarType.Empty:
                result = 0;
                break;
            case VarType.Single:
            case VarType.Double:
            case VarType.Date:
                result = RoundToInt64(ToDouble(value));
                break;
            case VarType.Currency:
                result = RoundToInt64(value.AsCurrency().ToDecimal());
                break;
            case VarType.Decimal:
                result = RoundToInt64(value.AsDecimal());
                break;
            case VarType.String:
                if (!NumberText.TryParse(value.AsVbaString().Chars, Culture, out var exact, out var wide, out var isExact))
                {
                    throw VbaErrors.TypeMismatch();
                }

                result = isExact ? RoundToInt64(exact) : RoundToInt64(wide);
                break;
            default:
                throw NotConvertible(value);
        }

        if (result < min || result > max)
        {
            throw VbaErrors.Overflow();
        }

        return result;
    }

    private static long RoundToInt64(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            throw VbaErrors.Overflow();
        }

        var rounded = Math.Round(value, MidpointRounding.ToEven);
        if (rounded < long.MinValue || rounded >= 9223372036854775808.0)
        {
            throw VbaErrors.Overflow();
        }

        return (long)rounded;
    }

    private static long RoundToInt64(decimal value)
    {
        var rounded = decimal.Round(value, 0, MidpointRounding.ToEven);
        if (rounded < long.MinValue || rounded > long.MaxValue)
        {
            throw VbaErrors.Overflow();
        }

        return (long)rounded;
    }

    private static string FixedPointText(decimal value, CultureInfo culture)
    {
        var text = DecimalText.Normalize(value).ToString(CultureInfo.InvariantCulture);
        var separator = culture.NumberFormat.NumberDecimalSeparator;
        return separator == "." ? text : text.Replace(".", separator, StringComparison.Ordinal);
    }

    /// <summary>Null raises 94, Nothing raises 91, other objects 438, everything else 13.</summary>
    internal static VbaException NotConvertible(in Variant value) => value.Type switch
    {
        VarType.Null => VbaErrors.InvalidUseOfNull(),
        VarType.Object when value.IsNothing => VbaErrors.ObjectVariableNotSet(),
        VarType.Object => new VbaException(VbaErrors.ObjectDoesNotSupportMember),
        _ => VbaErrors.TypeMismatch(),
    };
}
