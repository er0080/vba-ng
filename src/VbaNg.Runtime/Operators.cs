using System.Globalization;

namespace VbaNg.Runtime;

/// <summary>Which operands of an operator have a declared type of Variant (MS-VBAL 5.6.9.3): overflow widens only then.</summary>
[Flags]
public enum DeclaredTypes
{
    None = 0,
    LeftVariant = 1,
    RightVariant = 2,
    BothVariant = LeftVariant | RightVariant,

    /// <summary>The operation is typed Double or Single: a division by zero or an overflow hands back the IEEE result and is raised afterwards (docs/vba-quirks.md, floating point).</summary>
    Floating = 4,
}

/// <summary>The Option Compare mode of the module an operator or function runs in (MS-VBAL 5.2.1.1).</summary>
public enum CompareMode
{
    Binary = 0,
    Text = 1,
}

/// <summary>
/// The VBA operators on Variants (MS-VBAL 5.6.9). Generated code calls these for every operator
/// whose operands are not both statically typed primitives, passing which operands are declared
/// Variant, because that changes what an overflow does.
/// </summary>
public static class Operators
{
    /// <summary>MS-VBAL 5.6.9.3.1 + operator: adds numbers, concatenates two strings, and treats Empty as 0 or "".</summary>
    public static Variant Add(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None)
    {
        if (left.IsObject || right.IsObject)
        {
            return Add(Value(left), Value(right), declared);
        }

        // Two strings (or a string and Empty) concatenate; a string with a number adds.
        if ((left.IsString || left.IsEmpty) && (right.IsString || right.IsEmpty) && (left.IsString || right.IsString))
        {
            return Variant.FromVbaString(VbaString.Concat(
                left.IsEmpty ? ReadOnlySpan<byte>.Empty : left.AsVbaString().Bytes,
                right.IsEmpty ? ReadOnlySpan<byte>.Empty : right.AsVbaString().Bytes));
        }

        return Arithmetic(left, right, declared, ArithmeticKind.Add);
    }

    /// <summary>MS-VBAL 5.6.9.3.2 - operator.</summary>
    public static Variant Subtract(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None) =>
        Arithmetic(left, right, declared, ArithmeticKind.Subtract);

    /// <summary>MS-VBAL 5.6.9.3.3 * operator.</summary>
    public static Variant Multiply(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None) =>
        Arithmetic(left, right, declared, ArithmeticKind.Multiply);

    /// <summary>MS-VBAL 5.6.9.3.4 / operator: Double unless Single or Decimal operands say otherwise; 0/0 is error 6, x/0 error 11.</summary>
    public static Variant Divide(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None)
    {
        if (left.IsObject || right.IsObject)
        {
            return Divide(Value(left), Value(right), declared);
        }

        CheckOperands(left, right);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        var l = left.Type;
        var r = right.Type;
        VarType effective;
        if (l == VarType.Decimal || r == VarType.Decimal)
        {
            effective = VarType.Decimal;
        }
        else if (l == VarType.Single && r is VarType.Single or VarType.Byte or VarType.Boolean or VarType.Integer or VarType.Empty
            || r == VarType.Single && l is VarType.Byte or VarType.Boolean or VarType.Integer or VarType.Empty)
        {
            effective = VarType.Single;
        }
        else
        {
            effective = VarType.Double;
        }

        if (effective == VarType.Decimal)
        {
            var dividend = Coerce.ToDecimal(left);
            var divisor = Coerce.ToDecimal(right);
            if (divisor == 0)
            {
                throw VbaErrors.DivisionByZero();
            }

            try
            {
                return Variant.FromDecimal(dividend / divisor);
            }
            catch (OverflowException)
            {
                throw VbaErrors.Overflow();
            }
        }

        var x = Coerce.ToDouble(left);
        var y = Coerce.ToDouble(right);
        if (y == 0)
        {
            var number = x != 0 || (l is VarType.Single or VarType.Double or VarType.String or VarType.Date && r == VarType.Empty)
                ? VbaErrors.DivisionByZeroNumber
                : VbaErrors.OverflowNumber;
            if ((declared & DeclaredTypes.Floating) != 0)
            {
                // A typed Double or Single division stores its IEEE result before the error is raised (Errors golden).
                pendingFloating = number;
                return effective == VarType.Single ? Variant.FromSingle((float)x / (float)y) : Variant.FromDouble(x / y);
            }

            throw new VbaException(number);
        }

        if (effective == VarType.Single)
        {
            return Widen(Variant.FromSingle((float)x / (float)y), VarType.Single, declared, x / y);
        }

        return Widen(Variant.FromDouble(x / y), VarType.Double, declared, x / y);
    }

    /// <summary>MS-VBAL 5.6.9.3.5 \ operator: operands round to integers first; division by zero is error 11.</summary>
    public static Variant IntegerDivide(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None) =>
        IntegralOperation(left, right, declared, modulo: false);

    /// <summary>MS-VBAL 5.6.9.3.5 Mod operator: the remainder keeps the sign of the dividend, as Excel does.</summary>
    public static Variant Modulo(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None) =>
        IntegralOperation(left, right, declared, modulo: true);

    /// <summary>MS-VBAL 5.6.9.3.6 ^ operator: always Double; a negative base needs an integral exponent (error 5).</summary>
    public static Variant Power(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None)
    {
        if (left.IsObject || right.IsObject)
        {
            return Power(Value(left), Value(right), declared);
        }

        CheckOperands(left, right);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        var x = Coerce.ToDouble(left);
        var y = Coerce.ToDouble(right);
        if ((x < 0 && y != Math.Floor(y)) || (x == 0 && y < 0))
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var result = Math.Pow(x, y);
        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            if ((declared & DeclaredTypes.Floating) != 0)
            {
                pendingFloating = VbaErrors.OverflowNumber;
                return Variant.FromDouble(result);
            }

            throw VbaErrors.Overflow();
        }

        return Variant.FromDouble(result);
    }

    /// <summary>
    /// The floating-point error of the statement in progress (docs/vba-quirks.md, "Runtime,
    /// floating point"): a Double or Single operation that divides by zero or overflows hands
    /// back its IEEE result when the caller passes <see cref="DeclaredTypes.Floating"/>, and the
    /// error waits here until generated code checks, after the store of a typed result or before
    /// any other use of it.
    /// </summary>
    [ThreadStatic]
    private static int pendingFloating;

    /// <summary>Raises the floating-point error the statement left pending, if any.</summary>
    public static void CheckFloating()
    {
        var number = pendingFloating;
        if (number != 0)
        {
            pendingFloating = 0;
            throw new VbaException(number);
        }
    }

    /// <summary>Raises the pending floating-point error before <paramref name="value"/> is used anywhere but a typed store.</summary>
    public static Variant CheckFloating(in Variant value)
    {
        CheckFloating();
        return value;
    }

    /// <summary>A typed Double operation's result, once the pending floating-point error is raised.</summary>
    public static double CheckFloating(double value)
    {
        CheckFloating();
        return value;
    }

    // Typed operands through C# arithmetic (ROADMAP.md D-J, M7 E): generated code calls these when both operands
    // of +, -, or * have declared numeric types, in the result type MS-VBAL 5.6.9.3 gives them, and for / when that
    // type is Double (D-M). An integral result that does not fit raises 6; a Double's overflow is left pending for the
    // statement's check, as the Variant path does for typed operands (Operators and Errors goldens).

    public static byte AddByte(byte left, byte right) => FitByte(left + right);

    public static byte SubtractByte(byte left, byte right) => FitByte(left - right);

    public static byte MultiplyByte(byte left, byte right) => FitByte(left * right);

    public static short AddInt16(short left, short right) => FitInt16(left + right);

    public static short SubtractInt16(short left, short right) => FitInt16(left - right);

    public static short MultiplyInt16(short left, short right) => FitInt16(left * right);

    public static int AddInt32(int left, int right) => FitInt32((long)left + right);

    public static int SubtractInt32(int left, int right) => FitInt32((long)left - right);

    public static int MultiplyInt32(int left, int right) => FitInt32((long)left * right);

    public static long AddInt64(long left, long right)
    {
        var sum = unchecked(left + right);
        return ((left ^ sum) & (right ^ sum)) < 0 ? throw VbaErrors.Overflow() : sum;
    }

    public static long SubtractInt64(long left, long right)
    {
        var difference = unchecked(left - right);
        return ((left ^ right) & (left ^ difference)) < 0 ? throw VbaErrors.Overflow() : difference;
    }

    public static long MultiplyInt64(long left, long right)
    {
        var product = (Int128)left * right;
        return product < long.MinValue || product > long.MaxValue ? throw VbaErrors.Overflow() : (long)product;
    }

    public static double AddDouble(double left, double right) => Floated(left + right, left, right);

    public static double SubtractDouble(double left, double right) => Floated(left - right, left, right);

    public static double MultiplyDouble(double left, double right) => Floated(left * right, left, right);

    /// <summary>A / whose quotient is Double: by zero leaves 11 pending (0/0 leaves 6), and a quotient past the Double range leaves 6, after the IEEE result is stored, as Divide does under the Floating flag.</summary>
    public static double DivideDouble(double left, double right)
    {
        var quotient = left / right;
        if (right == 0)
        {
            pendingFloating = left != 0 ? VbaErrors.DivisionByZeroNumber : VbaErrors.OverflowNumber;
        }
        else if (double.IsInfinity(quotient) || double.IsNaN(quotient))
        {
            pendingFloating = VbaErrors.OverflowNumber;
        }

        return quotient;
    }

    private static byte FitByte(int value) => value is >= byte.MinValue and <= byte.MaxValue ? (byte)value : throw VbaErrors.Overflow();

    private static short FitInt16(int value) => value is >= short.MinValue and <= short.MaxValue ? (short)value : throw VbaErrors.Overflow();

    private static int FitInt32(long value) => value is >= int.MinValue and <= int.MaxValue ? (int)value : throw VbaErrors.Overflow();

    /// <summary>Finite operands that overflowed leave error 6 pending and hand back the IEEE result; an infinite operand from an earlier deferred error carries through (Errors golden).</summary>
    private static double Floated(double result, double left, double right)
    {
        if ((double.IsInfinity(result) || double.IsNaN(result)) && double.IsFinite(left) && double.IsFinite(right))
        {
            pendingFloating = VbaErrors.OverflowNumber;
        }

        return result;
    }

    /// <summary>MS-VBAL 5.6.9.3.7? Unary - operator: Byte negates as Integer; Date negates as Double and comes back as Date.</summary>
    public static Variant Negate(in Variant operand, bool declaredVariant = false)
    {
        if (operand.IsObject)
        {
            return Negate(Value(operand), declaredVariant);
        }

        CheckOperand(operand);
        switch (operand.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Byte:
                return Variant.FromInt16((short)-operand.AsByte());
            case VarType.Boolean:
            case VarType.Integer:
            case VarType.Empty:
                {
                    int value = -Coerce.ToInt16(operand);
                    return value is >= short.MinValue and <= short.MaxValue
                        ? Variant.FromInt16((short)value)
                        : declaredVariant ? Variant.FromInt32(value) : throw VbaErrors.Overflow();
                }

            case VarType.Long:
                {
                    long value = -(long)operand.AsInt32();
                    return value is >= int.MinValue and <= int.MaxValue
                        ? Variant.FromInt32((int)value)
                        : declaredVariant ? Variant.FromDouble(value) : throw VbaErrors.Overflow();
                }

            case VarType.LongLong:
                return operand.AsInt64() == long.MinValue ? throw VbaErrors.Overflow() : Variant.FromInt64(-operand.AsInt64());
            case VarType.Single:
                return Variant.FromSingle(-operand.AsSingle());
            case VarType.Double:
            case VarType.String:
                return Variant.FromDouble(-Coerce.ToDouble(operand));
            case VarType.Currency:
                return Variant.FromCurrency(Currency.Negate(operand.AsCurrency()));
            case VarType.Decimal:
                return Variant.FromDecimal(-operand.AsDecimal());
            case VarType.Date:
                {
                    var value = -operand.AsDate().Serial;
                    try
                    {
                        return Variant.FromDate(VbaDate.FromSerial(value));
                    }
                    catch (VbaException) when (declaredVariant)
                    {
                        return Variant.FromDouble(value);
                    }
                }

            default:
                throw Coerce.NotConvertible(operand);
        }
    }

    /// <summary>MS-VBAL 5.6.9.4 &amp; operator: Null &amp; Null is Null, Null alone is "", both operands become strings.</summary>
    public static Variant Concatenate(in Variant left, in Variant right)
    {
        if (left.IsObject || right.IsObject)
        {
            return Concatenate(Value(left), Value(right));
        }

        CheckOperands(left, right);
        if (left.IsNull && right.IsNull)
        {
            return Variant.Null;
        }

        // The bytes of both, so a byte string of odd length concatenates exactly (Strings golden).
        var l = left.IsNull ? default : Coerce.ToText(left);
        var r = right.IsNull ? default : Coerce.ToText(right);
        return Variant.FromVbaString(VbaString.Concat(l.Bytes, r.Bytes));
    }

    public static Variant Equal(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c == 0);

    public static Variant NotEqual(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c != 0);

    public static Variant LessThan(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c < 0);

    public static Variant GreaterThan(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c > 0);

    public static Variant LessThanOrEqual(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c <= 0);

    public static Variant GreaterThanOrEqual(in Variant left, in Variant right, DeclaredTypes declared = DeclaredTypes.None, CompareMode mode = CompareMode.Binary) =>
        Relational(left, right, declared, mode, static c => c >= 0);

    /// <summary>MS-VBAL 5.6.9.5 Is operator: reference equality of two object references.</summary>
    /// <summary>An operand as a value: an object stands for its default member (MS-VBAL 5.6.9.3), Nothing raises 91, no default member raises 438.</summary>
    private static Variant Value(in Variant operand) => operand.IsObject ? Coerce.ObjectValue(operand) : operand;

    public static Variant Is(in Variant left, in Variant right)
    {
        if (!left.IsObject || !right.IsObject)
        {
            throw new VbaException(VbaErrors.ObjectRequired);
        }

        // One interface pointer is one object, Nothing included; two pointers may still be one object behind two interfaces.
        if (left.InterfacePointer == right.InterfacePointer)
        {
            return Variant.True;
        }

        var leftObject = left.AsObject();
        var rightObject = right.AsObject();
        return Variant.FromBoolean(leftObject switch
        {
            // COM identity: two wrappers of one object are the same object (ARCHITECTURE.md section 6).
            IDispatchObject dispatch => dispatch.IsSameObject(rightObject),
            _ => rightObject is IDispatchObject other ? other.IsSameObject(leftObject) : ReferenceEquals(leftObject, rightObject),
        });
    }

    /// <summary>MS-VBAL 5.6.9.6 Like operator.</summary>
    public static Variant Like(in Variant text, in Variant pattern, CompareMode mode = CompareMode.Binary)
    {
        if (text.IsObject || pattern.IsObject)
        {
            return Like(Value(text), Value(pattern), mode);
        }

        CheckOperands(text, pattern);
        if (text.IsNull || pattern.IsNull)
        {
            return Variant.Null;
        }

        return Variant.FromBoolean(LikePattern.IsMatch(Coerce.ToText(text).Chars, Coerce.ToText(pattern).Chars, mode));
    }

    /// <summary>MS-VBAL 5.6.9.7.1 Not operator: bitwise on the operand's integral effective type.</summary>
    public static Variant Not(in Variant operand)
    {
        if (operand.IsObject)
        {
            return Not(Value(operand));
        }

        CheckOperand(operand);
        switch (operand.Type)
        {
            case VarType.Null:
                return Variant.Null;
            case VarType.Byte:
                return Variant.FromByte((byte)~operand.AsByte());
            case VarType.Boolean:
                return Variant.FromBoolean(!operand.AsBoolean());
            case VarType.Integer:
            case VarType.Empty:
                return Variant.FromInt16((short)~Coerce.ToInt16(operand));
            case VarType.LongLong:
                return Variant.FromInt64(~operand.AsInt64());
            default:
                return Variant.FromInt32(~Coerce.ToInt32(operand));
        }
    }

    /// <summary>MS-VBAL 5.6.9.7.2 And operator: Null And False is False, Null And nonzero is Null.</summary>
    public static Variant And(in Variant left, in Variant right) =>
        Logical(left, right, static (a, b) => a & b, LogicalKind.And);

    /// <summary>MS-VBAL 5.6.9.7.3 Or operator: Null Or True is True.</summary>
    public static Variant Or(in Variant left, in Variant right) =>
        Logical(left, right, static (a, b) => a | b, LogicalKind.Or);

    /// <summary>MS-VBAL 5.6.9.7.4 Xor operator: Null with anything is Null.</summary>
    public static Variant Xor(in Variant left, in Variant right) =>
        Logical(left, right, static (a, b) => a ^ b, LogicalKind.Other);

    /// <summary>MS-VBAL 5.6.9.7.5 Eqv operator.</summary>
    public static Variant Eqv(in Variant left, in Variant right) =>
        Logical(left, right, static (a, b) => ~(a ^ b), LogicalKind.Other);

    /// <summary>MS-VBAL 5.6.9.7.6 Imp operator: False Imp Null and Null Imp True are True.</summary>
    public static Variant Imp(in Variant left, in Variant right) =>
        Logical(left, right, static (a, b) => ~a | b, LogicalKind.Imp);

    private enum ArithmeticKind
    {
        Add,
        Subtract,
        Multiply,
    }

    private enum LogicalKind
    {
        And,
        Or,
        Imp,
        Other,
    }

    /// <summary>The binary arithmetic effective value type (MS-VBAL 5.6.9.3 table), for operands that are not Null.</summary>
    private static VarType EffectiveArithmeticType(VarType l, VarType r)
    {
        if (l == VarType.Decimal || r == VarType.Decimal)
        {
            return VarType.Decimal;
        }

        if (l == VarType.Date || r == VarType.Date)
        {
            return VarType.Date;
        }

        if (l == VarType.Currency || r == VarType.Currency)
        {
            return VarType.Currency;
        }

        if (l is VarType.Double or VarType.String || r is VarType.Double or VarType.String)
        {
            return VarType.Double;
        }

        if (l == VarType.Single || r == VarType.Single)
        {
            var other = l == VarType.Single ? r : l;
            return other is VarType.Long or VarType.LongLong ? VarType.Double : VarType.Single;
        }

        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        if (l == VarType.Long || r == VarType.Long)
        {
            return VarType.Long;
        }

        if ((l == VarType.Byte && r is VarType.Byte or VarType.Empty) || (r == VarType.Byte && l == VarType.Empty))
        {
            return VarType.Byte;
        }

        return VarType.Integer;
    }

    private static Variant Arithmetic(in Variant left, in Variant right, DeclaredTypes declared, ArithmeticKind kind)
    {
        if (left.IsObject || right.IsObject)
        {
            return Arithmetic(Value(left), Value(right), declared, kind);
        }

        CheckOperands(left, right);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        var effective = EffectiveArithmeticType(left.Type, right.Type);
        if (effective == VarType.Date && (kind == ArithmeticKind.Multiply || (kind == ArithmeticKind.Subtract && left.IsDate && right.IsDate)))
        {
            // Only + and - keep a Date; the difference of two Dates and any product is a Double (Operators goldens).
            effective = VarType.Double;
        }
        else if (effective == VarType.Currency && kind == ArithmeticKind.Multiply && (left.Type.IsFloatingPoint() || right.Type.IsFloatingPoint()))
        {
            // Currency times a floating-point value is a Double, although their sum is a Currency (Information golden).
            effective = VarType.Double;
        }

        switch (effective)
        {
            case VarType.Byte:
            case VarType.Integer:
            case VarType.Long:
                {
                    var a = Coerce.ToInt64(left);
                    var b = Coerce.ToInt64(right);
                    var result = kind switch
                    {
                        ArithmeticKind.Add => a + b,
                        ArithmeticKind.Subtract => a - b,
                        _ => a * b,
                    };
                    return Narrow(result, effective, declared);
                }

            case VarType.LongLong:
                {
                    var a = Coerce.ToInt64(left);
                    var b = Coerce.ToInt64(right);
                    try
                    {
                        var result = kind switch
                        {
                            ArithmeticKind.Add => checked(a + b),
                            ArithmeticKind.Subtract => checked(a - b),
                            _ => checked(a * b),
                        };
                        return Variant.FromInt64(result);
                    }
                    catch (OverflowException)
                    {
                        throw VbaErrors.Overflow();
                    }
                }

            case VarType.Single:
                {
                    var a = Coerce.ToSingle(left);
                    var b = Coerce.ToSingle(right);
                    var result = kind switch
                    {
                        ArithmeticKind.Add => a + b,
                        ArithmeticKind.Subtract => a - b,
                        _ => a * b,
                    };
                    var wide = kind switch
                    {
                        ArithmeticKind.Add => (double)a + b,
                        ArithmeticKind.Subtract => (double)a - b,
                        _ => (double)a * b,
                    };
                    return Widen(Variant.FromSingle(result), VarType.Single, declared, wide);
                }

            case VarType.Double:
            case VarType.Date:
                {
                    var a = Coerce.ToDouble(left);
                    var b = Coerce.ToDouble(right);
                    var result = kind switch
                    {
                        ArithmeticKind.Add => a + b,
                        ArithmeticKind.Subtract => a - b,
                        _ => a * b,
                    };
                    if ((double.IsInfinity(result) || double.IsNaN(result)) && double.IsFinite(a) && double.IsFinite(b))
                    {
                        // Finite operands overflowed; an infinite operand from an earlier deferred error just carries through (Errors golden).
                        return Overflowed(Variant.FromDouble(result), declared);
                    }

                    if (effective == VarType.Double)
                    {
                        return Variant.FromDouble(result);
                    }

                    try
                    {
                        return Variant.FromDate(VbaDate.FromSerial(result));
                    }
                    catch (VbaException) when ((declared & DeclaredTypes.BothVariant) != 0)
                    {
                        return Variant.FromDouble(result);
                    }
                }

            case VarType.Currency:
                {
                    var a = Coerce.ToCurrency(left);
                    var b = Coerce.ToCurrency(right);
                    return Variant.FromCurrency(kind switch
                    {
                        ArithmeticKind.Add => Currency.Add(a, b),
                        ArithmeticKind.Subtract => Currency.Subtract(a, b),
                        _ => Currency.Multiply(a, b),
                    });
                }

            default:
                {
                    var a = Coerce.ToDecimal(left);
                    var b = Coerce.ToDecimal(right);
                    try
                    {
                        return Variant.FromDecimal(kind switch
                        {
                            ArithmeticKind.Add => a + b,
                            ArithmeticKind.Subtract => a - b,
                            _ => a * b,
                        });
                    }
                    catch (OverflowException)
                    {
                        throw VbaErrors.Overflow();
                    }
                }
        }
    }

    /// <summary>Fits an integral result into the effective type, widening to Long or Double when a Variant operand allows it (MS-VBAL 5.6.9.3).</summary>
    private static Variant Narrow(long result, VarType effective, DeclaredTypes declared)
    {
        switch (effective)
        {
            case VarType.Byte when result is >= byte.MinValue and <= byte.MaxValue:
                return Variant.FromByte((byte)result);
            case VarType.Byte:
            case VarType.Integer when result is >= short.MinValue and <= short.MaxValue:
                if (effective == VarType.Integer)
                {
                    return Variant.FromInt16((short)result);
                }

                break;
            case VarType.Integer:
                break;
            case VarType.Long when result is >= int.MinValue and <= int.MaxValue:
                return Variant.FromInt32((int)result);
        }

        if ((declared & DeclaredTypes.BothVariant) == 0)
        {
            throw VbaErrors.Overflow();
        }

        if (result is >= short.MinValue and <= short.MaxValue)
        {
            return Variant.FromInt16((short)result);
        }

        if (result is >= int.MinValue and <= int.MaxValue)
        {
            return Variant.FromInt32((int)result);
        }

        return Variant.FromDouble(result);
    }

    /// <summary>A Single or Double result that overflowed: error 6 for declared operands, Double for a Single with a Variant operand.</summary>
    private static Variant Widen(Variant result, VarType effective, DeclaredTypes declared, double wide)
    {
        if (effective == VarType.Single)
        {
            var single = result.AsSingle();
            if (!float.IsInfinity(single) && !float.IsNaN(single))
            {
                return result;
            }

            if ((declared & DeclaredTypes.BothVariant) != 0 && !double.IsInfinity(wide) && !double.IsNaN(wide))
            {
                return Variant.FromDouble(wide);
            }

            return Overflowed(result, declared);
        }

        var value = result.AsDouble();
        if (double.IsInfinity(value) || double.IsNaN(value))
        {
            return Overflowed(result, declared);
        }

        return result;
    }

    /// <summary>An infinite or undefined floating-point result: raised now, or handed back with the error pending when the operation is typed (Errors golden).</summary>
    private static Variant Overflowed(Variant result, DeclaredTypes declared)
    {
        if ((declared & DeclaredTypes.Floating) != 0)
        {
            pendingFloating = VbaErrors.OverflowNumber;
            return result;
        }

        throw VbaErrors.Overflow();
    }

    private static Variant IntegralOperation(in Variant left, in Variant right, DeclaredTypes declared, bool modulo)
    {
        if (left.IsObject || right.IsObject)
        {
            return IntegralOperation(Value(left), Value(right), declared, modulo);
        }

        CheckOperands(left, right);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        var l = left.Type;
        var r = right.Type;
        VarType effective;
        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            effective = VarType.LongLong;
        }
        else if ((l is VarType.Boolean or VarType.Integer) && r is VarType.Single or VarType.Double or VarType.String or VarType.Currency or VarType.Date or VarType.Decimal)
        {
            effective = VarType.Integer;
        }
        else if (l.IsFloatingPoint() || l.IsFixedPoint() || l is VarType.String or VarType.Date
            || r.IsFloatingPoint() || r.IsFixedPoint() || r is VarType.String or VarType.Date)
        {
            effective = VarType.Long;
        }
        else if ((l == VarType.Byte && r == VarType.Empty) || (l == VarType.Empty && r == VarType.Byte))
        {
            effective = VarType.Integer;
        }
        else
        {
            effective = EffectiveArithmeticType(l, r);
        }

        var a = effective == VarType.LongLong ? Coerce.ToInt64(left) : Coerce.ToInt32(left);
        var b = effective == VarType.LongLong ? Coerce.ToInt64(right) : Coerce.ToInt32(right);
        if (b == 0)
        {
            throw VbaErrors.DivisionByZero();
        }

        if (effective == VarType.LongLong && a == long.MinValue && b == -1)
        {
            throw VbaErrors.Overflow();
        }

        var result = modulo ? a % b : a / b;
        return effective == VarType.LongLong ? Variant.FromInt64(result) : Narrow(result, effective, declared);
    }

    private static Variant Relational(in Variant left, in Variant right, DeclaredTypes declared, CompareMode mode, Func<int, bool> test)
    {
        if (left.IsObject || right.IsObject)
        {
            return Relational(Value(left), Value(right), declared, mode, test);
        }

        CheckOperands(left, right, allowError: true);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        var l = left.Type;
        var r = right.Type;
        if (l == VarType.Error || r == VarType.Error)
        {
            if (l != r)
            {
                throw VbaErrors.TypeMismatch();
            }

            return Variant.FromBoolean(test(left.AsError().Number.CompareTo(right.AsError().Number)));
        }

        // Both declared Variant, one a String and the other a number: the number sorts first (MS-VBAL 5.6.9.5 exception).
        if (declared == DeclaredTypes.BothVariant && ((l == VarType.String && r.IsNumeric()) || (r == VarType.String && l.IsNumeric())))
        {
            return Variant.FromBoolean(test(l == VarType.String ? 1 : -1));
        }

        // A declared String against a Variant compares as text whatever the Variant holds, Null aside (MS-VBAL 5.6.9.5;
        // Operators golden): a Variant holding 10 is less than "9".
        var variants = declared & DeclaredTypes.BothVariant;
        if ((variants == DeclaredTypes.LeftVariant && r == VarType.String) || (variants == DeclaredTypes.RightVariant && l == VarType.String))
        {
            return Variant.FromBoolean(test(TextCompare.Compare(Coerce.ToText(left), Coerce.ToText(right), mode)));
        }

        var effective = EffectiveRelationalType(l, r);
        int comparison;
        switch (effective)
        {
            case VarType.String:
                comparison = TextCompare.Compare(Coerce.ToText(left), Coerce.ToText(right), mode);
                break;
            case VarType.Boolean:
                {
                    // True is less than False: compare the underlying -1 and 0.
                    var a = Coerce.ToBoolean(left) ? -1 : 0;
                    var b = Coerce.ToBoolean(right) ? -1 : 0;
                    comparison = a.CompareTo(b);
                    break;
                }

            case VarType.Byte:
            case VarType.Integer:
            case VarType.Long:
            case VarType.LongLong:
                comparison = Coerce.ToInt64(left).CompareTo(Coerce.ToInt64(right));
                break;
            case VarType.Single:
                {
                    var a = Coerce.ToSingle(left);
                    var b = Coerce.ToSingle(right);
                    if (float.IsNaN(a) || float.IsNaN(b))
                    {
                        throw VbaErrors.Overflow();
                    }

                    comparison = a.CompareTo(b);
                    break;
                }

            case VarType.Double:
            case VarType.Date:
                {
                    var a = effective == VarType.Date ? Coerce.ToDate(left).Serial : Coerce.ToDouble(left);
                    var b = effective == VarType.Date ? Coerce.ToDate(right).Serial : Coerce.ToDouble(right);
                    if (double.IsNaN(a) || double.IsNaN(b))
                    {
                        throw VbaErrors.Overflow();
                    }

                    comparison = a.CompareTo(b);
                    break;
                }

            case VarType.Currency:
                comparison = Coerce.ToCurrency(left).CompareTo(Coerce.ToCurrency(right));
                break;
            default:
                comparison = Coerce.ToDecimal(left).CompareTo(Coerce.ToDecimal(right));
                break;
        }

        return Variant.FromBoolean(test(comparison));
    }

    /// <summary>The relational effective value type (MS-VBAL 5.6.9.5 table), for operands that are not Null or Error.</summary>
    private static VarType EffectiveRelationalType(VarType l, VarType r)
    {
        if (l == VarType.Decimal || r == VarType.Decimal)
        {
            return VarType.Decimal;
        }

        if (l == VarType.Date || r == VarType.Date)
        {
            return VarType.Date;
        }

        if (l == VarType.Currency || r == VarType.Currency)
        {
            return VarType.Currency;
        }

        if ((l == VarType.String || l == VarType.Empty) && (r == VarType.String || r == VarType.Empty) && (l == VarType.String || r == VarType.String))
        {
            return VarType.String;
        }

        if (l == VarType.Boolean && r is VarType.Boolean or VarType.String || r == VarType.Boolean && l is VarType.Boolean or VarType.String)
        {
            return VarType.Boolean;
        }

        if (l == VarType.Double || r == VarType.Double)
        {
            var other = l == VarType.Double ? r : l;
            return other == VarType.Single ? VarType.Single : VarType.Double;
        }

        if (l == VarType.Single || r == VarType.Single)
        {
            var other = l == VarType.Single ? r : l;
            return other is VarType.Long or VarType.LongLong ? VarType.Double : VarType.Single;
        }

        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        if (l == VarType.Long || r == VarType.Long)
        {
            return VarType.Long;
        }

        if (l == VarType.Byte && r is VarType.Byte or VarType.String or VarType.Empty || r == VarType.Byte && l is VarType.String or VarType.Empty)
        {
            return VarType.Byte;
        }

        return VarType.Integer;
    }

    private static Variant Logical(in Variant left, in Variant right, Func<long, long, long> operation, LogicalKind kind)
    {
        if (left.IsObject || right.IsObject)
        {
            return Logical(Value(left), Value(right), operation, kind);
        }

        CheckOperands(left, right);
        var l = left.Type;
        var r = right.Type;
        if (l == VarType.Null || r == VarType.Null)
        {
            if (l == VarType.Null && r == VarType.Null)
            {
                return Variant.Null;
            }

            // The Null rules (MS-VBAL 5.6.9.7): the other operand decides when its value settles the result.
            var other = l == VarType.Null ? right : left;
            var otherIsLeft = r == VarType.Null;
            var effectiveWithNull = EffectiveLogicalType(other.Type, VarType.Null);
            var value = Coerce.ToInt64(other);
            switch (kind)
            {
                case LogicalKind.And when value == 0:
                    return Narrow(0, effectiveWithNull);
                case LogicalKind.Or when value != 0:
                    return other.Type == VarType.Boolean ? Variant.True : Narrow(value, effectiveWithNull);
                case LogicalKind.Imp when otherIsLeft && value == 0:
                    return Variant.True;
                case LogicalKind.Imp when !otherIsLeft && value != 0:
                    return other.Type == VarType.Boolean ? Variant.True : Narrow(value, effectiveWithNull);
                default:
                    return Variant.Null;
            }
        }

        var effective = EffectiveLogicalType(l, r);
        var result = operation(Coerce.ToInt64(left), Coerce.ToInt64(right));
        return Narrow(result, effective);
    }

    /// <summary>The binary logical effective value type (MS-VBAL 5.6.9.7 table).</summary>
    private static VarType EffectiveLogicalType(VarType l, VarType r)
    {
        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        static bool IsLongClass(VarType t) => t.IsFloatingPoint() || t.IsFixedPoint() || t is VarType.Long or VarType.String or VarType.Date;
        if (IsLongClass(l) || IsLongClass(r))
        {
            return VarType.Long;
        }

        if (l == VarType.Byte && r is VarType.Byte or VarType.Null || r == VarType.Byte && l is VarType.Byte or VarType.Null)
        {
            return VarType.Byte;
        }

        if (l == VarType.Boolean && r is VarType.Boolean or VarType.Null || r == VarType.Boolean && l is VarType.Boolean or VarType.Null)
        {
            return VarType.Boolean;
        }

        return VarType.Integer;
    }

    private static Variant Narrow(long value, VarType effective) => effective switch
    {
        VarType.Byte => Variant.FromByte((byte)value),
        VarType.Boolean => Variant.FromBoolean(value != 0),
        VarType.Integer => Variant.FromInt16((short)value),
        VarType.LongLong => Variant.FromInt64(value),
        _ => Variant.FromInt32((int)value),
    };

    private static void CheckOperands(in Variant left, in Variant right, bool allowError = false)
    {
        CheckOperand(left, allowError);
        CheckOperand(right, allowError);
    }

    private static void CheckOperand(in Variant operand, bool allowError = false)
    {
        switch (operand.Type)
        {
            case VarType.Array:
                throw VbaErrors.TypeMismatch();
            case VarType.Error when !allowError:
                throw VbaErrors.TypeMismatch();
            case VarType.Object:
                throw Coerce.NotConvertible(operand);
        }
    }
}

/// <summary>String comparison under an Option Compare mode (MS-VBAL 5.6.9.5, String rows).</summary>
public static class TextCompare
{
    /// <summary>Binary compares UTF-16 code units; Text compares case-insensitively under the regional collation.</summary>
    public static int Compare(string left, string right, CompareMode mode)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        return Compare(left.AsSpan(), right.AsSpan(), mode);
    }

    public static int Compare(ReadOnlySpan<char> left, ReadOnlySpan<char> right, CompareMode mode) => mode == CompareMode.Binary
        ? Math.Sign(left.SequenceCompareTo(right))
        : Math.Sign(CultureInfo.CurrentCulture.CompareInfo.Compare(left, right, CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreWidth));

    /// <summary>Two BSTRs: the code units decide, and under binary comparison the one with more bytes is greater when they agree, so a trailing odd byte counts (Strings golden: StrComp(LeftB("hello", 3), "h") is 1); text comparison stays the collation's, where "ß" and "ss" are equal.</summary>
    public static int Compare(VbaString left, VbaString right, CompareMode mode)
    {
        var order = Compare(left.Chars, right.Chars, mode);
        return order != 0 || mode == CompareMode.Text ? order : left.ByteLength.CompareTo(right.ByteLength);
    }
}

/// <summary>Pattern matching for the Like operator (MS-VBAL 5.6.9.6).</summary>
public static class LikePattern
{
    public static bool IsMatch(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern, CompareMode mode)
    {
        return Match(text, 0, pattern, 0, mode);
    }

    private static bool Match(ReadOnlySpan<char> text, int t, ReadOnlySpan<char> pattern, int p, CompareMode mode)
    {
        while (p < pattern.Length)
        {
            var c = pattern[p];
            switch (c)
            {
                case '*':
                    // The rest of the pattern must itself be well formed before anything is matched.
                    Validate(pattern, p + 1);
                    for (var skip = t; skip <= text.Length; skip++)
                    {
                        if (Match(text, skip, pattern, p + 1, mode))
                        {
                            return true;
                        }
                    }

                    return false;
                case '?':
                    if (t >= text.Length)
                    {
                        return false;
                    }

                    t++;
                    p++;
                    break;
                case '#':
                    if (t >= text.Length || !char.IsAsciiDigit(text[t]))
                    {
                        return false;
                    }

                    t++;
                    p++;
                    break;
                case '[':
                    {
                        var end = ParseCharList(pattern, p, out var negate, out var start);
                        if (t >= text.Length)
                        {
                            return false;
                        }

                        var matched = InCharList(text[t], pattern, start, end, mode);
                        if (matched == negate)
                        {
                            return false;
                        }

                        t++;
                        p = end + 1;
                        break;
                    }

                default:
                    if (t >= text.Length || !CharEquals(text[t], c, mode))
                    {
                        return false;
                    }

                    t++;
                    p++;
                    break;
            }
        }

        return t == text.Length;
    }

    private static void Validate(ReadOnlySpan<char> pattern, int p)
    {
        while (p < pattern.Length)
        {
            if (pattern[p] == '[')
            {
                p = ParseCharList(pattern, p, out _, out _) + 1;
            }
            else
            {
                p++;
            }
        }
    }

    /// <summary>Returns the index of the closing bracket; the list starts at <paramref name="start"/>.</summary>
    private static int ParseCharList(ReadOnlySpan<char> pattern, int open, out bool negate, out int start)
    {
        var p = open + 1;
        negate = p < pattern.Length && pattern[p] == '!';
        if (negate)
        {
            p++;
        }

        start = p;
        // A leading "-" or "]" is literal in the list; then scan to the closing bracket.
        if (p < pattern.Length && pattern[p] == '-')
        {
            p++;
        }

        while (p < pattern.Length && pattern[p] != ']')
        {
            p++;
        }

        if (p >= pattern.Length)
        {
            throw VbaErrors.InvalidPatternString();
        }

        // Ranges must ascend.
        for (var i = start; i < p; i++)
        {
            if (pattern[i] == '-' && i > start && i + 1 < p && Compare(pattern[i - 1], pattern[i + 1], CompareMode.Binary) > 0)
            {
                throw VbaErrors.InvalidPatternString();
            }
        }

        return p;
    }

    private static bool InCharList(char c, ReadOnlySpan<char> pattern, int start, int end, CompareMode mode)
    {
        var i = start;
        while (i < end)
        {
            if (i + 2 < end && pattern[i + 1] == '-')
            {
                if (Compare(pattern[i], c, mode) <= 0 && Compare(c, pattern[i + 2], mode) <= 0)
                {
                    return true;
                }

                i += 3;
            }
            else
            {
                if (CharEquals(c, pattern[i], mode))
                {
                    return true;
                }

                i++;
            }
        }

        return false;
    }

    private static bool CharEquals(char a, char b, CompareMode mode) => Compare(a, b, mode) == 0;

    private static int Compare(char a, char b, CompareMode mode) =>
        mode == CompareMode.Binary ? a.CompareTo(b) : TextCompare.Compare(a.ToString(), b.ToString(), CompareMode.Text);
}
