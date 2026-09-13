using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>
/// A VBA Variant: a value type tag plus the value (MS-VBAL 2.1, ARCHITECTURE.md section 5).
/// Scalars live inline, so arithmetic on numbers never allocates (CLAUDE.md R20). A String is a
/// BSTR pointer (D20): copying the struct aliases it, the storage that holds it owns it through
/// <see cref="ObjectRefs.Assign(ref Variant, in Variant)"/> and <see cref="ObjectRefs.Release(ref Variant)"/>,
/// and a string this struct creates is a temporary of the current statement. An object is its
/// IDispatch pointer, as the VARIANT holds it, owned the same way through IUnknown's AddRef and
/// Release. A Decimal lies inline in the DECIMAL layout. An array is its SAFEARRAY descriptor,
/// owned the way a String is (M7 C4), and a record never travels in a Variant (M7 C3). The
/// default value is Empty.
/// </summary>
/// <remarks>
/// The layout is the VARIANT's (ARCHITECTURE.md section 5): the type tag at 0, the value at 8, and
/// for a Decimal the whole DECIMAL from offset 0, its reserved word being the tag, its scale at 2,
/// its sign at 3, its high 32 bits at 4, and its low 64 bits at 8. An array's descriptor is the
/// pointer at 8, and nothing in the struct is managed, so it is the VARIANT (M7 C4).
/// </remarks>
[DebuggerDisplay("{DebugView,nq}")]
[DebuggerTypeProxy(typeof(VariantDebugView))]
[StructLayout(LayoutKind.Explicit, Size = 24)]
public readonly struct Variant : IEquatable<Variant>
{
    [FieldOffset(0)]
    private readonly VarType type;

    /// <summary>A Decimal's scale in the low byte (offset 2) and its sign in the high byte (offset 3, 0x80 when negative).</summary>
    [FieldOffset(2)]
    private readonly ushort scaleAndSign;

    /// <summary>A Decimal's high 32 bits.</summary>
    [FieldOffset(4)]
    private readonly uint high;

    [FieldOffset(8)]
    private readonly long bits;

    private Variant(VarType type, long bits)
    {
        this.type = type;
        this.bits = bits;
    }

    private Variant(ushort scaleAndSign, uint high, long low)
    {
        type = VarType.Decimal;
        this.scaleAndSign = scaleAndSign;
        this.high = high;
        bits = low;
    }

    public static readonly Variant Empty;
    public static readonly Variant Null = new(VarType.Null, 0);
    public static readonly Variant Nothing = new(VarType.Object, 0);
    public static readonly Variant Missing = new(VarType.Error, ErrorValue.Missing.Scode);
    public static readonly Variant True = new(VarType.Boolean, -1);
    public static readonly Variant False = new(VarType.Boolean, 0);
    /// <summary>A null BSTR, which reads as "" and is what <c>vbNullString</c> and an unassigned String are (Memory golden); it owns nothing.</summary>
    public static readonly Variant EmptyString = new(VarType.String, 0);

    /// <summary>The value type. Arrays report <see cref="VarType.Array"/>; see <see cref="VarTypeValue"/> for what <c>VarType()</c> returns.</summary>
    public VarType Type => type < VarType.Array ? type : VarType.Array;

    /// <summary>What the <c>VarType</c> function returns: the element type is added for arrays (8204 for a Variant array).</summary>
    public VarType VarTypeValue => type;

    public bool IsEmpty => type == VarType.Empty;

    public bool IsNull => type == VarType.Null;

    public bool IsMissing => type == VarType.Error && bits == ErrorValue.Missing.Scode;

    public bool IsArray => type >= VarType.Array;

    public bool IsObject => type == VarType.Object;

    public bool IsNothing => type == VarType.Object && bits == 0;

    public bool IsString => type == VarType.String;

    public bool IsError => type == VarType.Error;

    public bool IsNumeric => type.IsNumeric();

    public bool IsBoolean => type == VarType.Boolean;

    public bool IsDate => type == VarType.Date;

    public static Variant FromInt16(short value) => new(VarType.Integer, value);

    public static Variant FromInt32(int value) => new(VarType.Long, value);

    public static Variant FromInt64(long value) => new(VarType.LongLong, value);

    public static Variant FromSingle(float value) => new(VarType.Single, BitConverter.SingleToInt32Bits(value));

    public static Variant FromDouble(double value) => new(VarType.Double, BitConverter.DoubleToInt64Bits(value));

    public static Variant FromCurrency(Currency value) => new(VarType.Currency, value.Scaled);

    public static Variant FromDate(VbaDate value) => new(VarType.Date, BitConverter.DoubleToInt64Bits(value.Serial));

    /// <summary>A new BSTR with the text, owned by the current statement; null is the null BSTR. A store copies it (<see cref="ObjectRefs.Assign(ref Variant, in Variant)"/>).</summary>
    public static Variant FromString(string? value) => value is null ? EmptyString : FromVbaString(VbaString.Alloc(value.AsSpan()));

    /// <summary>A Variant over a BSTR the caller owns; the statement takes ownership and frees it when it ends.</summary>
    public static Variant FromVbaString(VbaString value)
    {
        if (value.IsNull)
        {
            return EmptyString;
        }

        ObjectRefs.OwnedString(value.Pointer);
        return new Variant(VarType.String, value.Pointer);
    }

    /// <summary>A Variant over a BSTR its storage owns: no temporary is registered. For containers that copied the string themselves.</summary>
    internal static Variant FromOwnedString(nint bstr) => new(VarType.String, bstr);

    /// <summary>
    /// A view of a String variable's BSTR for an expression (D20): nothing is allocated or registered,
    /// the variable keeps the string, and a store into another slot copies it. Generated code uses
    /// it wherever a typed String reaches an operator or a Variant parameter.
    /// </summary>
    public static Variant ViewString(VbaString value) => new(VarType.String, value.Pointer);

    public static Variant FromBoolean(bool value) => value ? True : False;

    /// <summary>A Boolean member's two bytes as they are (ROADMAP.md M7 C): a value other than True's -1 or False's 0, written through memory, stays that value (Memory golden).</summary>
    public static Variant FromBoolean(VbaBoolean value) => new(VarType.Boolean, value.Bits);

    public static Variant FromByte(byte value) => new(VarType.Byte, value);

    /// <summary>A Decimal in the DECIMAL layout, inline: nothing is allocated (CLAUDE.md R20).</summary>
    public static Variant FromDecimal(decimal value)
    {
        Span<int> parts = stackalloc int[4];
        _ = decimal.GetBits(value, parts);
        var scale = (ushort)((parts[3] >> 16) & 0xFF);
        var sign = parts[3] < 0 ? (ushort)0x8000 : (ushort)0;
        return new Variant((ushort)(scale | sign), (uint)parts[2], (long)(((ulong)(uint)parts[1] << 32) | (uint)parts[0]));
    }

    public static Variant FromError(ErrorValue value) => new(VarType.Error, value.Scode);

    /// <summary>A view of an array its storage owns (D20): the tag VT_ARRAY with the element type, the payload the descriptor; a store copies the array (<see cref="ObjectRefs.Own(in Variant)"/>).</summary>
    public static Variant FromArray(VbaArray value) => new(VarType.Array | value.ElementType, value.Descriptor);

    /// <summary>
    /// An object reference as the VARIANT holds one, its IDispatch pointer (D20; ARCHITECTURE.md
    /// section 6): a view of the reference the object's holder keeps, which a store AddRefs
    /// (<see cref="ObjectRefs.Own(in Variant)"/>). A runtime object nothing references yet (one
    /// fresh from a call) becomes a temporary of the current statement first, which gives it its
    /// pointer. null is Nothing.
    /// </summary>
    public static Variant FromObject(object? value) => value switch
    {
        null => Nothing,
        Variant variant => variant,
        VbaArray array => FromArray(array),
        string text => FromString(text),
        RuntimeObject runtime => new Variant(VarType.Object, runtime.InterfacePointer != 0 ? runtime.InterfacePointer : ObjectRefs.Owned(runtime).InterfacePointer),
        IComInterface com => new Variant(VarType.Object, com.InterfacePointer),
        _ => throw new InvalidOperationException(VbaErrors.Invariant($"A {value.GetType().Name} has no COM identity, so a Variant cannot hold it.")),
    };

    /// <summary>A view of an IDispatch pointer its storage owns (a VARIANT from COM, a SAFEARRAY element): no reference is taken or registered.</summary>
    public static Variant ViewInterface(nint dispatch) => new(VarType.Object, dispatch);

    public static implicit operator Variant(short value) => FromInt16(value);

    public static implicit operator Variant(int value) => FromInt32(value);

    public static implicit operator Variant(long value) => FromInt64(value);

    public static implicit operator Variant(float value) => FromSingle(value);

    public static implicit operator Variant(double value) => FromDouble(value);

    public static implicit operator Variant(bool value) => FromBoolean(value);

    public static implicit operator Variant(byte value) => FromByte(value);

    public static implicit operator Variant(string? value) => FromString(value);

    public static implicit operator Variant(Currency value) => FromCurrency(value);

    public static implicit operator Variant(VbaDate value) => FromDate(value);

    public static implicit operator Variant(decimal value) => FromDecimal(value);

    public static implicit operator Variant(ErrorValue value) => FromError(value);

    /// <summary>The Integer value; the Variant must hold one (use <see cref="Coerce"/> to convert).</summary>
    public short AsInt16() => Expect(VarType.Integer) ? (short)bits : throw Mismatch(VarType.Integer);

    public int AsInt32() => Expect(VarType.Long) ? (int)bits : throw Mismatch(VarType.Long);

    public long AsInt64() => Expect(VarType.LongLong) ? bits : throw Mismatch(VarType.LongLong);

    public float AsSingle() => Expect(VarType.Single) ? BitConverter.Int32BitsToSingle((int)bits) : throw Mismatch(VarType.Single);

    public double AsDouble() => Expect(VarType.Double) ? BitConverter.Int64BitsToDouble(bits) : throw Mismatch(VarType.Double);

    public Currency AsCurrency() => Expect(VarType.Currency) ? Currency.FromScaled(bits) : throw Mismatch(VarType.Currency);

    public VbaDate AsDate() => Expect(VarType.Date) ? VbaDate.FromSerial(BitConverter.Int64BitsToDouble(bits)) : throw Mismatch(VarType.Date);

    /// <summary>The whole code units as a .NET string, a copy made on every call; the library moves to <see cref="AsVbaString"/> and spans (ROADMAP.md WP2 B).</summary>
    public string AsString() => AsVbaString().ToString();

    /// <summary>The BSTR the Variant holds; the Variant must hold a String. The struct aliases it and must not free it.</summary>
    public VbaString AsVbaString() => Expect(VarType.String) ? VbaString.Adopt((nint)bits) : throw Mismatch(VarType.String);

    /// <summary>The BSTR pointer of a String, 0 otherwise; what the ownership helpers copy and free.</summary>
    internal nint StringPointer => type == VarType.String ? (nint)bits : 0;

    /// <summary>The IDispatch pointer of an object, 0 for Nothing and for anything else; what the ownership helpers AddRef and Release, and what COM receives.</summary>
    public nint InterfacePointer => type == VarType.Object ? (nint)bits : 0;

    public bool AsBoolean() => Expect(VarType.Boolean) ? bits != 0 : throw Mismatch(VarType.Boolean);

    /// <summary>A Boolean's two bytes as the number VBA converts it to: -1 for True, 0 for False, and whatever else memory put there (Memory golden).</summary>
    internal short BooleanBits => Expect(VarType.Boolean) ? (short)bits : throw Mismatch(VarType.Boolean);

    public byte AsByte() => Expect(VarType.Byte) ? (byte)bits : throw Mismatch(VarType.Byte);

    public decimal AsDecimal() => Expect(VarType.Decimal)
        ? new decimal((int)(uint)bits, (int)(uint)((ulong)bits >> 32), (int)high, (scaleAndSign & 0x8000) != 0, (byte)(scaleAndSign & 0xFF))
        : throw Mismatch(VarType.Decimal);

    public ErrorValue AsError() => Expect(VarType.Error) ? new ErrorValue((int)bits) : throw Mismatch(VarType.Error);

    public VbaArray AsArray() => IsArray ? VbaArray.View((nint)bits, type & ~VarType.Array) : throw Mismatch(VarType.Array);

    /// <summary>The SAFEARRAY of an allocated array, 0 otherwise; what the ownership helpers copy and destroy.</summary>
    internal nint ArrayDescriptor => IsArray ? (nint)bits : 0;

    /// <summary>
    /// The object behind the pointer, or null for Nothing: a runtime object is itself (a class
    /// instance, a Collection), and a COM object is a view the Interop provider wraps around the
    /// pointer, which takes a reference of its own only when something stores it.
    /// </summary>
    public object? AsObject() => Expect(VarType.Object) ? ObjectOf((nint)bits) : throw Mismatch(VarType.Object);

    /// <summary>The raw 64 bits of an inline scalar: two's complement for integers, IEEE bits for Single, Double, and Date, the scaled Currency.</summary>
    public long RawBits => bits;

    public bool Equals(Variant other)
    {
        if (type != other.type)
        {
            return false;
        }

        return type switch
        {
            VarType.String => Bstr.Bytes((nint)bits).SequenceEqual(Bstr.Bytes((nint)other.bits)),
            VarType.Decimal => AsDecimal() == other.AsDecimal(),
            _ => bits == other.bits,
        };
    }

    public override bool Equals(object? obj) => obj is Variant other && Equals(other);

    public override int GetHashCode() => type switch
    {
        VarType.String => HashCode.Combine(type, AsVbaString().GetHashCode()),
        VarType.Decimal => HashCode.Combine(type, AsDecimal()),
        _ => HashCode.Combine(type, bits),
    };

    public static bool operator ==(Variant left, Variant right) => left.Equals(right);

    public static bool operator !=(Variant left, Variant right) => !left.Equals(right);

    /// <summary>A debugging rendering, not VBA's string conversion (see <see cref="Coerce.ToString(in Variant)"/>).</summary>
    public override string ToString() => DebugView;

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private string DebugView => Type switch
    {
        VarType.Empty => "Empty",
        VarType.Null => "Null",
        VarType.Integer => VbaErrors.Invariant($"{AsInt16()} (Integer)"),
        VarType.Long => VbaErrors.Invariant($"{AsInt32()} (Long)"),
        VarType.LongLong => VbaErrors.Invariant($"{AsInt64()} (LongLong)"),
        VarType.Single => VbaErrors.Invariant($"{AsSingle().ToString("R", CultureInfo.InvariantCulture)} (Single)"),
        VarType.Double => VbaErrors.Invariant($"{AsDouble().ToString("R", CultureInfo.InvariantCulture)} (Double)"),
        VarType.Currency => VbaErrors.Invariant($"{AsCurrency()} (Currency)"),
        VarType.Date => AsDate().DebugView + " (Date)",
        VarType.String => VbaErrors.Invariant($"\"{AsString()}\""),
        VarType.Boolean => AsBoolean() ? "True" : "False",
        VarType.Byte => VbaErrors.Invariant($"{AsByte()} (Byte)"),
        VarType.Decimal => VbaErrors.Invariant($"{AsDecimal()} (Decimal)"),
        VarType.Error => IsMissing ? "Missing" : AsError().ToString(),
        VarType.Array => VbaErrors.Invariant($"{AsArray().ElementType.TypeName()}({AsArray().BoundsText})"),
        VarType.Object => bits == 0 ? "Nothing" : VbaErrors.Invariant($"Object 0x{bits:X}"),
        _ => type.ToString(),
    };

    internal static object? ObjectOf(nint pointer) =>
        pointer == 0 ? null : (object?)RuntimeObject.FromPointer(pointer) ?? Com.Provider.Wrap(pointer);

    private bool Expect(VarType expected) => type == expected;

    private InvalidOperationException Mismatch(VarType expected) =>
        new(VbaErrors.Invariant($"The Variant holds a {type} value, not {expected}."));
}

/// <summary>What the locals window shows under a Variant (ARCHITECTURE.md section 9): an array's elements under their subscripts, or the object it holds; nothing for a scalar, whose value is the Variant's own line.</summary>
internal sealed class VariantDebugView(Variant value)
{
    [DebuggerBrowsable(DebuggerBrowsableState.RootHidden)]
    public object?[] Contents =>
        value.IsArray ? new VbaArrayDebugView(value.AsArray()).Elements
        : value.IsObject && value.AsObject() is { } target ? [target]
        : [];
}
