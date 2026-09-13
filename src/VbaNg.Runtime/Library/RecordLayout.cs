using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace VbaNg.Runtime.Library;

/// <summary>
/// How Put lays values out and Get reads them back (MS-VBAL 5.4.5.9): numbers in their binary
/// size, Boolean as two bytes, Currency scaled, Date as its serial, strings as ANSI bytes with a
/// two-byte length in Random mode and inside records, fixed-length strings padded with spaces,
/// Variants with a two-byte type tag, records field by field, arrays element by element.
/// </summary>
internal static class RecordLayout
{
    public static byte[] Encode(in Variant value, VarType type, int fixedLength, bool lengthPrefixed)
    {
        switch (type)
        {
            case VarType.Integer:
                return BitConverter.GetBytes(Coerce.ToInt16(value));
            case VarType.Long:
                return BitConverter.GetBytes(Coerce.ToInt32(value));
            case VarType.LongLong:
                return BitConverter.GetBytes(Coerce.ToInt64(value));
            case VarType.Single:
                return BitConverter.GetBytes(Coerce.ToSingle(value));
            case VarType.Double:
                return BitConverter.GetBytes(Coerce.ToDouble(value));
            case VarType.Currency:
                return BitConverter.GetBytes(Coerce.ToCurrency(value).Scaled);
            case VarType.Date:
                return BitConverter.GetBytes(Coerce.ToDate(value).Serial);
            case VarType.Boolean:
                return BitConverter.GetBytes(Coerce.ToBoolean(value) ? (short)-1 : (short)0);
            case VarType.Byte:
                return [Coerce.ToByte(value)];
            case VarType.String:
                {
                    var text = Coerce.ToString(value);
                    if (fixedLength > 0)
                    {
                        return FileSystem.Ansi.GetBytes(text.Length >= fixedLength ? text[..fixedLength] : text.PadRight(fixedLength));
                    }

                    var bytes = FileSystem.Ansi.GetBytes(text);
                    return lengthPrefixed ? [.. BitConverter.GetBytes((short)bytes.Length), .. bytes] : bytes;
                }

            case VarType.Variant:
                {
                    if (value.IsObject || value.IsArray)
                    {
                        throw VbaErrors.TypeMismatch();
                    }

                    var kind = value.Type == VarType.Empty ? VarType.Empty : value.Type;
                    var tag = BitConverter.GetBytes((short)kind);
                    if (kind is VarType.Empty or VarType.Null)
                    {
                        return tag;
                    }

                    if (kind == VarType.Decimal)
                    {
                        throw VbaErrors.TypeMismatch();
                    }

                    // A string inside a Variant always carries its length.
                    return [.. tag, .. Encode(value, kind, 0, kind == VarType.String || lengthPrefixed)];
                }

            default:
                throw VbaErrors.TypeMismatch();
        }
    }

    public static Variant Decode(FileChannel channel, VarType type, in Variant current, int fixedLength, bool lengthPrefixed)
    {
        switch (type)
        {
            case VarType.Integer:
                return Variant.FromInt16(BitConverter.ToInt16(channel.ReadBytes(2)));
            case VarType.Long:
                return Variant.FromInt32(BitConverter.ToInt32(channel.ReadBytes(4)));
            case VarType.LongLong:
                return Variant.FromInt64(BitConverter.ToInt64(channel.ReadBytes(8)));
            case VarType.Single:
                return Variant.FromSingle(BitConverter.ToSingle(channel.ReadBytes(4)));
            case VarType.Double:
                return Variant.FromDouble(BitConverter.ToDouble(channel.ReadBytes(8)));
            case VarType.Currency:
                return Variant.FromCurrency(Currency.FromScaled(BitConverter.ToInt64(channel.ReadBytes(8))));
            case VarType.Date:
                return Variant.FromDate(VbaDate.FromSerial(BitConverter.ToDouble(channel.ReadBytes(8))));
            case VarType.Boolean:
                return Variant.FromBoolean(BitConverter.ToInt16(channel.ReadBytes(2)) != 0);
            case VarType.Byte:
                return Variant.FromByte(channel.ReadBytes(1)[0]);
            case VarType.String:
                {
                    int length;
                    if (fixedLength > 0)
                    {
                        length = fixedLength;
                    }
                    else if (lengthPrefixed)
                    {
                        length = BitConverter.ToInt16(channel.ReadBytes(2));
                    }
                    else
                    {
                        length = current.IsString ? FileSystem.Ansi.GetByteCount(current.AsString()) : 0;
                    }

                    return Variant.FromString(FileSystem.Ansi.GetString(channel.ReadBytes(Math.Max(0, length))));
                }

            case VarType.Variant:
                {
                    var kind = (VarType)BitConverter.ToInt16(channel.ReadBytes(2));
                    if (kind is VarType.Empty)
                    {
                        return Variant.Empty;
                    }

                    if (kind == VarType.Null)
                    {
                        return Variant.Null;
                    }

                    return Decode(channel, kind, Variant.Empty, 0, kind == VarType.String || lengthPrefixed);
                }

            default:
                throw VbaErrors.TypeMismatch();
        }
    }

    public static byte[] EncodeRecord(object record)
    {
        var bytes = new List<byte>();
        foreach (var (field, type, fixedLength) in Fields(record))
        {
            using var embedded = EmbeddedMember.Open(field, record);
            var value = embedded is null ? FieldValue(field, record, type) : embedded.Array;
            switch (value)
            {
                case IVbaRecord nested:
                    bytes.AddRange(EncodeRecord(nested));
                    break;
                case VbaArray array:
                    foreach (var index in Indices(array))
                    {
                        bytes.AddRange(Encode(array.Get(index), array.ElementType, 0, false));
                    }

                    break;
                default:
                    bytes.AddRange(Encode(FromField(value), type, fixedLength, lengthPrefixed: true));
                    break;
            }
        }

        return bytes.ToArray();
    }

    public static void Read(FileChannel channel, object record)
    {
        foreach (var (field, type, fixedLength) in Fields(record))
        {
            using var embedded = EmbeddedMember.Open(field, record);
            var value = embedded is null ? FieldValue(field, record, type) : embedded.Array;
            switch (value)
            {
                case IVbaRecord nested:
                    try
                    {
                        Read(channel, nested);
                    }
                    finally
                    {
                        // The nested record is a boxed copy: what it reads, and what it freed, goes back into the outer record.
                        field.SetValue(record, nested);
                    }

                    break;
                case VbaArray array:
                    foreach (var index in Indices(array))
                    {
                        array.Set(index, Decode(channel, array.ElementType, array.Get(index), 0, false));
                    }

                    break;
                default:
                    SetField(field, record, Decode(channel, type, FromField(value), fixedLength, lengthPrefixed: true));
                    break;
            }
        }
    }

    /// <summary>Len of a record: the fields' sizes without padding, a fixed string one byte a character, an array of them one byte a character of each element (Types golden: 8 for Byte, String * 3 (0 To 1), Byte) and a variable-length string its eight-byte pointer whatever it holds (Types golden: 35 for Byte, Integer, Long, Double, Currency, String, String * 4).</summary>
    public static int Size(object record)
    {
        var total = 0;
        foreach (var (field, type, fixedLength) in Fields(record))
        {
            using var embedded = EmbeddedMember.Open(field, record);
            var value = embedded is null ? FieldValue(field, record, type) : embedded.Array;
            total += value switch
            {
                IVbaRecord nested => Size(nested),
                VbaArray array => array.Count * ScalarSize(array.ElementType, fixedLength, Variant.Empty),
                VbaString => 8,
                _ => ScalarSize(type, fixedLength, FromField(value)),
            };
        }

        return total;
    }

    /// <summary>The fields of a record with the size and alignment each takes in memory: the layout LenB measures and LSet copies (RecordImage). A fixed-size array lies inline, a dynamic one is its pointer.</summary>
    public static IEnumerable<(FieldInfo Field, VarType Type, int Size, int Alignment, int FixedLength)> MemoryFields(object record)
    {
        foreach (var (field, type, fixedLength) in Fields(record))
        {
            var (size, alignment) = LayoutOf(field, record, type, fixedLength);
            yield return (field, type, size, alignment, fixedLength);
        }
    }

    private static (int Size, int Alignment) LayoutOf(FieldInfo field, object record, VarType type, int fixedLength)
    {
        using var embedded = EmbeddedMember.Open(field, record);
        return (embedded is null ? FieldValue(field, record, type) : embedded.Array) switch
        {
            IVbaRecord nested => (MemorySize(nested), MemoryAlignment(nested)),
            VbaArray { IsFixedSize: true } array => (array.Count * ElementSize(array, fixedLength), ElementAlignment(array, fixedLength)),
            VbaArray => (8, 8),
            _ => MemoryLayout(type, fixedLength),
        };
    }

    /// <summary>The bytes one element of an array member takes in memory.</summary>
    public static int ElementSize(VbaArray array, int fixedLength) =>
        array.ElementType == VarType.UserDefinedType && array.Count > 0 ? MemorySize(array.BoxRecord(0)) : MemoryLayout(array.ElementType, fixedLength).Size;

    private static int ElementAlignment(VbaArray array, int fixedLength) =>
        array.ElementType == VarType.UserDefinedType && array.Count > 0 ? MemoryAlignment(array.BoxRecord(0)) : MemoryLayout(array.ElementType, fixedLength).Alignment;

    /// <summary>LenB of a record: every field at its natural alignment, fixed strings two bytes a character at one-byte alignment (Memory golden), and the total padded to the largest alignment.</summary>
    public static int MemorySize(object record)
    {
        var total = 0;
        var largest = 1;
        foreach (var (_, _, size, alignment, _) in MemoryFields(record))
        {
            total = Align(total, alignment) + size;
            largest = Math.Max(largest, alignment);
        }

        return Align(total, largest);
    }

    private static int MemoryAlignment(object record)
    {
        var largest = 1;
        foreach (var (_, _, _, alignment, _) in MemoryFields(record))
        {
            largest = Math.Max(largest, alignment);
        }

        return largest;
    }

    private static (int Size, int Alignment) MemoryLayout(VarType type, int fixedLength) => type switch
    {
        VarType.Integer or VarType.Boolean => (2, 2),
        VarType.Long or VarType.Single => (4, 4),
        VarType.LongLong or VarType.Double or VarType.Currency or VarType.Date => (8, 8),
        VarType.Byte => (1, 1),
        VarType.String => fixedLength > 0 ? (fixedLength * 2, 1) : (8, 8),
        VarType.Variant => (24, 8),
        _ => (8, 8),
    };

    private static int Align(int offset, int alignment) => (offset + alignment - 1) / alignment * alignment;

    private static int ScalarSize(VarType type, int fixedLength, in Variant value) => type switch
    {
        VarType.Integer or VarType.Boolean => 2,
        VarType.Long or VarType.Single => 4,
        VarType.LongLong or VarType.Double or VarType.Currency or VarType.Date => 8,
        VarType.Byte => 1,
        VarType.String => fixedLength > 0 ? fixedLength : 2 + (value.IsString ? FileSystem.Ansi.GetByteCount(value.AsString()) : 0),
        VarType.Variant => 16,
        _ => 0,
    };

    /// <summary>Every index tuple of an array in storage order (the first dimension varying fastest).</summary>
    public static IEnumerable<int[]> Indices(VbaArray array)
    {
        if (!array.IsAllocated)
        {
            yield break;
        }

        var rank = array.Rank;
        var counts = new int[rank];
        var total = 1;
        for (var dimension = 0; dimension < rank; dimension++)
        {
            counts[dimension] = array.UBound(dimension + 1) - array.LBound(dimension + 1) + 1;
            total *= counts[dimension];
        }

        for (var position = 0; position < total; position++)
        {
            var indices = new int[rank];
            var remainder = position;
            for (var dimension = 0; dimension < rank; dimension++)
            {
                indices[dimension] = array.LBound(dimension + 1) + (remainder % counts[dimension]);
                remainder /= counts[dimension];
            }

            yield return indices;
        }
    }

    /// <summary>The public fields of a generated record struct, boxed, in declaration order, with the VBA type and fixed string length each carries.</summary>
    private static IEnumerable<(FieldInfo Field, VarType Type, int FixedLength)> Fields(object record)
    {
        foreach (var field in record.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance).OrderBy(f => f.MetadataToken))
        {
            var attribute = field.GetCustomAttribute<VbaFieldAttribute>();
            var type = attribute?.Type ?? TypeOf(field.FieldType);
            yield return (field, type, attribute?.FixedLength ?? 0);
        }
    }

    private static VarType TypeOf(Type clrType)
    {
        if (clrType == typeof(short))
        {
            return VarType.Integer;
        }

        if (clrType == typeof(int))
        {
            return VarType.Long;
        }

        if (clrType == typeof(long))
        {
            return VarType.LongLong;
        }

        if (clrType == typeof(float))
        {
            return VarType.Single;
        }

        if (clrType == typeof(double))
        {
            return VarType.Double;
        }

        if (clrType == typeof(Currency))
        {
            return VarType.Currency;
        }

        if (clrType == typeof(VbaDate))
        {
            return VarType.Date;
        }

        if (clrType == typeof(VbaBoolean))
        {
            return VarType.Boolean;
        }

        if (clrType == typeof(bool))
        {
            return VarType.Boolean;
        }

        if (clrType == typeof(byte))
        {
            return VarType.Byte;
        }

        if (clrType == typeof(string) || clrType == typeof(VbaString) || clrType.IsDefined(typeof(InlineArrayAttribute), false))
        {
            return VarType.String;
        }

        return VarType.Variant;
    }

    /// <summary>A field's value for the file statements, Len, LenB, and LSet: a dynamic array member (ROADMAP.md M7 C8) as a view of its array, of the element type its attribute names.</summary>
    internal static object? FieldValue(FieldInfo field, object record, VarType type)
    {
        var value = field.GetValue(record);
        return value is VbaArrayMember member ? member.View(type) : value;
    }

    internal static Variant FromField(object? value) => value switch
    {
        null => Variant.Empty,
        Variant variant => variant,
        short s => Variant.FromInt16(s),
        int i => Variant.FromInt32(i),
        long l => Variant.FromInt64(l),
        float f => Variant.FromSingle(f),
        double d => Variant.FromDouble(d),
        Currency c => Variant.FromCurrency(c),
        VbaDate d => Variant.FromDate(d),
        bool b => Variant.FromBoolean(b),
        VbaBoolean b => Variant.FromBoolean(b),
        byte b => Variant.FromByte(b),
        string s => Variant.FromString(s),
        VbaString s => Variant.ViewString(s),
        { } chars when chars.GetType().IsDefined(typeof(InlineArrayAttribute), false) => Variant.FromString(ReadChars(chars)),
        _ => throw VbaErrors.TypeMismatch(),
    };

    /// <summary>A fixed-length string member, which lies inline in its record as its characters (ROADMAP.md M7 C5), read from its box.</summary>
    private static string ReadChars(object box)
    {
        var length = CharCount(box.GetType());
        var handle = GCHandle.Alloc(box, GCHandleType.Pinned);
        try
        {
            return Marshal.PtrToStringUni(handle.AddrOfPinnedObject(), length);
        }
        finally
        {
            handle.Free();
        }
    }

    /// <summary>The characters of a fixed-length string member's inline storage, which holds them as bytes, two a character, so it aligns at one byte as VBA's does (ROADMAP.md M7 C10).</summary>
    private static int CharCount(Type type) => type.GetCustomAttribute<InlineArrayAttribute>()!.Length / sizeof(char);

    /// <summary>A box of a fixed-length string member's characters holding the value, already padded to its length.</summary>
    private static object BoxChars(Type type, string value)
    {
        var box = RuntimeHelpers.GetUninitializedObject(type);
        var handle = GCHandle.Alloc(box, GCHandleType.Pinned);
        try
        {
            Marshal.Copy(value.ToCharArray(), 0, handle.AddrOfPinnedObject(), value.Length);
        }
        finally
        {
            handle.Free();
        }

        return box;
    }

    internal static void SetField(FieldInfo field, object record, in Variant value)
    {
        var clrType = field.FieldType;
        if (clrType == typeof(Variant))
        {
            field.SetValue(record, value);
            return;
        }

        var converted = Coerce.ToType(value, TypeOf(clrType));
        switch (converted.Type)
        {
            case VarType.Integer:
                field.SetValue(record, converted.AsInt16());
                break;
            case VarType.Long:
                field.SetValue(record, converted.AsInt32());
                break;
            case VarType.LongLong:
                field.SetValue(record, converted.AsInt64());
                break;
            case VarType.Single:
                field.SetValue(record, converted.AsSingle());
                break;
            case VarType.Double:
                field.SetValue(record, converted.AsDouble());
                break;
            case VarType.Currency:
                field.SetValue(record, converted.AsCurrency());
                break;
            case VarType.Date:
                field.SetValue(record, converted.AsDate());
                break;
            case VarType.Boolean:
                field.SetValue(record, clrType == typeof(VbaBoolean) ? VbaBoolean.FromBoolean(converted.AsBoolean()) : (object)converted.AsBoolean());
                break;
            case VarType.Byte:
                field.SetValue(record, converted.AsByte());
                break;
            case VarType.String when clrType == typeof(VbaString):
                {
                    // The field owns its BSTR (D20): the old one goes, a copy of the value stays.
                    var old = (VbaString)field.GetValue(record)!;
                    field.SetValue(record, ObjectRefs.Own(converted.AsVbaString()));
                    old.Free();
                    break;
                }

            case VarType.String when clrType.IsDefined(typeof(InlineArrayAttribute), false):
                field.SetValue(record, BoxChars(clrType, Strings.ToFixed(converted.AsString(), CharCount(clrType))));
                break;
            case VarType.String:
                field.SetValue(record, converted.AsString());
                break;
            default:
                field.SetValue(record, converted);
                break;
        }
    }
}

/// <summary>Marks a generated record field with the VBA type the file statements lay it out as, and the length of a fixed-length string.</summary>
[AttributeUsage(AttributeTargets.Field, AllowMultiple = false, Inherited = false)]
public sealed class VbaFieldAttribute(VarType type, int fixedLength = 0) : Attribute
{
    public VarType Type { get; } = type;

    public int FixedLength { get; } = fixedLength;
}

/// <summary>
/// A fixed-size array member of a boxed record, whose elements lie inline in it (ROADMAP.md M7
/// C6), opened as an array for the file statements, Len, LenB, and LSet: the member's box
/// pinned, a view over it, and on Dispose the box written back into the record, so what they
/// write through the array reaches the record. Null for any other member.
/// </summary>
internal sealed class EmbeddedMember : IDisposable
{
    private readonly FieldInfo field;
    private readonly object record;
    private readonly object box;
    private readonly char[]? chars;
    private readonly int length;
    private GCHandle handle;

    private EmbeddedMember(FieldInfo field, object record, object box, GCHandle handle, VbaArray array, char[]? chars = null, int length = 0)
    {
        this.field = field;
        this.record = record;
        this.box = box;
        this.handle = handle;
        this.chars = chars;
        this.length = length;
        Array = array;
    }

    public VbaArray Array { get; }

    public static EmbeddedMember? Open(FieldInfo field, object record)
    {
        var type = field.FieldType;
        if (!typeof(IEmbeddedArray).IsAssignableFrom(type))
        {
            return null;
        }

        var elementType = (VarType)type.GetProperty(nameof(IEmbeddedArray.ElementType), BindingFlags.Public | BindingFlags.Static)!.GetValue(null)!;
        var bounds = (int[])type.GetProperty(nameof(IEmbeddedArray.Bounds), BindingFlags.Public | BindingFlags.Static)!.GetValue(null)!;
        var box = field.GetValue(record)!;
        var handle = GCHandle.Alloc(box, GCHandleType.Pinned);
        var pairs = VbaArray.BoundPairs(bounds);
        if (type.GetField("First")!.FieldType.GetCustomAttribute<InlineArrayAttribute>() is { } inline)
        {
            // A fixed-length string array member (ROADMAP.md M7 C10) has no descriptor over its characters: it opens as an array of Strings of its own, whose elements go back on Dispose.
            var characters = inline.Length / sizeof(char);
            var all = new char[pairs.Aggregate(1, (n, b) => n * (b.Upper - b.Lower + 1)) * characters];
            Marshal.Copy(handle.AddrOfPinnedObject(), all, 0, all.Length);
            return new EmbeddedMember(field, record, box, handle, VbaArray.FromChars(all, pairs, characters), all, characters);
        }

        return new EmbeddedMember(field, record, box, handle, VbaArray.EmbeddedView(handle.AddrOfPinnedObject(), elementType, pairs, RecordKind.OfStorage(type)));
    }

    public void Dispose()
    {
        if (chars is not null)
        {
            Array.CopyCharsTo(chars, length);
            Marshal.Copy(chars, 0, handle.AddrOfPinnedObject(), chars.Length);
        }

        Array.Destroy();
        handle.Free();
        field.SetValue(record, box);
    }
}
