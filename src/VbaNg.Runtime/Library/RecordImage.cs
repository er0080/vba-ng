using System.Reflection;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The memory image of a record, as LSet between user-defined types copies it (MS-VBAL 5.4.3.6):
/// every field at its natural alignment, numbers in their binary size, fixed-length strings two
/// bytes a character, fixed-size arrays inline, and eight opaque bytes for anything held by
/// reference (a variable-length string, a dynamic array, an object). LSet writes the source's
/// image over the start of the destination's and leaves the rest of the destination alone: a
/// Long after the copied bytes keeps its value, and a fixed string that straddles the end takes
/// the bytes that reach it (Types golden). Until the memory model of M7 gives records a real
/// layout, this is the layout LenB reports, spelled out byte by byte; a fixed-size array member
/// lies inline in the record since M7 C6 and is read and written through a view of it.
/// </summary>
public static class RecordImage
{
    /// <summary>LSet destination = source between records: the destination boxed, its image overwritten, and the box written back whatever the read raises (ROADMAP.md M7 C3).</summary>
    public static void LSet<TDestination, TSource>(ref TDestination destination, in TSource source)
        where TDestination : struct, IVbaRecord
        where TSource : struct, IVbaRecord
    {
        object box = destination;
        var target = Write(box);
        var image = Write(source);
        Array.Copy(image, target, Math.Min(image.Length, target.Length));
        try
        {
            ReadInto(box, target, 0);
        }
        finally
        {
            destination = (TDestination)box;
        }
    }

    /// <summary>The copy of a record with a fixed-length string member that VBA hands a DLL (MS-VBAL 5.2.3.5; Declares golden): its memory image with each fixed-length string as its ANSI bytes, one a character.</summary>
    public static byte[] DeclareImage<T>(in T record)
        where T : struct, IVbaRecord => Write(record, ansi: true);

    /// <summary>After the call, what the DLL wrote into the copy goes back into the record, its fixed-length strings read as ANSI.</summary>
    public static void FromDeclareImage<T>(ref T record, byte[] image)
        where T : struct, IVbaRecord
    {
        ArgumentNullException.ThrowIfNull(image);
        object box = record;
        try
        {
            ReadInto(box, image, 0, ansi: true);
        }
        finally
        {
            record = (T)box;
        }
    }

    /// <summary>The record, boxed, laid out in memory.</summary>
    internal static byte[] Write(object record, bool ansi = false)
    {
        var buffer = new byte[ansi ? AnsiRecordLayout(record).Size : RecordLayout.MemorySize(record)];
        WriteInto(record, buffer, 0, ansi);
        return buffer;
    }

    private static void WriteInto(object record, byte[] buffer, int offset, bool ansi = false)
    {
        foreach (var slot in Slots(record, ansi))
        {
            using var embedded = EmbeddedMember.Open(slot.Field, record);
            var value = embedded is null ? RecordLayout.FieldValue(slot.Field, record, slot.Type) : embedded.Array;
            var at = offset + slot.Offset;
            switch (value)
            {
                case IVbaRecord nested:
                    WriteInto(nested, buffer, at, ansi);
                    break;
                case VbaArray { IsFixedSize: true } array:
                    {
                        var elementSize = RecordLayout.ElementSize(array, slot.FixedLength);
                        var i = 0;
                        foreach (var index in RecordLayout.Indices(array))
                        {
                            if (array.ElementType == VarType.UserDefinedType)
                            {
                                WriteInto(array.BoxRecord(i), buffer, at + i * elementSize, ansi);
                            }
                            else
                            {
                                WriteScalar(buffer.AsSpan(at + i * elementSize, elementSize), array.Get(index), array.ElementType, slot.FixedLength);
                            }

                            i++;
                        }

                        break;
                    }

                case VbaArray or Variant or null:
                    // A reference or a Variant: opaque here, the bytes stay zero.
                    break;
                case string when slot.FixedLength == 0:
                    break;
                default:
                    WriteScalar(buffer.AsSpan(at, slot.Size), RecordLayout.FromField(value), slot.Type, slot.FixedLength, ansi);
                    break;
            }
        }
    }

    private static void ReadInto(object record, byte[] buffer, int offset, bool ansi = false)
    {
        foreach (var slot in Slots(record, ansi))
        {
            using var embedded = EmbeddedMember.Open(slot.Field, record);
            var value = embedded is null ? RecordLayout.FieldValue(slot.Field, record, slot.Type) : embedded.Array;
            var at = offset + slot.Offset;
            switch (value)
            {
                case IVbaRecord nested:
                    try
                    {
                        ReadInto(nested, buffer, at, ansi);
                    }
                    finally
                    {
                        // The nested record is a boxed copy: it goes back into the outer record.
                        slot.Field.SetValue(record, nested);
                    }

                    break;
                case VbaArray { IsFixedSize: true } array:
                    {
                        var elementSize = RecordLayout.ElementSize(array, slot.FixedLength);
                        var i = 0;
                        foreach (var index in RecordLayout.Indices(array))
                        {
                            if (array.ElementType == VarType.UserDefinedType)
                            {
                                var element = array.BoxRecord(i);
                                try
                                {
                                    ReadInto(element, buffer, at + i * elementSize, ansi);
                                }
                                finally
                                {
                                    array.UnboxRecord(i, element);
                                }
                            }
                            else
                            {
                                array.Set(index, ReadScalar(buffer.AsSpan(at + i * elementSize, elementSize), array.ElementType, slot.FixedLength));
                            }

                            i++;
                        }

                        break;
                    }

                case VbaArray or Variant or null:
                    // A reference or a Variant keeps what the destination held.
                    break;
                case string when slot.FixedLength == 0:
                    break;
                default:
                    RecordLayout.SetField(slot.Field, record, ReadScalar(buffer.AsSpan(at, slot.Size), slot.Type, slot.FixedLength, ansi));
                    break;
            }
        }
    }

    private static void WriteScalar(Span<byte> span, in Variant value, VarType type, int fixedLength, bool ansi = false)
    {
        switch (type)
        {
            case VarType.Integer:
                BitConverter.TryWriteBytes(span, Coerce.ToInt16(value));
                break;
            case VarType.Boolean:
                BitConverter.TryWriteBytes(span, Coerce.ToBoolean(value) ? (short)-1 : (short)0);
                break;
            case VarType.Long:
                BitConverter.TryWriteBytes(span, Coerce.ToInt32(value));
                break;
            case VarType.Single:
                BitConverter.TryWriteBytes(span, Coerce.ToSingle(value));
                break;
            case VarType.LongLong:
                BitConverter.TryWriteBytes(span, Coerce.ToInt64(value));
                break;
            case VarType.Double:
                BitConverter.TryWriteBytes(span, Coerce.ToDouble(value));
                break;
            case VarType.Currency:
                BitConverter.TryWriteBytes(span, Coerce.ToCurrency(value).Scaled);
                break;
            case VarType.Date:
                BitConverter.TryWriteBytes(span, Coerce.ToDate(value).Serial);
                break;
            case VarType.Byte:
                span[0] = Coerce.ToByte(value);
                break;
            case VarType.String when fixedLength > 0:
                {
                    var padded = Strings.ToFixed(Coerce.ToString(value), fixedLength);
                    if (ansi)
                    {
                        // The copy a DLL receives holds the string's ANSI bytes, one a character (Declares golden).
                        var bytes = FileSystem.Ansi.GetBytes(padded);
                        bytes.AsSpan(0, Math.Min(bytes.Length, fixedLength)).CopyTo(span);
                        break;
                    }

                    for (var i = 0; i < fixedLength; i++)
                    {
                        BitConverter.TryWriteBytes(span.Slice(i * 2, 2), (ushort)padded[i]);
                    }

                    break;
                }

            default:
                break;
        }
    }

    private static Variant ReadScalar(ReadOnlySpan<byte> span, VarType type, int fixedLength, bool ansi = false)
    {
        switch (type)
        {
            case VarType.Integer:
                return Variant.FromInt16(BitConverter.ToInt16(span));
            case VarType.Boolean:
                return Variant.FromBoolean(BitConverter.ToInt16(span) != 0);
            case VarType.Long:
                return Variant.FromInt32(BitConverter.ToInt32(span));
            case VarType.Single:
                return Variant.FromSingle(BitConverter.ToSingle(span));
            case VarType.LongLong:
                return Variant.FromInt64(BitConverter.ToInt64(span));
            case VarType.Double:
                return Variant.FromDouble(BitConverter.ToDouble(span));
            case VarType.Currency:
                return Variant.FromCurrency(Currency.FromScaled(BitConverter.ToInt64(span)));
            case VarType.Date:
                return Variant.FromDate(VbaDate.FromSerial(BitConverter.ToDouble(span)));
            case VarType.Byte:
                return Variant.FromByte(span[0]);
            case VarType.String when fixedLength > 0:
                {
                    if (ansi)
                    {
                        return Variant.FromString(FileSystem.Ansi.GetString(span[..fixedLength]));
                    }

                    var characters = new char[fixedLength];
                    for (var i = 0; i < fixedLength; i++)
                    {
                        characters[i] = (char)BitConverter.ToUInt16(span.Slice(i * 2, 2));
                    }

                    return Variant.FromString(new string(characters));
                }

            default:
                return Variant.Empty;
        }
    }

    private readonly record struct Slot(FieldInfo Field, VarType Type, int Offset, int Size, int FixedLength);

    /// <summary>The fields with their offsets under natural alignment, the layout MemorySize measures.</summary>
    private static IEnumerable<Slot> Slots(object record, bool ansi = false)
    {
        var offset = 0;
        foreach (var (field, type, memorySize, memoryAlignment, fixedLength) in RecordLayout.MemoryFields(record))
        {
            var (size, alignment) = ansi ? AnsiLayout(field, record, type, fixedLength, (memorySize, memoryAlignment)) : (memorySize, memoryAlignment);
            offset = (offset + alignment - 1) / alignment * alignment;
            yield return new Slot(field, type, offset, size, fixedLength);
            offset += size;
        }
    }

    /// <summary>A member's size and alignment in the copy VBA hands a DLL: a fixed-length string one ANSI byte a character, a nested record in the same layout, anything else as in memory.</summary>
    private static (int Size, int Alignment) AnsiLayout(FieldInfo field, object record, VarType type, int fixedLength, (int Size, int Alignment) memory) =>
        type == VarType.String && fixedLength > 0 ? (fixedLength, 1)
        : field.GetValue(record) is IVbaRecord nested ? AnsiRecordLayout(nested)
        : memory;

    /// <summary>A record's size and alignment in the copy VBA hands a DLL.</summary>
    private static (int Size, int Alignment) AnsiRecordLayout(object record)
    {
        var end = 0;
        var largest = 1;
        foreach (var (field, type, memorySize, memoryAlignment, fixedLength) in RecordLayout.MemoryFields(record))
        {
            var (size, alignment) = AnsiLayout(field, record, type, fixedLength, (memorySize, memoryAlignment));
            end = ((end + alignment - 1) / alignment * alignment) + size;
            largest = Math.Max(largest, alignment);
        }

        return ((end + largest - 1) / largest * largest, largest);
    }
}
