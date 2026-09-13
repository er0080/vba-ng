using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>Arrays of user-defined types, Byte arrays as strings, and the array statements on Variants, for generated code.</summary>
public readonly unsafe partial struct VbaArray
{
    /// <summary>An array of records (MS-VBAL 5.2.3.3), each element a fresh value, laid out one after another in the array's native storage as VBA lays them out (ROADMAP.md M7 C5; Memory golden).</summary>
    public static VbaArray CreateRecords<T>(ReadOnlySpan<(int Lower, int Upper)> bounds, bool fixedSize = false)
        where T : unmanaged, IVbaRecord<T> =>
        CreateRecords(RecordKind<T>.Instance, bounds, fixedSize);

    /// <summary>ReDim on a declared array of records: the record type makes the elements even while the array has none. A fixed-size array, or one a For Each is walking, raises error 10.</summary>
    public static void ReDimRecords<T>(ref VbaArray array, ReadOnlySpan<(int Lower, int Upper)> bounds, bool preserve)
        where T : unmanaged, IVbaRecord<T>
    {
        if (array.IsFixedSize || array.IsLocked)
        {
            throw new VbaException(VbaErrors.ArrayFixedOrLocked);
        }

        Redim(ref array, bounds, preserve, RecordKind<T>.Instance);
    }

    /// <summary>
    /// An element of an array of records, as its storage: generated code reads it, writes its
    /// fields, and aliases it for a With block through the reference (MS-VBAL 5.4.2.7). The
    /// indices are scoped, so the reference outlives them.
    /// </summary>
    public ref T Record<T>(scoped ReadOnlySpan<int> indices)
        where T : unmanaged, IVbaRecord<T>
    {
        if (elementType != VarType.UserDefinedType)
        {
            throw VbaErrors.TypeMismatch();
        }

        var offset = Offset(indices);
        if (Head->ElementSize != (uint)sizeof(T))
        {
            throw VbaErrors.TypeMismatch();
        }

        return ref *(T*)At(offset);
    }

    /// <summary>A String assigned to a Byte array becomes its UTF-16 bytes (MS-VBAL 5.5.1.2.5 Let-coercion to and from Byte arrays): a temporary of the current statement, which the store copies.</summary>
    public static VbaArray FromString(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return ObjectRefs.Owned(FromBytes(System.Text.Encoding.Unicode.GetBytes(value)));
    }

    /// <summary>A String variable assigned to a Byte array: every byte of its BSTR (D20), a temporary of the current statement.</summary>
    public static VbaArray FromString(VbaString value) => ObjectRefs.Owned(FromBytes(value.Bytes));

    /// <summary>A Byte array holding exactly these bytes, every byte of a BSTR, an odd one included (D20); the caller owns it.</summary>
    public static VbaArray FromBytes(ReadOnlySpan<byte> bytes)
    {
        var array = bytes.Length == 0 ? Empty(VarType.Byte) : Create(VarType.Byte, [(0, bytes.Length - 1)]);
        bytes.CopyTo(new Span<byte>((void*)array.Head->Data, bytes.Length));
        return array;
    }

    /// <summary>ReDim on a Variant (MS-VBAL 5.4.3.3): resizes the array it holds, or gives it a new array of the element type.</summary>
    public static void ReDimVariant(ref Variant target, VarType elementType, ReadOnlySpan<(int Lower, int Upper)> bounds, bool preserve)
    {
        if (target.IsArray)
        {
            var array = target.AsArray();
            if (array.IsFixedSize || array.IsLocked)
            {
                throw new VbaException(VbaErrors.ArrayFixedOrLocked);
            }

            Redim(ref array, bounds, preserve, null);
            target = Variant.FromArray(array);
            return;
        }

        // The Variant takes a fresh array of its own; whatever it held (a string, an object) goes (D18, D20).
        var created = Create(elementType, bounds);
        ObjectRefs.Release(ref target);
        target = Variant.FromArray(created);
    }

    /// <summary>Erase on a Variant (MS-VBAL 5.4.3.4): the array it holds is erased; anything else raises 13.</summary>
    public static void Erase(ref Variant target)
    {
        if (!target.IsArray)
        {
            throw VbaErrors.TypeMismatch();
        }

        var array = target.AsArray();
        Erase(ref array);
        target = Variant.FromArray(array);
    }

    /// <summary>A Byte array assigned to a String variable: a new BSTR of exactly its bytes, owned by the current statement (Strings golden: an odd count survives).</summary>
    public VbaString ToText() => VbaString.Temporary(VbaString.AllocBytes(ByteSpan));

    /// <summary>The bytes of a Byte array, in order.</summary>
    public byte[] ToBytes() => ByteSpan.ToArray();

    /// <summary>A Byte array assigned to a String is read as UTF-16 bytes; an odd trailing byte is dropped (ROADMAP.md backlog: odd-byte strings).</summary>
    public string ToStringValue()
    {
        var bytes = ByteSpan;
        return System.Text.Encoding.Unicode.GetString(bytes[..(bytes.Length - (bytes.Length % 2))]);
    }

    /// <summary>The element at a storage position, boxed with what it owns, for the layouts of Put, Get, Len, and LSet, which read fields by reflection.</summary>
    internal object BoxRecord(int position) => Kind.Box(Head->Data, position);

    /// <summary>Writes a boxed element back to its position; what the box owns moves with it.</summary>
    internal void UnboxRecord(int position, object box) => Kind.Unbox(Head->Data, position, box);

    /// <summary>The elements of a Byte array as bytes; any other array is a type mismatch.</summary>
    private ReadOnlySpan<byte> ByteSpan => elementType != VarType.Byte
        ? throw VbaErrors.TypeMismatch()
        : descriptor == 0 ? default : new ReadOnlySpan<byte>((void*)Head->Data, Count);

    /// <summary>The operations of this array's record type, which its storage names (ROADMAP.md M7 C5), or, for records inline in another record, the view (C9).</summary>
    private RecordKind Kind => (Head->Features & FadfEmbedded) != 0 ? RecordKind.OfEmbedded(descriptor) : RecordKind.Of(Head->Data);

    private static VbaArray CreateRecords(RecordKind kind, ReadOnlySpan<(int Lower, int Upper)> bounds, bool fixedSize)
    {
        if (bounds.Length is 0 or > MaxRank)
        {
            throw VbaErrors.SubscriptOutOfRange();
        }

        var native = stackalloc SafeArrayBound[bounds.Length];
        var count = Dimensions(bounds, native);
        var data = RecordKind.Allocate(kind, count);
        kind.Fill(data, 0, count);
        return AllocateRecords(kind, native, bounds.Length, data, fixedSize);
    }

    /// <summary>A descriptor of its own for an array of records: the bounds and flags as for any array, the element size the record's, and the data the records one after another.</summary>
    private static VbaArray AllocateRecords(RecordKind kind, SafeArrayBound* bounds, int rank, nint data, bool fixedSize)
    {
        if (SafeArrayNative.SafeArrayAllocDescriptor((uint)rank, out var created) < 0 || created == 0)
        {
            RecordKind.Free(data);
            throw new VbaException(VbaErrors.OutOfMemory);
        }

        var head = (SafeArrayHeader*)created;
        head->Features = fixedSize ? (ushort)(FadfStatic | FadfFixedSize) : (ushort)0;
        head->ElementSize = (uint)kind.Size;
        head->Locks = 0;
        var stored = (SafeArrayBound*)(created + sizeof(SafeArrayHeader));
        for (var d = 0; d < rank; d++)
        {
            stored[rank - 1 - d] = bounds[d];
        }

        head->Data = data;
        Track(created);
        return new VbaArray(created, VarType.UserDefinedType);
    }
}

/// <summary>
/// The operations of one user-defined type over its arrays' native storage: fresh values,
/// copies, and releases, without reflection (ROADMAP.md M7 C3, C5). The storage holds the
/// records one after another, as VBA lays an array of them out, with the type's number in a
/// header before the first, so an array reached without its type (a release, a Variant's
/// ReDim) finds its operations.
/// </summary>
internal abstract unsafe class RecordKind
{
    private const int HeaderSize = 16;

    private static readonly List<RecordKind> Kinds = [];

    private static readonly Lock KindsLock = new();

    private static readonly Dictionary<nint, RecordKind> EmbeddedViews = [];

    private readonly int number;

    protected RecordKind(int size)
    {
        Size = size;
        lock (KindsLock)
        {
            number = Kinds.Count;
            Kinds.Add(this);
        }
    }

    /// <summary>The bytes one record takes in the array, the stride between elements.</summary>
    public int Size { get; }

    /// <summary>Zeroed storage for this many records of the kind; the caller fills it.</summary>
    public static nint Allocate(RecordKind kind, int count)
    {
        ArgumentNullException.ThrowIfNull(kind);
        var block = (byte*)NativeMemory.AllocZeroed((nuint)(HeaderSize + ((long)count * kind.Size)));
        *(int*)block = kind.number;
        return (nint)(block + HeaderSize);
    }

    /// <summary>Frees storage that <see cref="Allocate"/> made; what the records owned must already be released.</summary>
    public static void Free(nint data) => NativeMemory.Free((byte*)data - HeaderSize);

    /// <summary>The kind whose records a block of storage holds.</summary>
    public static RecordKind Of(nint data)
    {
        var index = *(int*)((byte*)data - HeaderSize);
        lock (KindsLock)
        {
            return Kinds[index];
        }
    }

    /// <summary>A view over records inline in another record (ROADMAP.md M7 C9) has no header before them: its kind is kept by descriptor until the view is destroyed.</summary>
    public static void Embed(nint descriptor, RecordKind kind)
    {
        lock (KindsLock)
        {
            EmbeddedViews[descriptor] = kind;
        }
    }

    public static void Unembed(nint descriptor)
    {
        lock (KindsLock)
        {
            EmbeddedViews.Remove(descriptor);
        }
    }

    public static RecordKind OfEmbedded(nint descriptor)
    {
        lock (KindsLock)
        {
            return EmbeddedViews[descriptor];
        }
    }

    /// <summary>The kind of the records an array member's inline storage holds, found by reflection on its element field; null for elements that are not records.</summary>
    public static RecordKind? OfStorage(Type storage)
    {
        ArgumentNullException.ThrowIfNull(storage);
        var element = storage.GetField("First")!.FieldType;
        return typeof(IVbaRecord).IsAssignableFrom(element)
            ? (RecordKind)typeof(RecordKind<>).MakeGenericType(element).GetField("Instance")!.GetValue(null)!
            : null;
    }

    /// <summary>The records in the range become fresh values; what they owned must already be released.</summary>
    public abstract void Fill(nint data, int start, int count);

    /// <summary>What the records in the range own goes; their storage stays.</summary>
    public abstract void Release(nint data, int start, int count);

    /// <summary>A copy of each record into other storage: strings, arrays, and references of their own (D18, D20).</summary>
    public abstract void CopyInto(nint source, nint target, int count);

    public abstract object Box(nint data, int position);

    public abstract void Unbox(nint data, int position, object box);
}

internal sealed unsafe class RecordKind<T> : RecordKind
    where T : unmanaged, IVbaRecord<T>
{
    public static readonly RecordKind<T> Instance = new();

    private RecordKind()
        : base(sizeof(T))
    {
    }

    public override void Fill(nint data, int start, int count)
    {
        var items = (T*)data;
        for (var i = start; i < start + count; i++)
        {
            items[i] = T.Fresh();
        }
    }

    public override void Release(nint data, int start, int count)
    {
        var items = (T*)data;
        for (var i = start; i < start + count; i++)
        {
            items[i].ReleaseReferences();
        }
    }

    public override void CopyInto(nint source, nint target, int count)
    {
        var from = (T*)source;
        var to = (T*)target;
        for (var i = 0; i < count; i++)
        {
            to[i] = from[i].Copy();
        }
    }

    public override object Box(nint data, int position) => ((T*)data)[position];

    public override void Unbox(nint data, int position, object box) => ((T*)data)[position] = (T)box;
}

/// <summary>The record kind of an array member's inline storage type, looked up once per type (ROADMAP.md M7 C9); null when its elements are not records.</summary>
/// <typeparam name="T">The member's inline storage.</typeparam>
internal static class EmbeddedStorage<T>
    where T : struct, IEmbeddedArray
{
    public static readonly RecordKind? Kind = T.ElementType == VarType.UserDefinedType ? RecordKind.OfStorage(typeof(T)) : null;
}
