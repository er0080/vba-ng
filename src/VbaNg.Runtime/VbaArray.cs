using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace VbaNg.Runtime;

/// <summary>
/// A VBA array (ARCHITECTURE.md section 5, "Arrays"; D20): a view of a SAFEARRAY descriptor and
/// the element type. The descriptor and the elements lie in native memory from oleaut32, laid out
/// as OLE Automation lays them out (column-major, the first index varying fastest, with the
/// feature flags VBA sets), and the storage that holds the view owns the array as it owns a
/// String (ROADMAP.md M7 C4): a store copies it (<see cref="Clone"/>), a release destroys it with
/// what its elements own, and a fresh array is a temporary of its statement
/// (<see cref="ObjectRefs.Owned(VbaArray)"/>). A dynamic array before its first ReDim has no
/// descriptor. A For Each locks the array it walks, so a ReDim or an Erase of it raises error 10
/// instead of freeing what the loop reads.
/// </summary>
[DebuggerDisplay("{DebugView,nq}")]
[DebuggerTypeProxy(typeof(VbaArrayDebugView))]
public readonly unsafe partial struct VbaArray : IEquatable<VbaArray>
{
    // VBA allows up to 60 dimensions.
    private const int MaxRank = 60;

    // The feature flags VBA adds for a fixed-size array (OAIdl.h; Memory golden: &H92 for a fixed Long array, &H80 for a dynamic one).
    private const ushort FadfStatic = 0x0002;
    private const ushort FadfFixedSize = 0x0010;

    // An array inside a structure, whose elements are the structure's (OAIdl.h FADF_EMBEDDED; ROADMAP.md M7 C6).
    private const ushort FadfEmbedded = 0x0004;

    private static long live;

#if DEBUG
    private static readonly HashSet<nint> LiveSet = [];
    private static readonly System.Threading.Lock LiveLock = new();
#endif

    private readonly nint descriptor;
    private readonly VarType elementType;

    private VbaArray(nint descriptor, VarType elementType)
    {
        this.descriptor = descriptor;
        this.elementType = elementType;
    }

    /// <summary>How many arrays the runtime has created and not destroyed; the golden replay checks that a case leaves it where it found it (ROADMAP.md WP2).</summary>
    public static long LiveCount => Interlocked.Read(ref live);

    /// <summary>The SAFEARRAY; 0 for a dynamic array that has not been dimensioned.</summary>
    public nint Descriptor => descriptor;

    /// <summary>The declared element type; Variant for <c>Dim a()</c> and <c>Array(...)</c>.</summary>
    public VarType ElementType => elementType;

    /// <summary>Declared with bounds (<c>Dim a(2)</c>): ReDim raises error 10, and Erase resets the elements instead of freeing them.</summary>
    public bool IsFixedSize => descriptor != 0 && (Head->Features & FadfFixedSize) != 0;

    /// <summary>False for a dynamic array that has not been dimensioned; LBound and UBound then raise error 9.</summary>
    public bool IsAllocated => descriptor != 0;

    /// <summary>The number of dimensions; 0 when unallocated.</summary>
    public int Rank => descriptor == 0 ? 0 : Head->Dimensions;

    /// <summary>The number of elements; 0 when unallocated.</summary>
    public int Count
    {
        get
        {
            if (descriptor == 0)
            {
                return 0;
            }

            var count = 1;
            var bounds = BoundsOf;
            for (var i = 0; i < Head->Dimensions; i++)
            {
                count *= (int)bounds[i].Count;
            }

            return count;
        }
    }

    /// <summary>Elements that own something (a BSTR, an interface pointer, a record's references, or any of these inside a Variant), so a release has more to do than free memory (ARCHITECTURE.md D18, D20).</summary>
    public bool HoldsReferences => elementType is VarType.String or VarType.Variant or VarType.Object or VarType.UserDefinedType;

    /// <summary>The bounds in VBA notation, for example "0 To 2, 1 To 3"; empty when unallocated.</summary>
    public string BoundsText
    {
        get
        {
            var text = new StringBuilder();
            for (var i = 1; i <= Rank; i++)
            {
                if (i > 1)
                {
                    text.Append(", ");
                }

                text.Append(VbaErrors.Invariant($"{LBound(i)} To {UBound(i)}"));
            }

            return text.ToString();
        }
    }

    /// <summary>A For Each is walking the array: its lock count is above zero.</summary>
    internal bool IsLocked => descriptor != 0 && Head->Locks != 0;

    private SafeArrayHeader* Head => (SafeArrayHeader*)descriptor;

    /// <summary>The bounds as the descriptor keeps them, the last dimension first.</summary>
    private SafeArrayBound* BoundsOf => (SafeArrayBound*)(descriptor + sizeof(SafeArrayHeader));

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private string DebugView => VbaErrors.Invariant($"{elementType.TypeName()}({BoundsText})");

    public static bool operator ==(VbaArray left, VbaArray right) => left.Equals(right);

    public static bool operator !=(VbaArray left, VbaArray right) => !left.Equals(right);

    /// <summary>A dynamic array with no dimensions yet (<c>Dim a() As Long</c>).</summary>
    public static VbaArray Unallocated(VarType elementType) => new(0, elementType);

    /// <summary>
    /// An array of fixed-length strings (Dim a(1) As String * 3): its elements that hold nothing take
    /// the value a String * n starts with, n Chr(0)s (MS-VBAL 5.2.3.1; Arrays golden), once the array
    /// is made and after a ReDim or an Erase. Every store into an element pads it, so an empty slot
    /// is one no store reached. Returns the array.
    /// </summary>
    public static VbaArray FixedStrings(VbaArray array, int length)
    {
        if (array.descriptor == 0 || array.elementType != VarType.String)
        {
            return array;
        }

        var zeros = new string('\0', length);
        for (var i = 0; i < array.Count; i++)
        {
            if (*(nint*)array.At(i) == 0)
            {
                array.StoreAt(i, Coerce.ToElementType(Variant.FromString(zeros), VarType.String));
            }
        }

        return array;
    }

    /// <summary>Allocates an array with the given bounds, its elements zeroed as VBA's start (0, a null String, False, Empty, Nothing); a lower bound above its upper bound raises error 9. The caller owns the array.</summary>
    public static VbaArray Create(VarType elementType, ReadOnlySpan<(int Lower, int Upper)> bounds, bool fixedSize = false)
    {
        if (bounds.Length is 0 or > MaxRank)
        {
            throw VbaErrors.SubscriptOutOfRange();
        }

        var native = stackalloc SafeArrayBound[bounds.Length];
        var count = Dimensions(bounds, native);
        return Allocate(elementType, native, bounds.Length, count, fixedSize);
    }

    /// <summary>A one-dimensional Variant array holding the given values, as <c>Array(...)</c> builds (MS-VBAL 6.1.2.6 Interaction, Array): a temporary of the current statement.</summary>
    public static VbaArray FromValues(ReadOnlySpan<Variant> values, int lowerBound = 0)
    {
        var array = ObjectRefs.Owned(values.Length == 0 ? Empty(VarType.Variant, lowerBound) : Create(VarType.Variant, [(lowerBound, lowerBound + values.Length - 1)]));
        var data = (Variant*)array.Head->Data;
        for (var i = 0; i < values.Length; i++)
        {
            data[i] = ObjectRefs.Own(values[i]);
        }

        return array;
    }

    /// <summary>An allocated array with no elements: the upper bound is one below the lower bound, as <c>Array()</c> and <c>Split("")</c> produce. The caller owns it.</summary>
    public static VbaArray Empty(VarType elementType, int lowerBound = 0)
    {
        var native = stackalloc SafeArrayBound[1];
        native[0] = new SafeArrayBound { Count = 0, Lower = lowerBound };
        return Allocate(elementType, native, 1, 0, fixedSize: false);
    }

    /// <summary>A one-dimensional array of the given element type holding the values, coerced to that type: a temporary of the current statement (D18, D20).</summary>
    public static VbaArray FromValues(VarType elementType, ReadOnlySpan<Variant> values, int lowerBound = 0)
    {
        var array = ObjectRefs.Owned(values.Length == 0 ? Empty(elementType, lowerBound) : Create(elementType, [(lowerBound, lowerBound + values.Length - 1)]));
        for (var i = 0; i < values.Length; i++)
        {
            array.StoreAt(i, Coerce.ToElementType(values[i], elementType));
        }

        return array;
    }

    /// <summary>The lower bound of a 1-based dimension (MS-VBAL 6.1.2.5 Information, LBound); error 9 when unallocated or out of range.</summary>
    public int LBound(int dimension = 1) => Bound(dimension).Lower;

    public int UBound(int dimension = 1)
    {
        var bound = Bound(dimension);
        return bound.Lower + (int)bound.Count - 1;
    }

    /// <summary>An element as a Variant, a view of what it owns as reads are (D20). A record never travels in a Variant (M7 C3), so an element of an array of records read or written this way is a type mismatch.</summary>
    public Variant this[ReadOnlySpan<int> indices]
    {
        get => Read(Offset(indices));
        set => StoreAt(Offset(indices), Coerce.ToElementType(value, elementType));
    }

    public Variant Get(ReadOnlySpan<int> indices) => this[indices];

    public void Set(ReadOnlySpan<int> indices, Variant value) => this[indices] = value;

    /// <summary>The element at a storage position (the first index varying fastest), as the indexer reads it.</summary>
    public Variant ElementAt(int position) => (uint)position < (uint)Count ? Read(position) : throw VbaErrors.SubscriptOutOfRange();

    /// <summary>VarPtr of an element (ARCHITECTURE.md section 5, "Pointers"): its address in the array's native storage, a record's included (ROADMAP.md M7 C5).</summary>
    public long ElementAddress(ReadOnlySpan<int> indices) => (long)At(Offset(indices));

    /// <summary>An element in the array's native storage as a T, by reference: what a Declare's ByRef parameter receives, the element's address, so the callee reaches the elements after it too (Declares golden). T is the element's own type, or byte for As Any.</summary>
    public ref T ElementRef<T>(scoped ReadOnlySpan<int> indices)
        where T : unmanaged
    {
        var offset = Offset(indices);
        if (elementType == VarType.UserDefinedType || (sizeof(T) != Head->ElementSize && typeof(T) != typeof(byte)))
        {
            throw VbaErrors.TypeMismatch();
        }

        return ref *(T*)At(offset);
    }

    /// <summary>
    /// Assignment copies arrays (MS-VBAL 5.5.1.2.5 Let-coercion to and from arrays): a new dynamic
    /// array with the same bounds whose elements own copies of their own (strings, references,
    /// nested arrays, records). The caller owns it; an unallocated array copies as unallocated.
    /// </summary>
    public VbaArray Clone()
    {
        if (descriptor == 0)
        {
            return this;
        }

        var rank = Rank;
        var native = stackalloc SafeArrayBound[rank];
        CopyBounds(native);
        if (elementType == VarType.UserDefinedType)
        {
            var kind = Kind;
            var records = RecordKind.Allocate(kind, Count);
            kind.CopyInto(Head->Data, records, Count);
            return AllocateRecords(kind, native, rank, records, fixedSize: false);
        }

        var count = Count;
        var copy = Allocate(elementType, native, rank, count, fixedSize: false);
        var from = (byte*)Head->Data;
        var to = (byte*)copy.Head->Data;
        switch (elementType)
        {
            case VarType.String:
                for (var i = 0; i < count; i++)
                {
                    ((nint*)to)[i] = Bstr.Copy(((nint*)from)[i]);
                }

                break;
            case VarType.Object:
                for (var i = 0; i < count; i++)
                {
                    var pointer = ((nint*)from)[i];
                    RuntimeObject.AddRefInterface(pointer);
                    ((nint*)to)[i] = pointer;
                }

                break;
            case VarType.Variant:
                for (var i = 0; i < count; i++)
                {
                    ((Variant*)to)[i] = ObjectRefs.Own(((Variant*)from)[i]);
                }

                break;
            default:
                {
                    var bytes = (long)count * Head->ElementSize;
                    Buffer.MemoryCopy(from, to, bytes, bytes);
                    break;
                }
        }

        return copy;
    }

    /// <summary>ReDim on a declared array variable (MS-VBAL 5.4.3.3): a fixed-size array, or one a For Each is walking, raises error 10.</summary>
    public static void ReDimDeclared(ref VbaArray array, ReadOnlySpan<(int Lower, int Upper)> bounds, bool preserve)
    {
        if (array.IsFixedSize || array.IsLocked)
        {
            throw new VbaException(VbaErrors.ArrayFixedOrLocked);
        }

        Redim(ref array, bounds, preserve, null);
    }

    /// <summary>Erase (MS-VBAL 5.4.3.4): a fixed-size array keeps its storage and its elements go back to their initial values; a dynamic array is destroyed and left unallocated. An array a For Each is walking raises error 10.</summary>
    public static void Erase(ref VbaArray array)
    {
        if (!array.IsAllocated)
        {
            return;
        }

        if (array.IsLocked)
        {
            throw new VbaException(VbaErrors.ArrayFixedOrLocked);
        }

        if (array.IsFixedSize)
        {
            array.Reset();
            return;
        }

        var old = array;
        array = Unallocated(old.elementType);
        old.Destroy();
    }

    public bool Equals(VbaArray other) => descriptor == other.descriptor && elementType == other.elementType;

    public override bool Equals(object? obj) => obj is VbaArray other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(descriptor, elementType);

    /// <summary>The view of an array a Variant or a statement's temporary holds.</summary>
    internal static VbaArray View(nint descriptor, VarType elementType) => new(descriptor, elementType);

    /// <summary>
    /// A fixed-size array member of a user-defined type (ROADMAP.md M7 C6): its elements lie inline
    /// in the record, as VBA lays them out, and this is a descriptor of its own over them, flagged
    /// embedded and fixed-size as OLE Automation marks an array inside a structure, and a temporary
    /// of the current statement. Destroying it frees only the descriptor; the elements stay the
    /// record's, so the record must stay where it is for the statement, as a local, a module-level
    /// variable, an instance's block, and an array's storage do (ARCHITECTURE.md D21).
    /// </summary>
    public static VbaArray Embedded<T>(ref T elements, ReadOnlySpan<(int Lower, int Upper)> bounds)
        where T : struct, IEmbeddedArray =>
        ObjectRefs.Owned(EmbeddedView((nint)Unsafe.AsPointer(ref elements), T.ElementType, bounds, EmbeddedStorage<T>.Kind));

    /// <summary>The initial value of a fixed-size array member of records (ROADMAP.md M7 C9): its inline storage, each element a fresh record.</summary>
    public static T FreshEmbedded<T>()
        where T : unmanaged, IEmbeddedArray
    {
        var elements = default(T);
        if (EmbeddedStorage<T>.Kind is { } kind)
        {
            kind.Fill((nint)(&elements), 0, sizeof(T) / kind.Size);
        }

        return elements;
    }

    /// <summary>A view over elements that lie elsewhere (a record's array member, in stable storage or pinned); the caller destroys it, which frees only the descriptor. Records name their <paramref name="kind"/>, which the view keeps until it is destroyed, since their storage here has no header (ROADMAP.md M7 C9).</summary>
    internal static VbaArray EmbeddedView(nint elements, VarType elementType, ReadOnlySpan<(int Lower, int Upper)> bounds, RecordKind? kind = null)
    {
        var native = stackalloc SafeArrayBound[bounds.Length];
        Dimensions(bounds, native);

        // The Ex form sets the type's flags (BSTR, VARIANT, DISPATCH) as SafeArrayCreate does, so a copy of the array copies what its elements hold; records carry no flag, as AllocateRecords makes them.
        nint created;
        var result = kind is null
            ? SafeArrayNative.SafeArrayAllocDescriptorEx((ushort)elementType, (uint)bounds.Length, out created)
            : SafeArrayNative.SafeArrayAllocDescriptor((uint)bounds.Length, out created);
        if (result < 0 || created == 0)
        {
            throw new VbaException(VbaErrors.OutOfMemory);
        }

        var head = (SafeArrayHeader*)created;
        if (kind is not null)
        {
            head->Features = 0;
            head->Locks = 0;
        }

        head->Features |= FadfEmbedded | FadfStatic | FadfFixedSize;

        // The Ex form leaves the element size to the caller.
        head->ElementSize = (uint)(kind?.Size ?? ElementSize(elementType));
        var stored = (SafeArrayBound*)(created + sizeof(SafeArrayHeader));
        for (var d = 0; d < bounds.Length; d++)
        {
            stored[bounds.Length - 1 - d] = native[d];
        }

        head->Data = elements;
        Track(created);
        if (kind is not null)
        {
            RecordKind.Embed(created, kind);
        }

        return new VbaArray(created, elementType);
    }

    /// <summary>A record's copy takes its own of what an embedded array member's elements hold, which came across byte for byte: strings copied, references added, Variants' contents owned (D18, D20).</summary>
    public static void OwnEmbedded<T>(ref T elements)
        where T : struct, IEmbeddedArray
    {
        var data = (byte*)Unsafe.AsPointer(ref elements);
        if (EmbeddedStorage<T>.Kind is { } kind)
        {
            // Records copy themselves where they lie (ROADMAP.md M7 C9).
            kind.CopyInto((nint)data, (nint)data, Unsafe.SizeOf<T>() / kind.Size);
            return;
        }

        OwnSlots(data, T.ElementType, Unsafe.SizeOf<T>() / ElementSize(T.ElementType));
    }

    /// <summary>What an embedded array member's elements hold goes, and the elements are left empty: a record's release (D18, D20).</summary>
    public static void ReleaseEmbedded<T>(ref T elements)
        where T : struct, IEmbeddedArray
    {
        var data = (byte*)Unsafe.AsPointer(ref elements);
        if (EmbeddedStorage<T>.Kind is { } kind)
        {
            kind.Release((nint)data, 0, Unsafe.SizeOf<T>() / kind.Size);
            return;
        }

        ReleaseSlots(data, T.ElementType, Unsafe.SizeOf<T>() / ElementSize(T.ElementType));
    }

    /// <summary>
    /// An element of a fixed-size array member of fixed-length strings (ROADMAP.md M7 C10), read.
    /// The member lies inline in its record as each element's UTF-16 characters, one after another
    /// at one-byte alignment, as VBA lays it out (Memory golden), with no descriptor over them:
    /// generated code reaches an element through these, and the whole array as a copy.
    /// </summary>
    public static string InlineChars<T>(ref T storage, scoped ReadOnlySpan<int> indices)
        where T : struct, IEmbeddedArray =>
        new(CharsAt(ref storage, indices));

    /// <summary>A store into an element of a fixed-length string array member: the value padded or cut to the element's length (MS-VBAL 5.4.3.1).</summary>
    public static void SetInlineChars<T>(ref T storage, scoped ReadOnlySpan<int> indices, string value)
        where T : struct, IEmbeddedArray
    {
        ArgumentNullException.ThrowIfNull(value);
        var target = CharsAt(ref storage, indices);
        Library.Strings.ToFixed(value, target.Length).AsSpan().CopyTo(target);
    }

    /// <summary>A fixed-length string array member as an array of its own, a copy of its elements and a temporary of the statement: what LBound, UBound, For Each, and a ByRef call see; <see cref="ToInlineChars"/> stores a ByRef call's changes back.</summary>
    public static VbaArray FromInlineChars<T>(in T storage)
        where T : struct, IEmbeddedArray =>
        ObjectRefs.Owned(FromChars(MemoryMarshal.CreateReadOnlySpan(ref Unsafe.As<T, char>(ref Unsafe.AsRef(in storage)), InlineShape<T>.Count * InlineShape<T>.Chars), InlineShape<T>.Bounds, InlineShape<T>.Chars));

    /// <summary>The elements of an array made by <see cref="FromInlineChars"/> back into the member, each padded or cut to its length.</summary>
    public static void ToInlineChars<T>(ref T storage, VbaArray array)
        where T : struct, IEmbeddedArray =>
        array.CopyCharsTo(MemoryMarshal.CreateSpan(ref Unsafe.As<T, char>(ref storage), InlineShape<T>.Count * InlineShape<T>.Chars), InlineShape<T>.Chars);

    /// <summary>A fixed-size array of Strings holding the fixed-length strings that lie one after another in <paramref name="all"/>, <paramref name="length"/> characters each.</summary>
    internal static VbaArray FromChars(ReadOnlySpan<char> all, ReadOnlySpan<(int Lower, int Upper)> bounds, int length)
    {
        var array = Create(VarType.String, bounds, fixedSize: true);
        for (var i = 0; i < all.Length / length; i++)
        {
            array.StoreAt(i, Coerce.ToElementType(Variant.FromString(new string(all.Slice(i * length, length))), VarType.String));
        }

        return array;
    }

    /// <summary>The elements of a String array as fixed-length strings of <paramref name="length"/> characters, one after another into <paramref name="all"/>.</summary>
    internal void CopyCharsTo(Span<char> all, int length)
    {
        for (var i = 0; i < all.Length / length; i++)
        {
            Library.Strings.ToFixed(Coerce.ToString(ElementAt(i)), length).AsSpan().CopyTo(all.Slice(i * length, length));
        }
    }

    /// <summary>An element's characters in a fixed-length string array member, by its indices (the first varying fastest, as in an array's storage); a subscript outside the bounds raises 9.</summary>
    private static Span<char> CharsAt<T>(ref T storage, scoped ReadOnlySpan<int> indices)
        where T : struct, IEmbeddedArray
    {
        var bounds = InlineShape<T>.Bounds;
        if (indices.Length != bounds.Length)
        {
            throw VbaErrors.SubscriptOutOfRange();
        }

        var position = 0;
        var stride = 1;
        for (var d = 0; d < bounds.Length; d++)
        {
            var index = (long)indices[d] - bounds[d].Lower;
            var count = (long)bounds[d].Upper - bounds[d].Lower + 1;
            if ((ulong)index >= (ulong)count)
            {
                throw VbaErrors.SubscriptOutOfRange();
            }

            position += (int)index * stride;
            stride *= (int)count;
        }

        var length = InlineShape<T>.Chars;
        return MemoryMarshal.CreateSpan(ref Unsafe.Add(ref Unsafe.As<T, char>(ref storage), position * length), length);
    }

    /// <summary>The bounds of <see cref="IEmbeddedArray.Bounds"/>, lower and upper in pairs, as dimensions.</summary>
    internal static (int Lower, int Upper)[] BoundPairs(int[] bounds)
    {
        var pairs = new (int Lower, int Upper)[bounds.Length / 2];
        for (var i = 0; i < pairs.Length; i++)
        {
            pairs[i] = (bounds[2 * i], bounds[(2 * i) + 1]);
        }

        return pairs;
    }

    /// <summary>The bytes one element of the type takes in an array's storage, as SafeArrayCreate sizes it.</summary>
    private static int ElementSize(VarType elementType) => elementType switch
    {
        VarType.Byte => 1,
        VarType.Integer or VarType.Boolean => 2,
        VarType.Long or VarType.Single => 4,
        VarType.Variant => sizeof(Variant),
        VarType.Decimal => 16,
        _ => 8,
    };

    /// <summary>
    /// ReDim (MS-VBAL 5.4.3.3): without Preserve the array is replaced; with Preserve the rank must
    /// match and only the last dimension's upper bound may change, else error 9, and the elements
    /// the new bounds keep move to the new array while the ones they drop are released. The same
    /// bounds leave the array where it is (Memory golden). An unallocated array accepts either
    /// form, and the array is untouched when the new bounds raise.
    /// </summary>
    internal static void Redim(ref VbaArray array, ReadOnlySpan<(int Lower, int Upper)> bounds, bool preserve, RecordKind? kind)
    {
        var old = array;
        var keep = preserve && old.IsAllocated;
        var rank = old.Rank;
        if (keep)
        {
            if (bounds.Length != rank)
            {
                throw VbaErrors.SubscriptOutOfRange();
            }

            for (var i = 0; i < rank - 1; i++)
            {
                if (bounds[i].Lower != old.LBound(i + 1) || bounds[i].Upper != old.UBound(i + 1))
                {
                    throw VbaErrors.SubscriptOutOfRange();
                }
            }

            // Preserve may change only the upper bound of the last dimension (Arrays golden).
            if (bounds[rank - 1].Lower != old.LBound(rank))
            {
                throw VbaErrors.SubscriptOutOfRange();
            }

            if (bounds[rank - 1].Upper == old.UBound(rank))
            {
                return;
            }
        }

        VbaArray replacement;
        if (old.elementType == VarType.UserDefinedType)
        {
            var records = kind ?? (old.IsAllocated ? old.Kind : throw new InvalidOperationException("ReDim of an array of records needs its record type (ReDimRecords)."));
            replacement = CreateRecords(records, bounds, fixedSize: false);
        }
        else
        {
            replacement = Create(old.elementType, bounds);
        }

        if (keep)
        {
            // Only the last dimension changed, and it varies slowest, so the kept elements are a prefix of the storage and move as one block.
            var inner = 1;
            for (var i = 1; i < rank; i++)
            {
                inner *= old.UBound(i) - old.LBound(i) + 1;
            }

            var kept = Math.Min(old.UBound(rank), bounds[rank - 1].Upper) - old.LBound(rank) + 1;
            MoveElements(old, replacement, kept * inner);
        }

        array = replacement;
        old.Destroy();
    }

    /// <summary>What the elements own goes (strings freed, references released, nested arrays destroyed, records' references dropped); the slots are left empty and the storage stays.</summary>
    internal void ReleaseElements()
    {
        if (descriptor == 0 || !HoldsReferences)
        {
            return;
        }

        var count = Count;
        if (elementType == VarType.UserDefinedType)
        {
            Kind.Release(Head->Data, 0, count);
            return;
        }

        ReleaseSlots((byte*)Head->Data, elementType, count);
    }

    /// <summary>What elements of a type that holds something own goes (strings freed, references released, Variants emptied); the slots are left empty.</summary>
    private static void ReleaseSlots(byte* data, VarType elementType, int count)
    {
        switch (elementType)
        {
            case VarType.String:
                {
                    var slots = (nint*)data;
                    for (var i = 0; i < count; i++)
                    {
                        var bstr = slots[i];
                        slots[i] = 0;
                        Bstr.Free(bstr);
                    }

                    break;
                }

            case VarType.Object:
                {
                    var slots = (nint*)data;
                    for (var i = 0; i < count; i++)
                    {
                        var pointer = slots[i];
                        slots[i] = 0;
                        RuntimeObject.ReleaseInterface(pointer);
                    }

                    break;
                }

            case VarType.Variant:
                {
                    var slots = (Variant*)data;
                    for (var i = 0; i < count; i++)
                    {
                        ObjectRefs.Release(ref slots[i]);
                    }

                    break;
                }
        }
    }

    /// <summary>Elements that came across byte for byte take their own of what they hold: strings copied, references added, Variants' contents owned (D18, D20).</summary>
    private static void OwnSlots(byte* data, VarType elementType, int count)
    {
        switch (elementType)
        {
            case VarType.String:
                {
                    var slots = (nint*)data;
                    for (var i = 0; i < count; i++)
                    {
                        slots[i] = Bstr.Copy(slots[i]);
                    }

                    break;
                }

            case VarType.Object:
                {
                    var slots = (nint*)data;
                    for (var i = 0; i < count; i++)
                    {
                        RuntimeObject.AddRefInterface(slots[i]);
                    }

                    break;
                }

            case VarType.Variant:
                {
                    var slots = (Variant*)data;
                    for (var i = 0; i < count; i++)
                    {
                        slots[i] = ObjectRefs.Own(slots[i]);
                    }

                    break;
                }
        }
    }

    /// <summary>Releases what the elements own and frees the array. The caller forgets every view of it; an array a For Each is walking never gets here.</summary>
    internal void Destroy()
    {
        if (descriptor == 0)
        {
            return;
        }

        if ((Head->Features & FadfEmbedded) != 0)
        {
            // An embedded array's elements are its record's: only the descriptor goes (ROADMAP.md M7 C6).
            if (elementType == VarType.UserDefinedType)
            {
                RecordKind.Unembed(descriptor);
            }

            Untrack(descriptor);
            Head->Data = 0;
            _ = SafeArrayNative.SafeArrayDestroyDescriptor(descriptor);
            return;
        }

        ReleaseElements();
        Untrack(descriptor);
        var head = Head;
        head->Features &= unchecked((ushort)~(FadfStatic | FadfFixedSize));
        if (elementType == VarType.UserDefinedType)
        {
            RecordKind.Free(head->Data);
            head->Data = 0;
            _ = SafeArrayNative.SafeArrayDestroyDescriptor(descriptor);
            return;
        }

        _ = SafeArrayNative.SafeArrayDestroy(descriptor);
    }

    /// <summary>A For Each starts walking the array.</summary>
    internal void Lock()
    {
        if (descriptor != 0)
        {
            Head->Locks++;
        }
    }

    /// <summary>The For Each that walked the array is done with it.</summary>
    internal void Unlock()
    {
        if (descriptor != 0 && Head->Locks != 0)
        {
            Head->Locks--;
        }
    }

    private static VbaArray Allocate(VarType elementType, SafeArrayBound* bounds, int rank, int count, bool fixedSize)
    {
        if (elementType == VarType.UserDefinedType)
        {
            throw new ArgumentException("An array of records is made by CreateRecords.", nameof(elementType));
        }

        var created = SafeArrayNative.SafeArrayCreate((ushort)elementType, (uint)rank, bounds);
        if (created == 0)
        {
            throw new VbaException(VbaErrors.OutOfMemory);
        }

        Track(created);
        var head = (SafeArrayHeader*)created;
        if (fixedSize)
        {
            head->Features |= FadfStatic | FadfFixedSize;
        }

        if (count > 0)
        {
            NativeMemory.Clear((void*)head->Data, (nuint)((long)count * head->ElementSize));
        }

        return new VbaArray(created, elementType);
    }

    /// <summary>Checks the bounds of a new array and writes them in dimension order: error 9 for a lower bound above its upper bound, error 7 for more elements than fit. The element count.</summary>
    private static int Dimensions(ReadOnlySpan<(int Lower, int Upper)> bounds, SafeArrayBound* native)
    {
        long count = 1;
        for (var i = 0; i < bounds.Length; i++)
        {
            var (lower, upper) = bounds[i];
            if (lower > upper)
            {
                throw VbaErrors.SubscriptOutOfRange();
            }

            var elements = (long)upper - lower + 1;
            count *= elements;
            if (count > int.MaxValue)
            {
                throw new VbaException(VbaErrors.OutOfMemory);
            }

            native[i] = new SafeArrayBound { Count = (uint)elements, Lower = lower };
        }

        return (int)count;
    }

    /// <summary>Moves the first elements of one array into another of the same type: what they own goes with them, and their old slots are left empty.</summary>
    private static void MoveElements(VbaArray from, VbaArray to, int count)
    {
        if (from.elementType == VarType.UserDefinedType)
        {
            // The fresh records the move overwrites own their fixed-size array members; the moved records' bytes carry what they own.
            to.Kind.Release(to.Head->Data, 0, count);
        }

        var bytes = (long)count * from.Head->ElementSize;
        Buffer.MemoryCopy((void*)from.Head->Data, (void*)to.Head->Data, bytes, bytes);
        NativeMemory.Clear((void*)from.Head->Data, (nuint)bytes);
    }

    /// <summary>The strings and arrays the elements of an array a DLL made hold become the runtime's too (D20).</summary>
    internal void AdoptElements()
    {
        if (descriptor == 0)
        {
            return;
        }

        var count = Count;
        if (elementType == VarType.String)
        {
            for (var i = 0; i < count; i++)
            {
                Bstr.Adopted(((nint*)Head->Data)[i]);
            }
        }
        else if (elementType == VarType.Variant)
        {
            for (var i = 0; i < count; i++)
            {
                Native.AdoptValue(((Variant*)Head->Data)[i]);
            }
        }
    }

    /// <summary>A SAFEARRAY a DLL created and handed the runtime (a Declare's result, a ByRef array it replaced): the runtime destroys it as its own from here.</summary>
    internal static void Adopted(nint descriptor)
    {
        if (descriptor != 0)
        {
            Track(descriptor);
        }
    }

    /// <summary>A SAFEARRAY of the runtime's that a DLL destroyed or dropped: it leaves the live arrays without being destroyed again.</summary>
    internal static void Abandoned(nint descriptor)
    {
        if (descriptor == 0)
        {
            return;
        }

#if DEBUG
        lock (LiveLock)
        {
            if (!LiveSet.Remove(descriptor))
            {
                return;
            }
        }
#endif
        Interlocked.Decrement(ref live);
    }

    private static void Track(nint created)
    {
        Interlocked.Increment(ref live);
#if DEBUG
        lock (LiveLock)
        {
            LiveSet.Add(created);
        }
#endif
    }

    private static void Untrack(nint destroyed)
    {
#if DEBUG
        lock (LiveLock)
        {
            if (!LiveSet.Remove(destroyed))
            {
                throw new InvalidOperationException(VbaErrors.Invariant($"SAFEARRAY 0x{destroyed:X} is not a live array of the runtime (destroyed twice, or never created here)."));
            }
        }
#endif
        Interlocked.Decrement(ref live);
    }

    private Variant Read(int position)
    {
        var at = At(position);
        return elementType switch
        {
            VarType.Integer => Variant.FromInt16(*(short*)at),
            VarType.Long => Variant.FromInt32(*(int*)at),
            VarType.LongLong => Variant.FromInt64(*(long*)at),
            VarType.Single => Variant.FromSingle(*(float*)at),
            VarType.Double => Variant.FromDouble(*(double*)at),
            VarType.Currency => Variant.FromCurrency(Currency.FromScaled(*(long*)at)),
            VarType.Date => Variant.FromDate(VbaDate.FromSerial(*(double*)at)),
            VarType.Boolean => Variant.FromBoolean(*(short*)at != 0),
            VarType.Byte => Variant.FromByte(*at),
            VarType.String => Variant.ViewString(VbaString.Adopt(*(nint*)at)),
            VarType.Object => Variant.ViewInterface(*(nint*)at),
            VarType.Variant => *(Variant*)at,
            _ => throw VbaErrors.TypeMismatch(),
        };
    }

    /// <summary>Stores a value already coerced to the element type: the element takes a copy of a string, a reference on an object, or its own copy of a Variant's contents, and what it held goes (D18, D20).</summary>
    private void StoreAt(int position, in Variant value)
    {
        var at = At(position);
        switch (elementType)
        {
            case VarType.Integer:
                *(short*)at = value.AsInt16();
                break;
            case VarType.Long:
                *(int*)at = value.AsInt32();
                break;
            case VarType.LongLong:
                *(long*)at = value.AsInt64();
                break;
            case VarType.Single:
                *(float*)at = value.AsSingle();
                break;
            case VarType.Double:
                *(double*)at = value.AsDouble();
                break;
            case VarType.Currency:
                *(long*)at = value.AsCurrency().Scaled;
                break;
            case VarType.Date:
                *(double*)at = value.AsDate().Serial;
                break;
            case VarType.Boolean:
                *(short*)at = value.AsBoolean() ? (short)-1 : (short)0;
                break;
            case VarType.Byte:
                *at = value.AsByte();
                break;
            case VarType.String:
                ObjectRefs.Assign(ref *(VbaString*)at, value.AsVbaString());
                break;
            case VarType.Object:
                {
                    var pointer = value.InterfacePointer;
                    RuntimeObject.AddRefInterface(pointer);
                    var old = *(nint*)at;
                    *(nint*)at = pointer;
                    RuntimeObject.ReleaseInterface(old);
                    break;
                }

            case VarType.Variant:
                ObjectRefs.Assign(ref *(Variant*)at, value);
                break;
            default:
                throw VbaErrors.TypeMismatch();
        }
    }

    /// <summary>Every element back to its initial value, what they owned released: Erase of a fixed-size array.</summary>
    private void Reset()
    {
        ReleaseElements();
        if (elementType == VarType.UserDefinedType)
        {
            Kind.Fill(Head->Data, 0, Count);
            return;
        }

        NativeMemory.Clear((void*)Head->Data, (nuint)((long)Count * Head->ElementSize));
    }

    /// <summary>The bounds in dimension order, the first dimension first, as SafeArrayCreate takes them.</summary>
    private void CopyBounds(SafeArrayBound* target)
    {
        var rank = Rank;
        for (var d = 0; d < rank; d++)
        {
            target[d] = BoundsOf[rank - 1 - d];
        }
    }

    private SafeArrayBound Bound(int dimension)
    {
        var rank = Rank;
        if (dimension < 1 || dimension > rank)
        {
            throw VbaErrors.SubscriptOutOfRange();
        }

        return BoundsOf[rank - dimension];
    }

    /// <summary>The storage position of an element: error 9 for the wrong number of indices or one out of its bounds.</summary>
    private int Offset(ReadOnlySpan<int> indices)
    {
        if (descriptor == 0 || indices.Length != Head->Dimensions)
        {
            throw VbaErrors.SubscriptOutOfRange();
        }

        var bounds = BoundsOf;
        var rank = indices.Length;
        var offset = 0;
        var stride = 1;
        for (var i = 0; i < rank; i++)
        {
            var bound = bounds[rank - 1 - i];
            var index = (long)indices[i] - bound.Lower;
            if ((ulong)index >= bound.Count)
            {
                throw VbaErrors.SubscriptOutOfRange();
            }

            offset += (int)index * stride;
            stride *= (int)bound.Count;
        }

        return offset;
    }

    private byte* At(int position) => (byte*)Head->Data + ((nint)position * (nint)Head->ElementSize);

    /// <summary>The SAFEARRAY header (OAIdl.h), 24 bytes on x64; the bounds follow it.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct SafeArrayHeader
    {
        public ushort Dimensions;
        public ushort Features;
        public uint ElementSize;
        public uint Locks;
        public nint Data;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SafeArrayBound
    {
        public uint Count;
        public int Lower;
    }
}

/// <summary>The oleaut32 entry points the runtime allocates and frees arrays with, so an array is one COM can take as it is (ARCHITECTURE.md section 5, "Arrays").</summary>
internal static unsafe partial class SafeArrayNative
{
    [LibraryImport("oleaut32.dll")]
    internal static partial int VariantCopy(Variant* destination, Variant* source);

    [LibraryImport("oleaut32.dll")]
    internal static partial nint SafeArrayCreate(ushort vt, uint dimensions, void* bounds);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayDestroy(nint array);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayAllocDescriptor(uint dimensions, out nint array);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayAllocDescriptorEx(ushort vt, uint dimensions, out nint array);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayDestroyDescriptor(nint array);
}

/// <summary>
/// What the locals window shows under an array (ARCHITECTURE.md section 9): its elements with
/// their VBA subscripts, as VBA's locals window lists them, read from the native storage the
/// debugger cannot see into (D20). The first thousand, so a large array stays browsable.
/// </summary>
internal sealed class VbaArrayDebugView(VbaArray array)
{
    private const int Shown = 1000;

    [DebuggerBrowsable(DebuggerBrowsableState.RootHidden)]
    public VbaArrayElementView[] Elements
    {
        get
        {
            var count = Math.Min(array.Count, Shown);
            var elements = new VbaArrayElementView[count];
            for (var position = 0; position < count; position++)
            {
                elements[position] = new VbaArrayElementView(Subscripts(position), array.ElementType == VarType.UserDefinedType ? array.BoxRecord(position) : array.ElementAt(position));
            }

            return elements;
        }
    }

    /// <summary>The subscripts of a storage position, the first varying fastest: "(1)", "(0, 2)".</summary>
    private string Subscripts(int position)
    {
        var text = new StringBuilder("(");
        for (var dimension = 1; dimension <= array.Rank; dimension++)
        {
            var count = array.UBound(dimension) - array.LBound(dimension) + 1;
            text.Append(dimension > 1 ? ", " : string.Empty).Append(VbaErrors.Invariant($"{array.LBound(dimension) + (position % count)}"));
            position /= count;
        }

        return text.Append(')').ToString();
    }
}

/// <summary>One element of an array in the locals window: its subscripts as the name, its value as the value.</summary>
[DebuggerDisplay("{Value}", Name = "{Subscripts,nq}")]
internal sealed record VbaArrayElementView(string Subscripts, object Value);

/// <summary>
/// A fixed-size array member's inline storage in a generated record (ROADMAP.md M7 C6): the
/// element type and the bounds, which the file statements, LSet, and the debugger read.
/// </summary>
public interface IEmbeddedArray
{
    static abstract VarType ElementType { get; }

    /// <summary>The bounds in pairs, lower then upper, the first dimension first.</summary>
    static abstract int[] Bounds { get; }
}

/// <summary>What the locals window shows under a fixed-length string array member of a record (ARCHITECTURE.md section 9; ROADMAP.md M7 C10): its elements under their subscripts, read from a copy.</summary>
/// <typeparam name="T">The member's inline storage.</typeparam>
public sealed class EmbeddedCharsView<T>(T elements)
    where T : struct, IEmbeddedArray
{
    [DebuggerBrowsable(DebuggerBrowsableState.RootHidden)]
    public object[] Elements
    {
        get
        {
            var local = elements;
            var copy = VbaArray.FromChars(MemoryMarshal.CreateReadOnlySpan(ref Unsafe.As<T, char>(ref local), InlineShape<T>.Count * InlineShape<T>.Chars), InlineShape<T>.Bounds, InlineShape<T>.Chars);
            try
            {
                return new VbaArrayDebugView(copy).Elements;
            }
            finally
            {
                copy.Destroy();
            }
        }
    }
}

/// <summary>What the locals window shows under an array member of a record (ARCHITECTURE.md section 9): its elements under their subscripts, read through a view of a pinned copy.</summary>
/// <typeparam name="T">The member's inline storage.</typeparam>
public sealed class EmbeddedArrayView<T>(T elements)
    where T : struct, IEmbeddedArray
{
    [DebuggerBrowsable(DebuggerBrowsableState.RootHidden)]
    public object[] Elements
    {
        get
        {
            object box = elements;
            var handle = GCHandle.Alloc(box, GCHandleType.Pinned);
            var view = VbaArray.EmbeddedView(handle.AddrOfPinnedObject(), T.ElementType, VbaArray.BoundPairs(T.Bounds), EmbeddedStorage<T>.Kind);
            try
            {
                return new VbaArrayDebugView(view).Elements;
            }
            finally
            {
                view.Destroy();
                handle.Free();
            }
        }
    }
}

/// <summary>The shape of a fixed-length string array member's inline storage (ROADMAP.md M7 C10), computed once per type: its bounds, its element count, and the characters of an element.</summary>
/// <typeparam name="T">The member's inline storage.</typeparam>
internal static class InlineShape<T>
    where T : struct, IEmbeddedArray
{
    public static readonly (int Lower, int Upper)[] Bounds = VbaArray.BoundPairs(T.Bounds);

    public static readonly int Count = Bounds.Aggregate(1, (n, b) => n * (b.Upper - b.Lower + 1));

    public static readonly int Chars = Unsafe.SizeOf<T>() / Count / sizeof(char);
}
