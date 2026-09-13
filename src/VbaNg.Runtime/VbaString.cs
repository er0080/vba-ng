using System.Diagnostics;

namespace VbaNg.Runtime;

/// <summary>
/// A VBA String: a BSTR (ARCHITECTURE.md section 5, D20). The struct is the pointer, so copying
/// it aliases the string; the storage that holds a string owns it and frees it through
/// <see cref="Free"/> when the storage is reassigned or goes out of scope, a temporary is freed
/// when its statement ends (<see cref="ObjectRefs"/>), and a store makes its own <see cref="Copy"/>.
/// <see cref="Null"/> is <c>vbNullString</c>, a null pointer that reads as empty; an assigned
/// <c>""</c> is a real BSTR of length zero (Memory golden). <see cref="Length"/> counts whole
/// code units and <see cref="ByteLength"/> every byte, so the byte functions are exact.
/// </summary>
[DebuggerDisplay("{DebugView,nq}")]
public readonly struct VbaString : IEquatable<VbaString>
{
    private readonly nint pointer;

    private VbaString(nint pointer) => this.pointer = pointer;

    /// <summary><c>vbNullString</c>: no BSTR at all.</summary>
    public static VbaString Null => default;

    /// <summary>The BSTR pointer: the address of the first code unit, or 0.</summary>
    public nint Pointer => pointer;

    public bool IsNull => pointer == 0;

    public int ByteLength => Bstr.ByteLength(pointer);

    /// <summary>What Len reports: the whole code units.</summary>
    public int Length => ByteLength / 2;

    public ReadOnlySpan<char> Chars => Bstr.Chars(pointer);

    public ReadOnlySpan<byte> Bytes => Bstr.Bytes(pointer);

    /// <summary>A new BSTR with these code units.</summary>
    public static VbaString Alloc(ReadOnlySpan<char> text) => new(Bstr.Alloc(text));

    /// <summary>A new BSTR with exactly these bytes.</summary>
    public static VbaString AllocBytes(ReadOnlySpan<byte> bytes) => new(Bstr.AllocBytes(bytes));

    /// <summary>A new BSTR of the given length whose code units are still to be written through <see cref="Bstr.MutableChars"/>.</summary>
    public static VbaString AllocLength(int length) => new(Bstr.AllocLength(length));

    /// <summary>A new BSTR holding the bytes of both, in order.</summary>
    public static VbaString Concat(ReadOnlySpan<byte> left, ReadOnlySpan<byte> right)
    {
        var result = new VbaString(Bstr.AllocByteLength(left.Length + right.Length));
        var bytes = Bstr.MutableBytes(result.pointer);
        left.CopyTo(bytes);
        right.CopyTo(bytes[left.Length..]);
        return result;
    }

    /// <summary>A new BSTR with the text of a .NET string; null is <c>vbNullString</c>.</summary>
    public static VbaString From(string? text) => text is null ? Null : Alloc(text.AsSpan());

    /// <summary>A string literal of a module: a BSTR that lives as long as the process, viewed by every use and copied by every store (D20).</summary>
    public static VbaString Literal(string text) => new(Bstr.AllocLiteral(text.AsSpan()));

    /// <summary>Adopts a BSTR the caller owns (one the runtime allocated, or one a COM method handed over).</summary>
    public static VbaString Adopt(nint bstr) => new(bstr);

    /// <summary>A new BSTR with the text, owned by the current statement: what a string literal or a String-returning intrinsic evaluates to in generated code (D20); a store copies it.</summary>
    public static VbaString Temporary(string? text) => text is null ? Null : Variant.FromString(text).AsVbaString();

    /// <summary>The BSTR handed over to the current statement, which frees it when it ends unless a store copied it first.</summary>
    public static VbaString Temporary(VbaString text) => Variant.FromVbaString(text).AsVbaString();

    /// <summary>A BSTR with the same bytes, owned by the caller.</summary>
    public VbaString Copy() => new(Bstr.Copy(pointer));

    /// <summary>Frees the BSTR. Every copy of this struct is dangling afterwards; the storage that owned it must forget it.</summary>
    public void Free() => Bstr.Free(pointer);

    /// <summary>The whole code units as a .NET string; a null or empty BSTR gives the empty string.</summary>
    public override string ToString() => pointer == 0 || ByteLength < 2 ? string.Empty : new string(Chars);

    /// <summary>Byte-for-byte equality, so a null BSTR equals a zero-length one and an odd byte counts.</summary>
    public bool Equals(VbaString other) => Bytes.SequenceEqual(other.Bytes);

    public override bool Equals(object? obj) => obj is VbaString other && Equals(other);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.AddBytes(Bytes);
        return hash.ToHashCode();
    }

    public static bool operator ==(VbaString left, VbaString right) => left.Equals(right);

    public static bool operator !=(VbaString left, VbaString right) => !left.Equals(right);

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private string DebugView => pointer == 0 ? "vbNullString" : VbaErrors.Invariant($"\"{ToString()}\" ({ByteLength} bytes)");
}
