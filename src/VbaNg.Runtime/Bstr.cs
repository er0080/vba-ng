using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>
/// BSTR allocation (ARCHITECTURE.md section 5, D20): SysAllocStringLen and its kin from
/// oleaut32, so a string the runtime makes is one Excel and any COM method can take as it is,
/// with the four-byte byte-length prefix, UTF-16 code units, and a terminator, and any byte
/// count, odd included. A debug build keeps the set of live allocations, so a double free
/// throws instead of corrupting the heap, and <see cref="LiveCount"/> lets a test check that
/// everything a case allocated was freed (ROADMAP.md WP2 exit criterion).
/// </summary>
public static partial class Bstr
{
    private static long live;

#if DEBUG
    private static readonly HashSet<nint> LiveSet = [];
    private static readonly Lock LiveLock = new();
#endif

    /// <summary>How many BSTRs the runtime has allocated and not freed.</summary>
    public static long LiveCount => Interlocked.Read(ref live);

    /// <summary>A BSTR holding the code units; an empty span gives a real zero-length BSTR, as VBA's <c>""</c> is (Memory golden).</summary>
    public static unsafe nint Alloc(ReadOnlySpan<char> text)
    {
        fixed (char* chars = text)
        {
            return Track(SysAllocStringLen((nint)chars, (uint)text.Length));
        }
    }

    /// <summary>A BSTR holding exactly these bytes, so a byte string of odd length is exact (Strings golden).</summary>
    public static unsafe nint AllocBytes(ReadOnlySpan<byte> bytes)
    {
        fixed (byte* data = bytes)
        {
            return Track(SysAllocStringByteLen((nint)data, (uint)bytes.Length));
        }
    }

    /// <summary>A BSTR of the given length in code units, uninitialized; the caller writes them through <see cref="MutableChars"/>.</summary>
    public static nint AllocLength(int length) => Track(SysAllocStringLen(0, (uint)length));

    /// <summary>A BSTR of the given byte count, uninitialized; the caller writes them through <see cref="MutableBytes"/>.</summary>
    public static nint AllocByteLength(int byteLength) => Track(SysAllocStringByteLen(0, (uint)byteLength));

    /// <summary>A BSTR with the same bytes; null stays null.</summary>
    public static nint Copy(nint bstr) => bstr == 0 ? 0 : AllocBytes(Bytes(bstr));

    /// <summary>
    /// A BSTR a module keeps for the life of the process: a string literal, which every use views
    /// and no store frees (D20). It is outside <see cref="LiveCount"/>, and a debug build refuses
    /// to free it, since it was never tracked.
    /// </summary>
    public static unsafe nint AllocLiteral(ReadOnlySpan<char> text)
    {
        fixed (char* chars = text)
        {
            var bstr = SysAllocStringLen((nint)chars, (uint)text.Length);
            return bstr == 0 ? throw new VbaException(VbaErrors.OutOfMemory) : bstr;
        }
    }

    /// <summary>Frees a BSTR; null is ignored. A debug build throws on a pointer it did not allocate or already freed.</summary>
    public static void Free(nint bstr)
    {
        if (bstr == 0)
        {
            return;
        }

#if DEBUG
        lock (LiveLock)
        {
            if (!LiveSet.Remove(bstr))
            {
                throw new InvalidOperationException(VbaErrors.Invariant($"BSTR 0x{bstr:X} is not a live allocation of the runtime (freed twice, or never allocated here)."));
            }
        }
#endif
        SysFreeString(bstr);
        Interlocked.Decrement(ref live);
    }

    /// <summary>A BSTR a DLL allocated and left in the runtime's storage (a Declare's ByRef Variant): the runtime owns it from here and frees it as its own (D20).</summary>
    internal static void Adopted(nint bstr)
    {
        if (bstr != 0)
        {
            Track(bstr);
        }
    }

    /// <summary>A BSTR of the runtime's that a DLL freed, or overwrote without freeing as VBA would leak it: it leaves the live set without being freed again.</summary>
    internal static void Abandoned(nint bstr)
    {
        if (bstr == 0)
        {
            return;
        }

#if DEBUG
        lock (LiveLock)
        {
            if (!LiveSet.Remove(bstr))
            {
                return;
            }
        }
#endif
        Interlocked.Decrement(ref live);
    }

    /// <summary>The byte count from the prefix; 0 for a null BSTR.</summary>
    public static unsafe int ByteLength(nint bstr) => bstr == 0 ? 0 : *(int*)(bstr - 4);

    /// <summary>The whole code units: the byte count halved, so a trailing odd byte is not a character (Strings golden: LeftB("hello", 3) reads as "h").</summary>
    public static unsafe ReadOnlySpan<char> Chars(nint bstr) => bstr == 0 ? default : new ReadOnlySpan<char>((void*)bstr, ByteLength(bstr) / 2);

    public static unsafe ReadOnlySpan<byte> Bytes(nint bstr) => bstr == 0 ? default : new ReadOnlySpan<byte>((void*)bstr, ByteLength(bstr));

    /// <summary>The code units as a writable span, for the Mid statement and the byte functions that write in place.</summary>
    public static unsafe Span<char> MutableChars(nint bstr) => bstr == 0 ? default : new Span<char>((void*)bstr, ByteLength(bstr) / 2);

    public static unsafe Span<byte> MutableBytes(nint bstr) => bstr == 0 ? default : new Span<byte>((void*)bstr, ByteLength(bstr));

    /// <summary>
    /// Builds a BSTR piece by piece in a pooled buffer and allocates it once at the end, for the
    /// functions that assemble a result (Replace, Join) without a .NET string on the way (R20).
    /// </summary>
    public ref struct Builder
    {
        private char[] buffer;
        private int length;

        public Builder(int capacity)
        {
            buffer = System.Buffers.ArrayPool<char>.Shared.Rent(Math.Max(capacity, 16));
            length = 0;
        }

        public void Append(ReadOnlySpan<char> text)
        {
            if (length + text.Length > buffer.Length)
            {
                var larger = System.Buffers.ArrayPool<char>.Shared.Rent(Math.Max(buffer.Length * 2, length + text.Length));
                buffer.AsSpan(0, length).CopyTo(larger);
                System.Buffers.ArrayPool<char>.Shared.Return(buffer);
                buffer = larger;
            }

            text.CopyTo(buffer.AsSpan(length));
            length += text.Length;
        }

        /// <summary>The BSTR with everything appended, owned by the caller; the buffer goes back to the pool.</summary>
        public VbaString ToVbaString()
        {
            var result = VbaString.Alloc(buffer.AsSpan(0, length));
            System.Buffers.ArrayPool<char>.Shared.Return(buffer);
            buffer = [];
            length = 0;
            return result;
        }
    }

    private static nint Track(nint bstr)
    {
        if (bstr == 0)
        {
            throw new VbaException(VbaErrors.OutOfMemory);
        }

#if DEBUG
        lock (LiveLock)
        {
            LiveSet.Add(bstr);
        }
#endif
        Interlocked.Increment(ref live);
        return bstr;
    }

    [LibraryImport("oleaut32.dll")]
    private static partial nint SysAllocStringLen(nint text, uint length);

    [LibraryImport("oleaut32.dll")]
    private static partial nint SysAllocStringByteLen(nint text, uint byteLength);

    [LibraryImport("oleaut32.dll")]
    private static partial void SysFreeString(nint bstr);
}
