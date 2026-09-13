using System.Runtime.InteropServices;

using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>
/// What a <c>Declare</c> needs at run time (MS-VBAL 5.2.3.5, Declares golden): the ANSI buffer a
/// <c>ByVal ... As String</c> parameter travels in, which the callee may write into, and the
/// errors VBA raises when the library or the entry point is not there.
/// </summary>
public static class Native
{
    /// <summary>
    /// A string as a <c>ByVal ... As String</c> parameter reaches a DLL: the text as
    /// null-terminated ANSI in unmanaged memory. VBA copies the buffer back into the variable
    /// afterwards, keeping the string's length, so a callee that writes into it is visible to the
    /// caller (Declares golden: CharUpperBuffA).
    /// </summary>
    public readonly struct AnsiBuffer : IDisposable
    {
        private readonly int length;

        internal AnsiBuffer(string value)
        {
            var bytes = Strings.Ansi.GetBytes(value ?? string.Empty);
            length = bytes.Length;
            Pointer = Marshal.AllocHGlobal(length + 1);
            Marshal.Copy(bytes, 0, Pointer, length);
            Marshal.WriteByte(Pointer, length, 0);
        }

        /// <summary>The address the DLL receives.</summary>
        public nint Pointer { get; }

        /// <summary>The buffer as the caller's string: the same number of characters it had, whatever the callee wrote.</summary>
        public string Read()
        {
            if (length == 0)
            {
                return string.Empty;
            }

            var bytes = new byte[length];
            Marshal.Copy(Pointer, bytes, 0, length);
            return Strings.Ansi.GetString(bytes);
        }

        /// <summary>The buffer as a String the current statement owns, for the ByVal String parameter it came from (D20).</summary>
        public VbaString ReadText() => VbaString.Temporary(Read());

        public void Dispose() => Marshal.FreeHGlobal(Pointer);
    }

    public static AnsiBuffer Ansi(string value) => new(value);

    /// <summary>
    /// A String member of the copy of a record handed to a Declare: its BSTR becomes one holding
    /// the text's ANSI bytes, as VBA converts it on the way, so the callee never sees the
    /// variable's own string (Memory golden). vbNullString stays null.
    /// </summary>
    public static void ToAnsi(ref VbaString member)
    {
        if (member.Pointer == 0)
        {
            return;
        }

        var ansi = VbaString.AllocBytes(Strings.Ansi.GetBytes(member.ToString()));
        var old = member;
        member = ansi;
        old.Free();
    }

    /// <summary>After the call, the member's ANSI bytes become text again in a new BSTR; the ANSI one lives until the statement ends, so no String the copy back makes can take its place.</summary>
    public static void FromAnsi(ref VbaString member)
    {
        if (member.Pointer == 0)
        {
            return;
        }

        var text = VbaString.Alloc(Strings.Ansi.GetString(member.Bytes));
        ObjectRefs.OwnedString(member.Pointer);
        member = text;
    }

    public static AnsiBuffer Ansi(VbaString value) => new(value.ToString());

    /// <summary>Error 53 when the library is not on the machine (Declares golden).</summary>
    public static VbaException LibraryNotFound(string library) =>
        new(53, "File not found: " + library);

    /// <summary>Error 453 when the library has no such entry point (Declares golden).</summary>
    public static VbaException EntryPointNotFound(string entryPoint, string library) =>
        new(453, $"Can't find DLL entry point {entryPoint} in {library}");

    /// <summary>
    /// After a Declare call, what the DLL left in a ByRef Variant becomes the runtime's, as VBA
    /// takes it (Declares golden: VariantChangeType writes a String into one). A BSTR the DLL
    /// allocated joins the live allocations, and the one it replaced leaves them, freed by the
    /// DLL or, when the DLL overwrote it without freeing, leaked as VBA leaks it (D20). An array
    /// a DLL creates in a Variant is not adopted yet (ROADMAP.md backlog).
    /// </summary>
    public static void TakeOver(in Variant before, ref Variant after)
    {
        if (before.StringPointer != after.StringPointer)
        {
            Bstr.Abandoned(before.StringPointer);
            Bstr.Adopted(after.StringPointer);
        }

        if (before.ArrayDescriptor != after.ArrayDescriptor)
        {
            VbaArray.Abandoned(before.ArrayDescriptor);
            if (after.IsArray)
            {
                AdoptArray(after.AsArray());
            }
        }
    }

    /// <summary>After a Declare call, an array passed ByRef whose descriptor the DLL replaced: the new one becomes the runtime's, and the old one leaves the live arrays, destroyed by the DLL or leaked as VBA leaks it (D20).</summary>
    public static void TakeOver(nint before, ref VbaArray after)
    {
        if (before == after.Descriptor)
        {
            return;
        }

        VbaArray.Abandoned(before);
        AdoptArray(after);
    }

    /// <summary>A Declare's array result, a SAFEARRAY the DLL created (Declares golden: SafeArrayCreateVector): the runtime's from here, and a temporary of the statement the caller stores (D20).</summary>
    public static VbaArray ArrayResult(nint descriptor, VarType elementType)
    {
        if (descriptor == 0)
        {
            return VbaArray.Unallocated(elementType);
        }

        var array = VbaArray.View(descriptor, elementType);
        AdoptArray(array);
        return ObjectRefs.Owned(array);
    }

    /// <summary>A Declare's Variant result, the VARIANT the DLL wrote through the hidden result pointer: what it holds becomes the runtime's, a temporary of the statement (D20).</summary>
    public static Variant VariantResult(Variant result)
    {
        AdoptValue(result);
        return ObjectRefs.Adopt(ref result);
    }

    /// <summary>VBE7's rtcCallByName (MS-VBAL 6.1.2.6 CallByName under its runtime name): the runtime answers it, the name a BSTR pointer and the arguments a Variant array (Declares golden).</summary>
    public static Variant CallByName(object? target, nint name, int callType, in VbaArray arguments)
    {
        var values = new Variant[arguments.IsAllocated ? arguments.Count : 0];
        for (var i = 0; i < values.Length; i++)
        {
            values[i] = arguments.ElementAt(i);
        }

        return LateBound.CallByName(Variant.FromObject(target), VbaString.Adopt(name).ToString(), callType, values);
    }

    /// <summary>
    /// An object passed ByVal to a Declare parameter of a type library's interface: the pointer to
    /// that interface, as VBA hands it, queried from the object and released when the statement
    /// ends (Declares golden: IPicture). Nothing is a null pointer; an object without the
    /// interface is a type mismatch.
    /// </summary>
    public static nint InterfacePointer(object? value, in Guid iid)
    {
        var pointer = Variant.FromObject(value).InterfacePointer;
        if (pointer == 0)
        {
            return 0;
        }

        if (Marshal.QueryInterface(pointer, in iid, out var result) < 0 || result == 0)
        {
            throw VbaErrors.TypeMismatch();
        }

        ObjectRefs.OwnedInterface(result);
        return result;
    }

    /// <summary>A String passed ByRef to a Declare: a BSTR of its own holding the text's ANSI bytes, whose address the DLL receives, as VBA converts it on the way (Declares golden). vbNullString stays null.</summary>
    public static nint AnsiCopy(VbaString value) => value.Pointer == 0 ? 0 : VbaString.AllocBytes(Strings.Ansi.GetBytes(value.ToString())).Pointer;

    /// <summary>After the call, the BSTR the DLL left (the copy, or one it put in the copy's place) becomes the variable's text again, its bytes read as ANSI (Declares golden).</summary>
    public static void FromAnsiCopy(nint original, nint current, ref VbaString target)
    {
        if (current != original)
        {
            Bstr.Abandoned(original);
            Bstr.Adopted(current);
        }

        var text = current == 0 ? VbaString.Null : VbaString.Alloc(Strings.Ansi.GetString(Bstr.Bytes(current)));
        Bstr.Free(current);
        var old = target;
        target = text;
        old.Free();
    }

    /// <summary>An array a DLL made, and the strings and arrays its elements hold, become the runtime's (D20).</summary>
    internal static void AdoptArray(VbaArray array)
    {
        VbaArray.Adopted(array.Descriptor);
        array.AdoptElements();
    }

    /// <summary>A value a DLL made: its BSTR or its array becomes the runtime's; an object came with its reference.</summary>
    internal static void AdoptValue(in Variant value)
    {
        Bstr.Adopted(value.StringPointer);
        if (value.IsArray)
        {
            AdoptArray(value.AsArray());
        }
    }
}
