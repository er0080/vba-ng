using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>
/// A dynamic array member of a user-defined type as VBA lays it out: the eight-byte pointer to its
/// SAFEARRAY descriptor, 0 until the member is first dimensioned (ARCHITECTURE.md D20; ROADMAP.md
/// M7 C8; Memory golden). Generated code reaches the array through a view of the member's declared
/// element type (<see cref="View"/>); a statement that can replace the descriptor (ReDim, Erase, a
/// ByRef call) works on a view and stores it back, and an assignment goes through
/// <see cref="Assign"/>.
/// </summary>
[StructLayout(LayoutKind.Sequential, Size = 8)]
[DebuggerDisplay("{DebugView,nq}")]
public readonly struct VbaArrayMember
{
    private readonly nint descriptor;

    private VbaArrayMember(nint descriptor) => this.descriptor = descriptor;

    /// <summary>The SAFEARRAY; 0 for a member not yet dimensioned.</summary>
    public nint Descriptor => descriptor;

    internal string DebugView => descriptor == 0 ? "(not dimensioned)" : "SAFEARRAY 0x" + descriptor.ToString("X", System.Globalization.CultureInfo.InvariantCulture);

    public static implicit operator VbaArrayMember(VbaArray array) => FromArray(array);

    public static VbaArrayMember FromArray(VbaArray array) => new(array.Descriptor);

    /// <summary>The array, as a view of the member's descriptor with the element type the member is declared with.</summary>
    public VbaArray View(VarType elementType) => VbaArray.View(descriptor, elementType);

    /// <summary>A store into the member: the array it held goes, and it takes <paramref name="value"/>, which is a copy of its own (MS-VBAL 5.4.3.1).</summary>
    public static VbaArrayMember Assign(VbaArrayMember member, VarType elementType, VbaArray value)
    {
        var view = member.View(elementType);
        ObjectRefs.AssignArray(ref view, value);
        return view;
    }

    /// <summary>The member's storage goes: the array and what its elements own are destroyed, and the member is 0 again.</summary>
    public static void Release(ref VbaArrayMember member, VarType elementType)
    {
        var view = member.View(elementType);
        ObjectRefs.Release(ref view);
        member = default;
    }
}
