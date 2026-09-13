using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>The non-generic face of an <see cref="ObjectSlot{T}"/>, for a host calling a procedure by reflection: the slot of a ByRef object parameter (a UDF's Range, an event's object).</summary>
public interface IObjectSlot
{
    /// <summary>The interface pointer the slot holds a reference on; 0 for Nothing.</summary>
    nint Pointer { get; }

    /// <summary>Set slot = value: the new object gains a reference before the old one loses its own.</summary>
    void Assign(object? value);
}

/// <summary>
/// A typed object variable's storage (ARCHITECTURE.md section 5, D20; ROADMAP.md M7 E3): the
/// interface pointer itself, one machine word as in VBA, owning one reference on the object.
/// Locals, parameters, module-level and class-level variables, Function results, and members of
/// user-defined types of an object type are slots, so VarPtr reports where the pointer lies and
/// CopyMemory into the variable writes the pointer without AddRef, as it does in VBA (Memory
/// golden). Generated code reads the object through <see cref="Target"/> and stores and releases
/// through <see cref="ObjectRefs"/>.
/// </summary>
/// <typeparam name="T">The declared class: a project class, a runtime class, a COM type's wrapper, or object for As Object.</typeparam>
[DebuggerDisplay("{DebugView,nq}")]
[StructLayout(LayoutKind.Sequential)]
public struct ObjectSlot<T> : IObjectSlot
    where T : class
{
    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private nint pointer;

    public readonly nint Pointer => pointer;

    /// <summary>
    /// The object the pointer is, borrowing the slot's reference: a runtime object itself, a COM
    /// object through a view; null for Nothing. A pointer to an object of another class, which
    /// only CopyMemory can put there, is a type mismatch.
    /// </summary>
    public readonly T? Target => pointer == 0 ? null : Variant.ObjectOf(pointer) as T ?? throw VbaErrors.TypeMismatch();

    /// <summary>What the locals window shows (ARCHITECTURE.md section 9): Nothing, or the object's TypeName and its pointer, as VBA's shows the class.</summary>
    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly string DebugView => pointer == 0 ? "Nothing" : VbaErrors.Invariant($"{Library.Information.ObjectTypeName(Variant.ObjectOf(pointer)!)} 0x{pointer:X}");

    void IObjectSlot.Assign(object? value) => ObjectRefs.Assign(ref this, (T?)value);

    /// <summary>Puts a pointer in the slot and hands back the one it held: the references move, no count changes.</summary>
    internal nint Exchange(nint value)
    {
        var old = pointer;
        pointer = value;
        return old;
    }
}
