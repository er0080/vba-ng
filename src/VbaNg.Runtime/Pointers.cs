using System.Runtime.CompilerServices;

namespace VbaNg.Runtime;

/// <summary>
/// VarPtr and ObjPtr (ARCHITECTURE.md section 5, "Pointers"; ROADMAP.md M7 E1): the address of a
/// variable's storage and the interface pointer of an object. Generated code hands VarPtr the
/// variable by reference, and the compiler lets that through only for storage whose address is
/// stable for the variable's lifetime (a local or a parameter on the stack, a module-level
/// variable that holds no managed reference), so the address means what it means in VBA; StrPtr
/// is the String's own BSTR pointer. What CopyMemory then does with an address is VBA's
/// business, a wrong copy included (D20).
/// </summary>
public static unsafe class Pointers
{
    /// <summary>VarPtr of a variable: the address of its storage.</summary>
    public static long Address<T>(ref T variable) => (long)Unsafe.AsPointer(ref variable);

    /// <summary>The byte at an address, by reference: what a Declare's ByRef As Any parameter receives when the call passes a pointer ByVal, as VBA hands the DLL that pointer (MS-VBAL 5.2.3.5).</summary>
    public static ref byte At(long address) => ref Unsafe.AsRef<byte>((void*)address);

    /// <summary>ObjPtr: the interface pointer an object reference is, 0 for Nothing; anything but an object is a type mismatch.</summary>
    public static long ObjPtr(in Variant value) => value.IsObject ? value.InterfacePointer : throw VbaErrors.TypeMismatch();
}
