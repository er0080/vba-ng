using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>
/// A Boolean member of a user-defined type as VBA lays it out: two bytes holding VARIANT_BOOL's
/// value, True -1 and False 0 (ARCHITECTURE.md D20; ROADMAP.md M7 C). The bits stay as they are,
/// so a member that memory wrote another value into keeps it: it reads as True and converts to that
/// number (Memory golden: CInt of a member holding 1 is 1). It reads as a Boolean wherever one is
/// wanted.
/// </summary>
[StructLayout(LayoutKind.Sequential, Size = 2)]
[DebuggerDisplay("{DebugView,nq}")]
public readonly struct VbaBoolean
{
    private readonly short bits;

    private VbaBoolean(short bits) => this.bits = bits;

    /// <summary>The two bytes as a number: -1 for True, 0 for False, and whatever else memory put there.</summary>
    public short Bits => bits;

    internal string DebugView => bits switch
    {
        -1 => "True",
        0 => "False",
        _ => "True (" + bits.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")",
    };

    public static implicit operator VbaBoolean(bool value) => FromBoolean(value);

    public static implicit operator bool(VbaBoolean value) => value.ToBoolean();

    public static VbaBoolean FromBoolean(bool value) => new(value ? (short)-1 : (short)0);

    public bool ToBoolean() => bits != 0;

    public override string ToString() => DebugView;
}
