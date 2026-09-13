using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>
/// Converts between <see cref="VbaArray"/> and SAFEARRAY in both directions (ARCHITECTURE.md
/// section 6, "Arrays"): the bounds of every dimension survive, typed arrays keep their element
/// type, and Variant arrays carry each element as its own VARIANT.
/// </summary>
internal static unsafe class SafeArrayMarshal
{
    private const ushort FadfStatic = 0x0002;
    private const ushort FadfFixedSize = 0x0010;

    public static ushort ElementVt(VarType elementType) => elementType switch
    {
        VarType.Integer => Vt.I2,
        VarType.Long => Vt.I4,
        VarType.LongLong => Vt.I8,
        VarType.Single => Vt.R4,
        VarType.Double => Vt.R8,
        VarType.Currency => Vt.Cy,
        VarType.Date => Vt.Date,
        VarType.String => Vt.Bstr,
        VarType.Boolean => Vt.Bool,
        VarType.Byte => Vt.UI1,
        VarType.Object => Vt.Dispatch,
        _ => Vt.Variant,
    };

    public static VarType ElementType(ushort vt) => vt switch
    {
        Vt.I2 => VarType.Integer,
        Vt.I4 or Vt.Int => VarType.Long,
        Vt.I8 => VarType.LongLong,
        Vt.R4 => VarType.Single,
        Vt.R8 => VarType.Double,
        Vt.Cy => VarType.Currency,
        Vt.Date => VarType.Date,
        Vt.Bstr => VarType.String,
        Vt.Bool => VarType.Boolean,
        Vt.UI1 => VarType.Byte,
        Vt.Dispatch or Vt.Unknown => VarType.Object,
        _ => VarType.Variant,
    };

    /// <summary>
    /// The array for a COM call: oleaut32's copy of its SAFEARRAY, since the array is one already
    /// (M7 C4), its strings copied byte for byte, its references AddRef'd, and its Variants copied,
    /// without the fixed-size flags, so the callee owns and may destroy it; null for an unallocated
    /// array, as VBA passes one.
    /// </summary>
    public static nint ToNative(VbaArray array)
    {
        if (!array.IsAllocated)
        {
            return 0;
        }

        Check(NativeMethods.SafeArrayCopy(array.Descriptor, out var copy));
        *(ushort*)(copy + 2) &= unchecked((ushort)~(FadfStatic | FadfFixedSize));
        return copy;
    }

    /// <summary>A VbaArray copied from a SAFEARRAY the caller still owns.</summary>
    public static VbaArray FromNative(nint safeArray, ushort declaredVt)
    {
        var rank = (int)NativeMethods.SafeArrayGetDim(safeArray);
        var vt = NativeMethods.SafeArrayGetVartype(safeArray, out var actualVt) == NativeMethods.SOk ? actualVt : declaredVt;
        var bounds = new (int Lower, int Upper)[rank];
        var counts = new int[rank];
        var total = 1;
        for (var dimension = 0; dimension < rank; dimension++)
        {
            Check(NativeMethods.SafeArrayGetLBound(safeArray, (uint)(dimension + 1), out var lower));
            Check(NativeMethods.SafeArrayGetUBound(safeArray, (uint)(dimension + 1), out var upper));
            bounds[dimension] = (lower, upper);
            counts[dimension] = upper - lower + 1;
            total *= counts[dimension];
        }

        if (total == 0 && rank == 1)
        {
            // An array with no elements, as Scripting.Dictionary's Keys of an empty dictionary: lower To lower - 1, which Create
            // would refuse as ReDim does (Objects golden: UBound -1, bounds 0 and -1, For Each runs no times).
            return VbaArray.Empty(ElementType(vt), bounds[0].Lower);
        }

        var array = VbaArray.Create(ElementType(vt), bounds);
        if (total == 0)
        {
            return array;
        }

        var elementSize = (int)NativeMethods.SafeArrayGetElemsize(safeArray);
        Check(NativeMethods.SafeArrayAccessData(safeArray, out var data));
        try
        {
            var indices = new int[rank];
            for (var position = 0; position < total; position++)
            {
                var remainder = position;
                for (var dimension = 0; dimension < rank; dimension++)
                {
                    indices[dimension] = bounds[dimension].Lower + (remainder % counts[dimension]);
                    remainder /= counts[dimension];
                }

                array.Set(indices, VariantMarshal.FromPayload(vt, (byte*)data + (long)position * elementSize));
            }
        }
        finally
        {
            _ = NativeMethods.SafeArrayUnaccessData(safeArray);
        }

        return array;
    }

    private static void Check(int hr)
    {
        if (hr < 0)
        {
            throw ComErrors.FromHResult(hr);
        }
    }
}
