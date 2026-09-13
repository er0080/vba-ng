using System.Runtime.InteropServices;

using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>
/// Converts between the runtime's <see cref="Variant"/> and the OLE VARIANT, by hand (D7): every
/// VBA type maps to its VARENUM code, Currency to VT_CY, Date to VT_DATE, Missing to VT_ERROR
/// with DISP_E_PARAMNOTFOUND, Nothing to a null VT_DISPATCH, and arrays to SAFEARRAYs.
/// </summary>
internal static unsafe class VariantMarshal
{
    /// <summary>Writes a Variant into a VARIANT the caller owns; <see cref="Clear"/> releases what this allocated.</summary>
    public static void ToNative(in Variant value, VARIANT* target)
    {
        target->vt = Vt.Empty;
        target->data = 0;
        switch (value.Type)
        {
            case VarType.Empty:
                break;
            case VarType.Null:
                target->vt = Vt.Null;
                break;
            case VarType.Integer:
                target->vt = Vt.I2;
                *(short*)&target->data = value.AsInt16();
                break;
            case VarType.Long:
                target->vt = Vt.I4;
                *(int*)&target->data = value.AsInt32();
                break;
            case VarType.LongLong:
                target->vt = Vt.I8;
                target->data = value.AsInt64();
                break;
            case VarType.Single:
                target->vt = Vt.R4;
                *(float*)&target->data = value.AsSingle();
                break;
            case VarType.Double:
                target->vt = Vt.R8;
                *(double*)&target->data = value.AsDouble();
                break;
            case VarType.Currency:
                target->vt = Vt.Cy;
                target->data = value.AsCurrency().Scaled;
                break;
            case VarType.Date:
                target->vt = Vt.Date;
                *(double*)&target->data = value.AsDate().Serial;
                break;
            case VarType.String:
                // The String's own BSTR, byte for byte and a null one as null, as VBA hands a String to COM (D20); the callee owns the copy.
                target->vt = Vt.Bstr;
                target->pointer = CopyBstr(value.AsVbaString().Pointer);
                break;
            case VarType.Boolean:
                target->vt = Vt.Bool;
                *(short*)&target->data = value.AsBoolean() ? (short)-1 : (short)0;
                break;
            case VarType.Byte:
                target->vt = Vt.UI1;
                *(byte*)&target->data = value.AsByte();
                break;
            case VarType.Decimal:
                WriteDecimal(target, value.AsDecimal());
                break;
            case VarType.Error:
                target->vt = Vt.Error;
                *(int*)&target->data = value.AsError().Scode;
                break;
            case VarType.Object:
                target->vt = Vt.Dispatch;
                target->pointer = AddRefInterface(value);
                break;
            case VarType.Array:
                {
                    // The array's descriptor already carries its element type, which is the VARENUM code (M7 C4).
                    var array = value.AsArray();
                    target->vt = (ushort)(Vt.Array | (ushort)array.ElementType);
                    target->pointer = SafeArrayMarshal.ToNative(array);
                    break;
                }

            default:
                throw VbaErrors.TypeMismatch();
        }
    }

    /// <summary>
    /// The interface pointer a VARIANT, a ByRef slot, or a SAFEARRAY element holds for an object,
    /// with a reference of its own for the callee to release: a COM object's IDispatch, or a
    /// runtime object's own, since a class instance, a Collection, Err, or an enumerator crosses as
    /// itself (ARCHITECTURE.md section 5, "Class modules"); 0 for Nothing or Empty.
    /// </summary>
    public static nint AddRefInterface(in Variant value)
    {
        if (value.IsEmpty || value.IsNothing)
        {
            return 0;
        }

        if (!value.IsObject)
        {
            throw VbaErrors.TypeMismatch();
        }

        // The Variant holds the object as its IDispatch pointer (D20): the callee gets that pointer with a reference of its own.
        var pointer = value.InterfacePointer;
        NativeMethods.AddRef(pointer);
        return pointer;
    }

    /// <summary>Reads a VARIANT, following VT_BYREF, into a Variant; COM objects come back wrapped, arrays copied.</summary>
    public static Variant FromNative(VARIANT* source)
    {
        var vt = source->vt;
        var type = (ushort)(vt & Vt.TypeMask);
        if ((vt & Vt.ByRef) != 0)
        {
            var reference = source->pointer;
            if (reference == 0)
            {
                return Variant.Empty;
            }

            if (type == Vt.Variant)
            {
                return FromNative((VARIANT*)reference);
            }

            if ((vt & Vt.Array) != 0)
            {
                return FromArray(*(nint*)reference, type);
            }

            return FromPayload(type, (byte*)reference);
        }

        if ((vt & Vt.Array) != 0)
        {
            return FromArray(source->pointer, type);
        }

        return type switch
        {
            Vt.Empty => Variant.Empty,
            Vt.Null => Variant.Null,
            Vt.Decimal => Variant.FromDecimal(ReadDecimal(source)),
            _ => FromPayload(type, (byte*)&source->data),
        };
    }

    /// <summary>A value of a known VARENUM type at an address: a VARIANT's payload or a SAFEARRAY element.</summary>
    public static Variant FromPayload(ushort type, byte* data)
    {
        switch (type)
        {
            case Vt.Empty:
                return Variant.Empty;
            case Vt.Null:
                return Variant.Null;
            case Vt.I2:
                return Variant.FromInt16(*(short*)data);
            case Vt.I4:
            case Vt.Int:
                return Variant.FromInt32(*(int*)data);
            case Vt.R4:
                return Variant.FromSingle(*(float*)data);
            case Vt.R8:
                return Variant.FromDouble(*(double*)data);
            case Vt.Cy:
                return Variant.FromCurrency(Currency.FromScaled(*(long*)data));
            case Vt.Date:
                return Variant.FromDate(VbaDate.FromSerial(*(double*)data));
            case Vt.Bstr:
                return FromBstr(*(nint*)data);
            case Vt.Dispatch:
            case Vt.Unknown:
                return FromInterface(*(nint*)data);
            case Vt.Error:
                return Variant.FromError(new ErrorValue(*(int*)data));
            case Vt.Bool:
                return Variant.FromBoolean(*(short*)data != 0);
            case Vt.Variant:
                return FromNative((VARIANT*)data);
            case Vt.Decimal:
                return Variant.FromDecimal(ReadDecimal((VARIANT*)data));
            case Vt.I1:
                return Variant.FromInt16(*(sbyte*)data);
            case Vt.UI1:
                return Variant.FromByte(*data);
            case Vt.UI2:
                return Variant.FromInt32(*(ushort*)data);
            case Vt.UI4:
            case Vt.UInt:
                return Variant.FromInt32(unchecked((int)*(uint*)data));
            case Vt.I8:
                return Variant.FromInt64(*(long*)data);
            case Vt.UI8:
                return Variant.FromInt64(unchecked((long)*(ulong*)data));
            default:
                // A type VBA cannot hold in a Variant (VT_RECORD, VT_PTR, ...): "Variable uses an Automation type not supported in Visual Basic".
                throw new VbaException(458);
        }
    }

    /// <summary>
    /// An object a call returned, as the Variant holds it: its IDispatch pointer, with a reference
    /// that is a temporary of the current statement (ARCHITECTURE.md D18, D20), released when the
    /// statement ends unless a store took one of its own, on the thread that made it, as VBA
    /// releases the objects a call returns. The pointer passed in is not consumed. A runtime
    /// object's own block answers QueryInterface with itself, so a class instance COM kept comes
    /// back as itself. No wrapper is made until code asks for one (<see cref="Variant.AsObject"/>).
    /// </summary>
    public static Variant FromInterface(nint pointer)
    {
        if (pointer == 0)
        {
            return Variant.Nothing;
        }

        var dispatch = NativeMethods.QueryInterface(pointer, in NativeMethods.IidIDispatch);
        if (dispatch == 0)
        {
            throw new VbaException(VbaErrors.ClassDoesNotSupportAutomation);
        }

        return ObjectRefs.OwnedInterface(dispatch);
    }

    public static void Clear(VARIANT* variant) => _ = NativeMethods.VariantClear(variant);

    /// <summary>
    /// Writes a value through a VT_BYREF VARIANT, as an event handler's ByRef parameter reaches
    /// the source (Cancel = True): the pointed-to scalar takes the coerced value, a pointed-to
    /// VARIANT is cleared and rewritten, a BSTR is freed and replaced.
    /// </summary>
    public static void WriteByRef(VARIANT* variant, in Variant value)
    {
        if ((variant->vt & Vt.ByRef) == 0 || variant->pointer == 0)
        {
            return;
        }

        var target = variant->pointer;
        switch ((ushort)(variant->vt & Vt.TypeMask))
        {
            case Vt.Variant:
                Clear((VARIANT*)target);
                ToNative(value, (VARIANT*)target);
                break;
            case Vt.Bool:
                *(short*)target = Coerce.ToBoolean(value) ? (short)-1 : (short)0;
                break;
            case Vt.I2:
                *(short*)target = Coerce.ToInt16(value);
                break;
            case Vt.I4:
            case Vt.Int:
                *(int*)target = Coerce.ToInt32(value);
                break;
            case Vt.I8:
                *(long*)target = Coerce.ToInt64(value);
                break;
            case Vt.R4:
                *(float*)target = Coerce.ToSingle(value);
                break;
            case Vt.R8:
                *(double*)target = Coerce.ToDouble(value);
                break;
            case Vt.Cy:
                *(long*)target = Coerce.ToCurrency(value).Scaled;
                break;
            case Vt.Date:
                *(double*)target = Coerce.ToDate(value).Serial;
                break;
            case Vt.UI1:
                *(byte*)target = Coerce.ToByte(value);
                break;
            case Vt.Bstr:
                {
                    var old = *(nint*)target;
                    if (old != 0)
                    {
                        NativeMethods.SysFreeString(old);
                    }

                    *(nint*)target = value.IsString ? CopyBstr(value.AsVbaString().Pointer) : AllocBstr(Coerce.ToString(value));
                    break;
                }

            case Vt.Dispatch:
            case Vt.Unknown:
                {
                    var old = *(nint*)target;
                    *(nint*)target = AddRefInterface(value);
                    if (old != 0)
                    {
                        NativeMethods.Release(old);
                    }

                    break;
                }

            default:
                break;
        }
    }

    public static nint AllocBstr(string text)
    {
        fixed (char* chars = text)
        {
            var bstr = NativeMethods.SysAllocStringLen(chars, (uint)text.Length);
            return bstr != 0 || text.Length == 0 ? bstr : throw new VbaException(VbaErrors.OutOfMemory);
        }
    }

    /// <summary>A copy of a BSTR for COM, byte for byte (an odd trailing byte included), from oleaut32 so the callee can free it; a null BSTR stays null.</summary>
    public static nint CopyBstr(nint bstr)
    {
        if (bstr == 0)
        {
            return 0;
        }

        var length = NativeMethods.SysStringByteLen(bstr);
        var copy = NativeMethods.SysAllocStringByteLen((byte*)bstr, length);
        return copy != 0 ? copy : throw new VbaException(VbaErrors.OutOfMemory);
    }

    /// <summary>A String a COM call handed back, as a BSTR of the statement's, byte for byte; a null BSTR comes back as vbNullString.</summary>
    public static Variant FromBstr(nint bstr) =>
        bstr == 0 ? Variant.EmptyString : Variant.FromVbaString(VbaString.AllocBytes(new ReadOnlySpan<byte>((void*)bstr, (int)NativeMethods.SysStringByteLen(bstr))));

    public static string ReadBstr(nint bstr) =>
        bstr == 0 ? string.Empty : new string((char*)bstr, 0, (int)NativeMethods.SysStringLen(bstr));

    /// <summary>A SAFEARRAY read into an array of the statement: its elements go when the statement ends unless a store copied them (D18, D20).</summary>
    private static Variant FromArray(nint array, ushort elementType) =>
        array == 0 ? Variant.FromArray(VbaArray.Unallocated(SafeArrayMarshal.ElementType(elementType))) : Variant.FromArray(ObjectRefs.Owned(SafeArrayMarshal.FromNative(array, elementType)));

    /// <summary>The DECIMAL layout shares the VARIANT: wReserved (the vt), scale, sign, Hi32, Lo64.</summary>
    private static decimal ReadDecimal(VARIANT* variant)
    {
        var bytes = (byte*)variant;
        var scale = bytes[2];
        var negative = bytes[3] != 0;
        var hi = *(uint*)(bytes + 4);
        var lo = *(ulong*)(bytes + 8);
        return new decimal((int)(uint)lo, (int)(uint)(lo >> 32), (int)hi, negative, scale);
    }

    private static void WriteDecimal(VARIANT* variant, decimal value)
    {
        Span<int> bits = stackalloc int[4];
        decimal.GetBits(value, bits);
        var bytes = (byte*)variant;
        variant->vt = Vt.Decimal;
        bytes[2] = (byte)((bits[3] >> 16) & 0xFF);
        bytes[3] = (bits[3] & unchecked((int)0x80000000)) != 0 ? (byte)0x80 : (byte)0;
        *(uint*)(bytes + 4) = (uint)bits[2];
        *(ulong*)(bytes + 8) = (uint)bits[0] | ((ulong)(uint)bits[1] << 32);
    }
}
