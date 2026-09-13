using System.Runtime.InteropServices;

namespace VbaNg.Interop;

/// <summary>OLE Automation entry points and the structures the invoke path passes to them (D7: no C# dynamic, no built-in IDispatch binder).</summary>
internal static unsafe partial class NativeMethods
{
    public const int LocaleUserDefault = 0x400;
    public const uint ClsctxServer = 0x1 | 0x4; // CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER

    public const int DispidValue = 0;
    public const int DispidNewEnum = -4;
    public const int DispidPropertyPut = -3;
    public const int DispidUnknown = -1;

    public const int SOk = 0;
    public const int SFalse = 1;
    public const int DispEUnknownName = unchecked((int)0x80020006);
    public const int DispEException = unchecked((int)0x80020009);
    public const int DispEMemberNotFound = unchecked((int)0x80020003);
    public const int DispEParamNotFound = unchecked((int)0x80020004);
    public const int DispETypeMismatch = unchecked((int)0x80020005);
    public const int DispENoNamedArgs = unchecked((int)0x80020007);
    public const int DispEBadVarType = unchecked((int)0x80020008);
    public const int DispEOverflow = unchecked((int)0x8002000A);
    public const int DispEBadIndex = unchecked((int)0x8002000B);
    public const int DispEUnknownLcid = unchecked((int)0x8002000C);
    public const int DispEArrayIsLocked = unchecked((int)0x8002000D);
    public const int DispEBadParamCount = unchecked((int)0x8002000E);
    public const int DispEParamNotOptional = unchecked((int)0x8002000F);
    public const int DispENotACollection = unchecked((int)0x80020011);
    public const int EInvalidArg = unchecked((int)0x80070057);
    public const int EOutOfMemory = unchecked((int)0x8007000E);
    public const int ENoInterface = unchecked((int)0x80004002);

    public static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
    public static readonly Guid IidIDispatch = new("00020400-0000-0000-C000-000000000046");
    public static readonly Guid IidIEnumVariant = new("00020404-0000-0000-C000-000000000046");
    public static readonly Guid IidIProvideClassInfo = new("B196B283-BAB4-101A-B69C-00AA00341D07");
    public static readonly Guid IidNull = Guid.Empty;

    [LibraryImport("ole32.dll", StringMarshalling = StringMarshalling.Utf16)]
    public static partial int CLSIDFromProgID(string progId, out Guid clsid);

    [LibraryImport("ole32.dll")]
    public static partial int CoCreateInstance(in Guid clsid, nint outer, uint context, in Guid iid, out nint result);

    [LibraryImport("ole32.dll", StringMarshalling = StringMarshalling.Utf16)]
    public static partial int CoGetObject(string name, nint bindOptions, in Guid iid, out nint result);

    [LibraryImport("oleaut32.dll")]
    public static partial int GetActiveObject(in Guid clsid, nint reserved, out nint unknown);

    [LibraryImport("oleaut32.dll")]
    public static partial nint SysAllocStringLen(char* text, uint length);

    [LibraryImport("oleaut32.dll")]
    public static partial void SysFreeString(nint bstr);

    [LibraryImport("oleaut32.dll")]
    public static partial uint SysStringLen(nint bstr);

    [LibraryImport("oleaut32.dll")]
    public static partial nint SysAllocStringByteLen(byte* bytes, uint length);

    [LibraryImport("oleaut32.dll")]
    public static partial uint SysStringByteLen(nint bstr);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayCopy(nint array, out nint copy);

    [LibraryImport("oleaut32.dll")]
    public static partial int VariantClear(VARIANT* variant);

    [LibraryImport("oleaut32.dll")]
    public static partial nint SafeArrayCreate(ushort vt, uint dimensions, SAFEARRAYBOUND* bounds);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayDestroy(nint array);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayAccessData(nint array, out nint data);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayUnaccessData(nint array);

    [LibraryImport("oleaut32.dll")]
    public static partial uint SafeArrayGetDim(nint array);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayGetLBound(nint array, uint dimension, out int bound);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayGetUBound(nint array, uint dimension, out int bound);

    [LibraryImport("oleaut32.dll")]
    public static partial int SafeArrayGetVartype(nint array, out ushort vt);

    [LibraryImport("oleaut32.dll")]
    public static partial uint SafeArrayGetElemsize(nint array);

    /// <summary>The <paramref name="index"/>th entry of a COM object's vtable.</summary>
    public static nint Slot(nint comObject, int index) => (*(nint**)comObject)[index];

    /// <summary>IUnknown::AddRef through the vtable.</summary>
    public static void AddRef(nint comObject) => ((delegate* unmanaged[Stdcall]<nint, uint>)Slot(comObject, 1))(comObject);

    /// <summary>IUnknown::Release through the vtable.</summary>
    public static void Release(nint comObject) => ((delegate* unmanaged[Stdcall]<nint, uint>)Slot(comObject, 2))(comObject);

    /// <summary>IUnknown::QueryInterface; the result is null when the interface is not supported.</summary>
    public static nint QueryInterface(nint comObject, in Guid iid)
    {
        fixed (Guid* id = &iid)
        {
            nint result;
            var hr = ((delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)Slot(comObject, 0))(comObject, id, &result);
            return hr == SOk ? result : 0;
        }
    }
}

/// <summary>The OLE VARIANT (24 bytes on x64): the type tag, then a union that also serves as the DECIMAL payload from offset 0.</summary>
[StructLayout(LayoutKind.Explicit, Size = 24)]
internal struct VARIANT
{
    [FieldOffset(0)]
    public ushort vt;

    [FieldOffset(8)]
    public long data;

    [FieldOffset(8)]
    public nint pointer;
}

[StructLayout(LayoutKind.Sequential)]
internal struct DISPPARAMS
{
    public nint rgvarg;
    public nint rgdispidNamedArgs;
    public uint cArgs;
    public uint cNamedArgs;
}

[StructLayout(LayoutKind.Explicit, Size = 64)]
internal struct EXCEPINFO
{
    [FieldOffset(0)]
    public ushort wCode;

    [FieldOffset(8)]
    public nint bstrSource;

    [FieldOffset(16)]
    public nint bstrDescription;

    [FieldOffset(24)]
    public nint bstrHelpFile;

    [FieldOffset(32)]
    public uint dwHelpContext;

    [FieldOffset(48)]
    public nint pfnDeferredFillIn;

    [FieldOffset(56)]
    public int scode;
}

[StructLayout(LayoutKind.Sequential)]
internal struct SAFEARRAYBOUND
{
    public uint cElements;
    public int lLbound;
}

/// <summary>The VARENUM type codes the marshaler handles.</summary>
internal static class Vt
{
    public const ushort Empty = 0;
    public const ushort Null = 1;
    public const ushort I2 = 2;
    public const ushort I4 = 3;
    public const ushort R4 = 4;
    public const ushort R8 = 5;
    public const ushort Cy = 6;
    public const ushort Date = 7;
    public const ushort Bstr = 8;
    public const ushort Dispatch = 9;
    public const ushort Error = 10;
    public const ushort Bool = 11;
    public const ushort Variant = 12;
    public const ushort Unknown = 13;
    public const ushort Decimal = 14;
    public const ushort I1 = 16;
    public const ushort UI1 = 17;
    public const ushort UI2 = 18;
    public const ushort UI4 = 19;
    public const ushort I8 = 20;
    public const ushort UI8 = 21;
    public const ushort Int = 22;
    public const ushort UInt = 23;
    public const ushort Array = 0x2000;
    public const ushort ByRef = 0x4000;
    public const ushort TypeMask = 0x0FFF;
}
