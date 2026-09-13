using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VbaNg.Interop;

/// <summary>One event of a source interface: its dispid and the names of its parameters.</summary>
public sealed record ComEvent(string Name, int DispId, IReadOnlyList<string> Parameters);

/// <summary>
/// The default source interface of a COM object, found through its class information
/// (ARCHITECTURE.md section 6, "Events"): the IID to advise on and the events by name with their
/// dispids. Works for Excel's Workbook and Worksheet objects and for ActiveX controls alike, so
/// the host never needs a type library model to hook events.
/// </summary>
public sealed class EventSource
{
    private const int MemberIdNil = -1;

    private EventSource(Guid interfaceId, string interfaceName, IReadOnlyDictionary<string, ComEvent> events, bool declaredOrder)
    {
        InterfaceId = interfaceId;
        InterfaceName = interfaceName;
        Events = events;
        DeclaredOrder = declaredOrder;
    }

    /// <summary>
    /// True when the source fires its events with rgvarg in declared order rather than the reversed
    /// order IDispatch prescribes, as Excel does (its WorkbookBeforeClose arrives as Wb, Cancel).
    /// Sources described from a type library model are Office objects and get this; sources
    /// described from their own class information follow the convention.
    /// </summary>
    public bool DeclaredOrder { get; }

    public Guid InterfaceId { get; }

    public string InterfaceName { get; }

    /// <summary>Events by name, case-insensitively.</summary>
    public IReadOnlyDictionary<string, ComEvent> Events { get; }

    /// <summary>
    /// The default source of a coclass from a type library model, for objects that expose no class
    /// information at run time (Excel's own objects): Workbook gives WorkbookEvents, Worksheet gives
    /// DocEvents, Application gives AppEvents. Null when the library has no such coclass or source.
    /// </summary>
    public static EventSource? FromLibrary(VbaNg.Runtime.TypeLibraries.ComLibrary library, string coclassName)
    {
        ArgumentNullException.ThrowIfNull(library);
        ArgumentNullException.ThrowIfNull(coclassName);
        var coclass = library.FindType(coclassName);
        if (coclass?.DefaultSource is not { } sourceName)
        {
            return null;
        }

        var dot = sourceName.IndexOf('.', StringComparison.Ordinal);
        var source = library.FindType(dot < 0 ? sourceName : sourceName[(dot + 1)..]);
        if (source is null)
        {
            return null;
        }

        var events = new Dictionary<string, ComEvent>(StringComparer.OrdinalIgnoreCase);
        foreach (var member in source.Members)
        {
            if (member.Kind == VbaNg.Runtime.TypeLibraries.ComMemberKind.Method)
            {
                events.TryAdd(member.Name, new ComEvent(member.Name, member.DispId, member.Parameters.Select(p => p.Name).ToList()));
            }
        }

        return new EventSource(source.Guid, source.Name, events, declaredOrder: true);
    }

    /// <summary>The default source of an object, or null when it exposes no class information or no source interface.</summary>
    public static EventSource? Describe(ComObject target)
    {
        ArgumentNullException.ThrowIfNull(target);
        var classInfo = NativeMethods.QueryInterface(target.DispatchPointer(), in NativeMethods.IidIProvideClassInfo);
        if (classInfo == 0)
        {
            return null;
        }

        try
        {
            nint typeInfoPointer;
            unsafe
            {
                var hr = ((delegate* unmanaged[Stdcall]<nint, nint*, int>)NativeMethods.Slot(classInfo, 3))(classInfo, &typeInfoPointer);
                if (hr < 0 || typeInfoPointer == 0)
                {
                    return null;
                }
            }

            var coclass = (ITypeInfo)Marshal.GetObjectForIUnknown(typeInfoPointer);
            Marshal.Release(typeInfoPointer);
            try
            {
                return FromCoClass(coclass);
            }
            finally
            {
                Marshal.ReleaseComObject(coclass);
            }
        }
        finally
        {
            NativeMethods.Release(classInfo);
        }
    }

    private static EventSource? FromCoClass(ITypeInfo coclass)
    {
        coclass.GetTypeAttr(out var attrPointer);
        int implCount;
        try
        {
            implCount = Marshal.PtrToStructure<TYPEATTR>(attrPointer).cImplTypes;
        }
        finally
        {
            coclass.ReleaseTypeAttr(attrPointer);
        }

        for (var i = 0; i < implCount; i++)
        {
            coclass.GetImplTypeFlags(i, out var flags);
            if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) == 0 || (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) == 0)
            {
                continue;
            }

            coclass.GetRefTypeOfImplType(i, out var reference);
            coclass.GetRefTypeInfo(reference, out var source);
            try
            {
                return FromSource(source);
            }
            finally
            {
                Marshal.ReleaseComObject(source);
            }
        }

        return null;
    }

    private static EventSource FromSource(ITypeInfo source)
    {
        source.GetDocumentation(MemberIdNil, out var name, out _, out _, out _);
        source.GetTypeAttr(out var attrPointer);
        Guid iid;
        int functionCount;
        try
        {
            var attr = Marshal.PtrToStructure<TYPEATTR>(attrPointer);
            iid = attr.guid;
            functionCount = attr.cFuncs;
        }
        finally
        {
            source.ReleaseTypeAttr(attrPointer);
        }

        var events = new Dictionary<string, ComEvent>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < functionCount; i++)
        {
            source.GetFuncDesc(i, out var descPointer);
            try
            {
                var desc = Marshal.PtrToStructure<FUNCDESC>(descPointer);
                var names = new string[desc.cParams + 1];
                source.GetNames(desc.memid, names, names.Length, out var count);
                if (count == 0)
                {
                    continue;
                }

                events.TryAdd(names[0], new ComEvent(names[0], desc.memid, names.Skip(1).Take(count - 1).ToList()));
            }
            finally
            {
                source.ReleaseFuncDesc(descPointer);
            }
        }

        return new EventSource(iid, name, events, declaredOrder: false);
    }
}
