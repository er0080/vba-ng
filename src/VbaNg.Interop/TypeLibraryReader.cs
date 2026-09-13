using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Interop;

/// <summary>
/// Reads a registered type library into the compact <see cref="ComLibrary"/> model
/// (ARCHITECTURE.md section 6, "Type library reader") with no Excel running: coclasses with their
/// default and source interfaces, interfaces with members, dispids, parameter types and defaults,
/// enums and module constants, aliases. Dual interfaces are read from their dispatch view, so a
/// member's type is the value VBA sees, never the HRESULT of the vtable view.
/// </summary>
public static unsafe partial class TypeLibraryReader
{
    private const int MemberIdNil = -1;
    private const int RegKindNone = 2;

    [LibraryImport("oleaut32.dll")]
    private static partial int LoadRegTypeLib(in Guid guid, ushort majorVersion, ushort minorVersion, int lcid, out nint typeLib);

    [LibraryImport("oleaut32.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int LoadTypeLibEx(string path, int registerKind, out nint typeLib);

    /// <summary>Reads a library registered under a guid and version; throws when it is not registered.</summary>
    public static ComLibrary Read(Guid libraryId, int majorVersion, int minorVersion)
    {
        var hr = LoadRegTypeLib(in libraryId, (ushort)majorVersion, (ushort)minorVersion, 0, out var pointer);
        if (hr < 0 || pointer == 0)
        {
            throw new FileNotFoundException(string.Create(CultureInfo.InvariantCulture, $"Type library {libraryId:B} version {majorVersion}.{minorVersion} is not registered (HRESULT 0x{hr:X8})."));
        }

        return ReadPointer(pointer);
    }

    /// <summary>Reads a type library from a file (.tlb, .olb, .dll, .exe) without registering it.</summary>
    public static ComLibrary ReadFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var hr = LoadTypeLibEx(path, RegKindNone, out var pointer);
        if (hr < 0 || pointer == 0)
        {
            throw new FileNotFoundException(string.Create(CultureInfo.InvariantCulture, $"Type library {path} could not be loaded (HRESULT 0x{hr:X8})."), path);
        }

        return ReadPointer(pointer);
    }

    private static ComLibrary ReadPointer(nint pointer)
    {
        var library = (ITypeLib)Marshal.GetObjectForIUnknown(pointer);
        Marshal.Release(pointer);
        try
        {
            return ReadLibrary(library);
        }
        finally
        {
            Marshal.ReleaseComObject(library);
        }
    }

    private static ComLibrary ReadLibrary(ITypeLib library)
    {
        library.GetDocumentation(MemberIdNil, out var name, out _, out _, out _);
        library.GetLibAttr(out var attrPointer);
        TYPELIBATTR attr;
        try
        {
            attr = Marshal.PtrToStructure<TYPELIBATTR>(attrPointer);
        }
        finally
        {
            library.ReleaseTLibAttr(attrPointer);
        }

        var types = new List<ComType>();
        var count = library.GetTypeInfoCount();
        for (var i = 0; i < count; i++)
        {
            library.GetTypeInfo(i, out var info);
            try
            {
                types.Add(ReadType(info));
            }
            finally
            {
                Marshal.ReleaseComObject(info);
            }
        }

        return new ComLibrary(name, attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, types) { Format = ComLibrary.CurrentFormat };
    }

    private static ComType ReadType(ITypeInfo info)
    {
        info.GetDocumentation(MemberIdNil, out var name, out _, out _, out _);
        info.GetTypeAttr(out var attrPointer);
        try
        {
            var attr = Marshal.PtrToStructure<TYPEATTR>(attrPointer);
            var kind = attr.typekind switch
            {
                TYPEKIND.TKIND_ENUM => ComTypeKind.Enum,
                TYPEKIND.TKIND_RECORD => ComTypeKind.Record,
                TYPEKIND.TKIND_MODULE => ComTypeKind.Module,
                TYPEKIND.TKIND_INTERFACE => ComTypeKind.Interface,
                TYPEKIND.TKIND_DISPATCH => ComTypeKind.Dispatch,
                TYPEKIND.TKIND_COCLASS => ComTypeKind.CoClass,
                TYPEKIND.TKIND_ALIAS => ComTypeKind.Alias,
                _ => ComTypeKind.Union,
            };

            var members = new List<ComMember>();
            for (var i = 0; i < attr.cFuncs; i++)
            {
                var member = ReadFunction(info, i);
                if (member is not null)
                {
                    members.Add(member);
                }
            }

            for (var i = 0; i < attr.cVars; i++)
            {
                members.Add(ReadVariable(info, i, kind));
            }

            var implements = new List<ComImplemented>();
            for (var i = 0; i < attr.cImplTypes; i++)
            {
                info.GetRefTypeOfImplType(i, out var reference);
                info.GetImplTypeFlags(i, out var flags);
                implements.Add(new ComImplemented(
                    QualifiedName(info, reference),
                    (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) != 0,
                    (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0,
                    (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FRESTRICTED) != 0));
            }

            var alias = kind == ComTypeKind.Alias ? ReadTypeRef(info, attr.tdescAlias) : null;
            return new ComType(
                name,
                kind,
                attr.guid,
                members,
                implements,
                alias,
                (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FHIDDEN) != 0,
                (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FRESTRICTED) != 0,
                (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FAPPOBJECT) != 0,
                (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FDUAL) != 0,
                // An interface without TYPEFLAG_FNONEXTENSIBLE may have members the library does not list; VBA binds unknown members on it late (docs/vba-quirks.md, MSForms.Control).
                (kind is ComTypeKind.Interface or ComTypeKind.Dispatch) && (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE) == 0);
        }
        finally
        {
            info.ReleaseTypeAttr(attrPointer);
        }
    }

    /// <summary>A FUNCDESC as VBA sees it: on a vtable interface the [retval] parameter is the return type and HRESULT is dropped.</summary>
    private static ComMember? ReadFunction(ITypeInfo info, int index)
    {
        info.GetFuncDesc(index, out var descPointer);
        try
        {
            var desc = Marshal.PtrToStructure<FUNCDESC>(descPointer);
            var names = new string[desc.cParams + 1];
            info.GetNames(desc.memid, names, names.Length, out var nameCount);
            var name = nameCount > 0 ? names[0] : string.Empty;
            var kind = desc.invkind switch
            {
                INVOKEKIND.INVOKE_PROPERTYGET => ComMemberKind.PropertyGet,
                INVOKEKIND.INVOKE_PROPERTYPUT => ComMemberKind.PropertyPut,
                INVOKEKIND.INVOKE_PROPERTYPUTREF => ComMemberKind.PropertyPutRef,
                _ => ComMemberKind.Method,
            };

            var parameters = new List<ComParameter>();
            ComTypeRef? returnType = ReadTypeRef(info, desc.elemdescFunc.tdesc);
            var elementSize = Marshal.SizeOf<ELEMDESC>();
            for (var i = 0; i < desc.cParams; i++)
            {
                var element = Marshal.PtrToStructure<ELEMDESC>(desc.lprgelemdescParam + i * elementSize);
                var flags = element.desc.paramdesc.wParamFlags;
                var type = ReadTypeRef(info, element.tdesc);
                if ((flags & PARAMFLAG.PARAMFLAG_FRETVAL) != 0)
                {
                    returnType = type with { IsPointer = false };
                    continue;
                }

                if ((flags & PARAMFLAG.PARAMFLAG_FLCID) != 0)
                {
                    continue;
                }

                var isOut = (flags & PARAMFLAG.PARAMFLAG_FOUT) != 0;
                var isOptional = (flags & PARAMFLAG.PARAMFLAG_FOPT) != 0;
                var isParamArray = desc.cParamsOpt == -1 && i == desc.cParams - 1;
                ComConstant? defaultValue = null;
                if ((flags & PARAMFLAG.PARAMFLAG_FHASDEFAULT) != 0 && element.desc.paramdesc.lpVarValue != 0)
                {
                    // PARAMDESCEX: a ULONG size, then the VARIANT at offset 8.
                    defaultValue = ReadConstant((VARIANT*)(element.desc.paramdesc.lpVarValue + 8));
                }

                parameters.Add(new ComParameter(
                    i + 1 < nameCount ? names[i + 1] : string.Create(CultureInfo.InvariantCulture, $"arg{i}"),
                    type.IsPointer && !type.IsUserDefined ? type with { IsPointer = false } : type,
                    isOptional || defaultValue is not null,
                    type.IsPointer && (isOut || !IsInterface(info, type)),
                    isOut,
                    isParamArray,
                    defaultValue));
            }

            if (returnType is { VarType: (int)VarEnum.VT_VOID } || returnType is { VarType: (int)VarEnum.VT_HRESULT })
            {
                returnType = null;
            }

            return new ComMember(
                name,
                desc.memid,
                kind,
                returnType,
                parameters,
                ((FUNCFLAGS)desc.wFuncFlags & FUNCFLAGS.FUNCFLAG_FHIDDEN) != 0,
                ((FUNCFLAGS)desc.wFuncFlags & FUNCFLAGS.FUNCFLAG_FRESTRICTED) != 0,
                null);
        }
        finally
        {
            info.ReleaseFuncDesc(descPointer);
        }
    }

    private static ComMember ReadVariable(ITypeInfo info, int index, ComTypeKind owner)
    {
        info.GetVarDesc(index, out var descPointer);
        try
        {
            var desc = Marshal.PtrToStructure<VARDESC>(descPointer);
            var names = new string[1];
            info.GetNames(desc.memid, names, 1, out var nameCount);
            var name = nameCount > 0 ? names[0] : string.Empty;
            var type = ReadTypeRef(info, desc.elemdescVar.tdesc);
            var kind = desc.varkind switch
            {
                VARKIND.VAR_CONST => ComMemberKind.Constant,
                VARKIND.VAR_DISPATCH => ComMemberKind.PropertyGet,
                _ => ComMemberKind.Field,
            };
            var value = desc.varkind == VARKIND.VAR_CONST && desc.desc.lpvarValue != 0 ? ReadConstant((VARIANT*)desc.desc.lpvarValue) : null;
            var hidden = ((VARFLAGS)desc.wVarFlags & VARFLAGS.VARFLAG_FHIDDEN) != 0;
            var restricted = ((VARFLAGS)desc.wVarFlags & VARFLAGS.VARFLAG_FRESTRICTED) != 0;
            return new ComMember(name, desc.memid, kind, type, [], hidden, restricted, value);
        }
        finally
        {
            info.ReleaseVarDesc(descPointer);
        }
    }

    private static ComTypeRef ReadTypeRef(ITypeInfo info, TYPEDESC desc)
    {
        switch ((VarEnum)desc.vt)
        {
            case VarEnum.VT_PTR:
                {
                    var inner = ReadTypeRef(info, Marshal.PtrToStructure<TYPEDESC>(desc.lpValue));
                    return inner with { IsPointer = true };
                }

            case VarEnum.VT_SAFEARRAY:
                {
                    var inner = ReadTypeRef(info, Marshal.PtrToStructure<TYPEDESC>(desc.lpValue));
                    return inner with { IsArray = true };
                }

            case VarEnum.VT_CARRAY:
                {
                    // ARRAYDESC starts with the element TYPEDESC.
                    var inner = ReadTypeRef(info, Marshal.PtrToStructure<TYPEDESC>(desc.lpValue));
                    return inner with { IsArray = true };
                }

            case VarEnum.VT_USERDEFINED:
                return new ComTypeRef(ComTypeRef.UserDefined, QualifiedName(info, (int)desc.lpValue), false, false);

            default:
                return new ComTypeRef(desc.vt, null, false, false);
        }
    }

    /// <summary>True when a pointer type refers to an interface or coclass, which is an object reference rather than a ByRef value.</summary>
    private static bool IsInterface(ITypeInfo info, ComTypeRef type)
    {
        if (!type.IsUserDefined)
        {
            return false;
        }

        // Interfaces and coclasses reached through a pointer are object references; enums, records, and aliases are values.
        return type.TypeName is not null && resolvedKinds.GetValueOrDefault(type.TypeName) is ComTypeKind.Interface or ComTypeKind.Dispatch or ComTypeKind.CoClass;
    }

    [ThreadStatic]
    private static Dictionary<string, ComTypeKind>? kinds;

    private static Dictionary<string, ComTypeKind> resolvedKinds => kinds ??= new Dictionary<string, ComTypeKind>(StringComparer.OrdinalIgnoreCase);

    /// <summary>The name of a referenced type as <c>Library.Type</c>, remembering its kind for pointer classification.</summary>
    private static string QualifiedName(ITypeInfo info, int reference)
    {
        info.GetRefTypeInfo(reference, out var target);
        try
        {
            target.GetDocumentation(MemberIdNil, out var name, out _, out _, out _);
            target.GetContainingTypeLib(out var library, out _);
            string libraryName;
            try
            {
                library.GetDocumentation(MemberIdNil, out libraryName, out _, out _, out _);
            }
            finally
            {
                Marshal.ReleaseComObject(library);
            }

            var qualified = libraryName + "." + name;
            if (!resolvedKinds.ContainsKey(qualified))
            {
                target.GetTypeAttr(out var attrPointer);
                try
                {
                    var attr = Marshal.PtrToStructure<TYPEATTR>(attrPointer);
                    resolvedKinds[qualified] = attr.typekind switch
                    {
                        TYPEKIND.TKIND_INTERFACE => ComTypeKind.Interface,
                        TYPEKIND.TKIND_DISPATCH => ComTypeKind.Dispatch,
                        TYPEKIND.TKIND_COCLASS => ComTypeKind.CoClass,
                        TYPEKIND.TKIND_ENUM => ComTypeKind.Enum,
                        TYPEKIND.TKIND_RECORD => ComTypeKind.Record,
                        TYPEKIND.TKIND_ALIAS => ComTypeKind.Alias,
                        TYPEKIND.TKIND_MODULE => ComTypeKind.Module,
                        _ => ComTypeKind.Union,
                    };
                }
                finally
                {
                    target.ReleaseTypeAttr(attrPointer);
                }
            }

            return qualified;
        }
        finally
        {
            Marshal.ReleaseComObject(target);
        }
    }

    /// <summary>
    /// A constant or a parameter default as the library stores it. An object-valued default (the
    /// Office library has parameters defaulting to a null IDispatch) and anything else that has no
    /// text form is recorded by its type alone, never by throwing: one odd constant must not make
    /// a whole library unreadable.
    /// </summary>
    private static ComConstant ReadConstant(VARIANT* variant)
    {
        var value = VariantMarshal.FromNative(variant);
        string text;
        try
        {
            text = value.Type switch
            {
                Runtime.VarType.String => value.AsString(),
                Runtime.VarType.Boolean => value.AsBoolean() ? "True" : "False",
                Runtime.VarType.Empty or Runtime.VarType.Null or Runtime.VarType.Object or Runtime.VarType.Error => string.Empty,
                _ when value.IsObject || value.IsArray => string.Empty,
                _ => Runtime.Coerce.ToString(value),
            };
        }
        catch (Runtime.VbaException)
        {
            text = string.Empty;
        }

        return new ComConstant((int)value.Type, text);
    }
}
