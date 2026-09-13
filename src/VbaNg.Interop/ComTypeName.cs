namespace VbaNg.Interop;

/// <summary>
/// The name <c>TypeName</c> reports for a COM object: the coclass name from
/// <c>IProvideClassInfo</c> when the object offers it (Excel's objects, the Scripting objects),
/// otherwise the name of the type info behind its <c>IDispatch</c> with a leading underscore
/// removed, and "Object" when it has no type information at all.
/// </summary>
internal static unsafe class ComTypeName
{
    private const int MemberIdNil = -1;

    public static string Of(nint dispatch)
    {
        var classInfo = NativeMethods.QueryInterface(dispatch, in NativeMethods.IidIProvideClassInfo);
        if (classInfo != 0)
        {
            try
            {
                nint typeInfo;
                var hr = ((delegate* unmanaged[Stdcall]<nint, nint*, int>)NativeMethods.Slot(classInfo, 3))(classInfo, &typeInfo);
                if (hr == NativeMethods.SOk && typeInfo != 0)
                {
                    var name = NameOf(typeInfo);
                    if (name is not null)
                    {
                        return name;
                    }
                }
            }
            finally
            {
                NativeMethods.Release(classInfo);
            }
        }

        uint count;
        if (((delegate* unmanaged[Stdcall]<nint, uint*, int>)NativeMethods.Slot(dispatch, 3))(dispatch, &count) == NativeMethods.SOk && count > 0)
        {
            nint typeInfo;
            var hr = ((delegate* unmanaged[Stdcall]<nint, uint, uint, nint*, int>)NativeMethods.Slot(dispatch, 4))(dispatch, 0, NativeMethods.LocaleUserDefault, &typeInfo);
            if (hr == NativeMethods.SOk && typeInfo != 0)
            {
                var name = NameOf(typeInfo);
                if (name is not null)
                {
                    return name.StartsWith('_') ? name[1..] : name;
                }
            }
        }

        return "Object";
    }

    /// <summary>ITypeInfo::GetDocumentation(MEMBERID_NIL) for the type's own name; releases the type info.</summary>
    private static string? NameOf(nint typeInfo)
    {
        try
        {
            nint name;
            var hr = ((delegate* unmanaged[Stdcall]<nint, int, nint*, nint*, uint*, nint*, int>)NativeMethods.Slot(typeInfo, 12))(typeInfo, MemberIdNil, &name, null, null, null);
            if (hr != NativeMethods.SOk || name == 0)
            {
                return null;
            }

            var text = VariantMarshal.ReadBstr(name);
            NativeMethods.SysFreeString(name);
            return text.Length > 0 ? text : null;
        }
        finally
        {
            NativeMethods.Release(typeInfo);
        }
    }
}
