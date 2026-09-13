using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>Turns the outcome of an <c>IDispatch::Invoke</c> into the run-time error VBA raises for it (CLAUDE.md R5).</summary>
internal static unsafe class ComErrors
{
    /// <summary>
    /// DISP_E_EXCEPTION: the object described its error. The number is <c>wCode</c> when set,
    /// otherwise the SCODE with VBA's facility stripped (Excel's 0x800A03EC is 1004); the
    /// description, source, and help file are the object's own, which is what Err reports.
    /// </summary>
    public static VbaException FromExceptionInfo(EXCEPINFO* info)
    {
        if (info->pfnDeferredFillIn != 0)
        {
            ((delegate* unmanaged[Stdcall]<EXCEPINFO*, int>)info->pfnDeferredFillIn)(info);
        }

        var number = info->wCode != 0 ? info->wCode : new ErrorValue(info->scode).Number;
        if (info->wCode == 0 && number == info->scode)
        {
            // An object that passes on another's failure reports that HRESULT here, a DISP_E code among them, which reads as the
            // HRESULT itself would (Objects golden: Scripting.Dictionary refusing a Collection as an Item value is 450).
            number = FromHResult(info->scode).Number;
        }

        var description = TakeBstr(ref info->bstrDescription);
        var source = TakeBstr(ref info->bstrSource);
        var helpFile = TakeBstr(ref info->bstrHelpFile);
        return new VbaException(
            number,
            description.Length > 0 ? description : null,
            source.Length > 0 ? source : null,
            helpFile.Length > 0 ? helpFile : string.Empty,
            (int)info->dwHelpContext);
    }

    /// <summary>A failed HRESULT without exception info: the DISP_E codes map to VBA's numbers, anything else is an Automation error carrying the HRESULT.</summary>
    public static VbaException FromHResult(int hr) => hr switch
    {
        NativeMethods.DispEMemberNotFound or NativeMethods.DispEUnknownName => new VbaException(VbaErrors.ObjectDoesNotSupportMember),
        NativeMethods.DispEParamNotFound => new VbaException(VbaErrors.NamedArgumentNotFound),
        NativeMethods.DispETypeMismatch or NativeMethods.DispEBadVarType => VbaErrors.TypeMismatch(),
        NativeMethods.DispENoNamedArgs => new VbaException(VbaErrors.NoNamedArguments),
        NativeMethods.DispEOverflow => VbaErrors.Overflow(),
        NativeMethods.DispEBadIndex => VbaErrors.SubscriptOutOfRange(),
        NativeMethods.DispEUnknownLcid => new VbaException(VbaErrors.UnsupportedLocale),
        NativeMethods.DispEArrayIsLocked => new VbaException(VbaErrors.ArrayFixedOrLocked),
        NativeMethods.DispEBadParamCount => new VbaException(VbaErrors.WrongNumberOfArguments),
        NativeMethods.DispEParamNotOptional => new VbaException(VbaErrors.ArgumentNotOptional),
        NativeMethods.DispENotACollection => new VbaException(451),
        NativeMethods.EInvalidArg => VbaErrors.InvalidProcedureCall(),
        NativeMethods.EOutOfMemory => new VbaException(VbaErrors.OutOfMemory),
        NativeMethods.ENoInterface => new VbaException(VbaErrors.ClassDoesNotSupportAutomation),
        _ => new VbaException(hr),
    };

    private static string TakeBstr(ref nint bstr)
    {
        if (bstr == 0)
        {
            return string.Empty;
        }

        var text = VariantMarshal.ReadBstr(bstr);
        NativeMethods.SysFreeString(bstr);
        bstr = 0;
        return text;
    }
}
