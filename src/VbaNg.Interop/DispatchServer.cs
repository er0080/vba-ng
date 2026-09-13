using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>
/// IDispatch::Invoke for the objects the runtime implements (<see cref="RuntimeObject"/>;
/// ARCHITECTURE.md section 5, "Class modules"), which the runtime's vtable thunk reaches through
/// <see cref="IComProvider.Invoke"/>. The DISPPARAMS become Variants (positional arguments are
/// stored right to left after the named ones), the member runs through the object's own dispid
/// table (<see cref="IDispatchObject"/>), the result goes to the caller's VARIANT, and a run-time
/// error becomes DISP_E_EXCEPTION with an EXCEPINFO: the SCODE is 0x800A0000 plus the number (a
/// number that is already an HRESULT, from vbObjectError, as it is), with the error's
/// description, source, help file, and context, which is how <see cref="ComErrors"/> reads such
/// an error back.
/// </summary>
internal static unsafe class DispatchServer
{
    private const int EFail = unchecked((int)0x80004005);

    public static int Invoke(IDispatchObject target, int dispId, ushort flags, nint parameters, nint result, nint exception)
    {
        // The arguments and the result are temporaries of this call (ARCHITECTURE.md D18, D20); ToNative copies what the caller keeps.
        using var frame = ObjectRefs.Frame();
        var dispParams = (DISPPARAMS*)parameters;
        var count = dispParams is null ? 0 : (int)dispParams->cArgs;
        var namedCount = dispParams is null ? 0 : (int)dispParams->cNamedArgs;
        var natives = dispParams is null ? null : (VARIANT*)dispParams->rgvarg;
        var namedIds = dispParams is null ? null : (int*)dispParams->rgdispidNamedArgs;
        if (result != 0)
        {
            ((VARIANT*)result)->vt = Vt.Empty;
            ((VARIANT*)result)->data = 0;
        }

        try
        {
            var kind = (InvokeKind)(flags & 0xF);
            var positional = count - namedCount;
            if ((kind & (InvokeKind.PropertyPut | InvokeKind.PropertyPutRef)) != 0)
            {
                // A property put: the value is the argument named DISPID_PROPERTYPUT; the others are the indices.
                if (namedCount != 1 || namedIds[0] != NativeMethods.DispidPropertyPut)
                {
                    return NativeMethods.DispEParamNotFound;
                }

                var indices = new Variant[positional];
                for (var i = 0; i < positional; i++)
                {
                    indices[i] = VariantMarshal.FromNative(natives + (count - 1 - i));
                }

                target.Put(dispId, indices, VariantMarshal.FromNative(natives), (kind & InvokeKind.PropertyPutRef) != 0);
                return NativeMethods.SOk;
            }

            var arguments = new Variant[count];
            for (var i = 0; i < positional; i++)
            {
                arguments[i] = VariantMarshal.FromNative(natives + (count - 1 - i));
            }

            for (var i = 0; i < namedCount; i++)
            {
                arguments[positional + i] = VariantMarshal.FromNative(natives + i);
            }

            var value = namedCount == 0
                ? target.Invoke(dispId, kind, arguments)
                : target.InvokeNamed(dispId, kind, arguments, new ReadOnlySpan<int>(namedIds, namedCount));
            if (result != 0)
            {
                VariantMarshal.ToNative(value, (VARIANT*)result);
            }

            return NativeMethods.SOk;
        }
        catch (VbaException error) when (error.Number is VbaErrors.WrongNumberOfArguments or VbaErrors.ArgumentNotOptional)
        {
            // Arguments that do not fit the member: the codes OLE Automation's own dispatch returns, as VBA's objects do, which a COM
            // caller passes on (Scripting.Dictionary evaluating a Collection's default value) and VBA reads as 450 and 449 (Objects golden).
            return error.Number == VbaErrors.WrongNumberOfArguments ? NativeMethods.DispEBadParamCount : NativeMethods.DispEParamNotOptional;
        }
        catch (VbaException error)
        {
            var scode = error.Number is > 0 and <= 0xFFFF ? unchecked((int)(0x800A0000u | (uint)error.Number)) : error.Number;
            Describe((EXCEPINFO*)exception, scode, error.Description, error.Source, error.HelpFile, error.HelpContext);
            return NativeMethods.DispEException;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            Describe((EXCEPINFO*)exception, EFail, ex.Message, null, null, 0);
            return NativeMethods.DispEException;
        }
    }

    private static void Describe(EXCEPINFO* info, int scode, string? description, string? source, string? helpFile, int helpContext)
    {
        if (info is null)
        {
            return;
        }

        *info = default;
        info->scode = scode;
        info->bstrDescription = VariantMarshal.AllocBstr(description ?? string.Empty);
        info->bstrSource = VariantMarshal.AllocBstr(source ?? string.Empty);
        info->bstrHelpFile = VariantMarshal.AllocBstr(helpFile ?? string.Empty);
        info->dwHelpContext = (uint)helpContext;
    }
}
