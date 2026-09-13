using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>
/// Creates and finds COM objects for <c>CreateObject</c> and <c>GetObject</c> (ARCHITECTURE.md
/// section 7, "Creation"): ProgIDs resolve through the registry, running instances through the
/// Running Object Table, and file names through their monikers. Hosts install it with
/// <see cref="Com.Provider"/>.
/// </summary>
public class ComProvider : IComProvider
{
    /// <summary>IDispatch::Invoke from COM on a class instance or another object the runtime implements (<see cref="DispatchServer"/>).</summary>
    public int Invoke(IDispatchObject target, int dispId, ushort flags, nint parameters, nint result, nint exception, nint argumentError) =>
        DispatchServer.Invoke(target, dispId, flags, parameters, result, exception);

    /// <summary>A view over an IDispatch pointer a Variant holds (<see cref="ComObject.View"/>).</summary>
    public virtual IDispatchObject Wrap(nint dispatch) => ComObject.View(dispatch);

    /// <summary>New on a COM class: the object is a temporary of the current statement until a store keeps it (ARCHITECTURE.md D18).</summary>
    public object CreateInstance(Guid classId)
    {
        var hr = NativeMethods.CoCreateInstance(in classId, 0, NativeMethods.ClsctxServer, in NativeMethods.IidIDispatch, out var dispatch);
        return hr >= 0 && dispatch != 0 ? ObjectRefs.Owned(new ComObject(dispatch)) : throw new VbaException(VbaErrors.CannotCreateObject);
    }

    /// <summary>A running instance of the application object's class when one is registered in the Running Object Table, otherwise a new one.</summary>
    public virtual object GetAppObject(Guid libraryId, Guid classId)
    {
        if (NativeMethods.GetActiveObject(in classId, 0, out var unknown) >= 0 && unknown != 0)
        {
            try
            {
                return ObjectRefs.Owned(ComObject.FromPointer(unknown));
            }
            finally
            {
                NativeMethods.Release(unknown);
            }
        }

        return CreateInstance(classId);
    }

    public object CreateObject(string progId, string? serverName)
    {
        ArgumentNullException.ThrowIfNull(progId);
        if (!string.IsNullOrEmpty(serverName))
        {
            // Remote activation (CoCreateInstanceEx with a server name) is not supported yet (ROADMAP.md backlog).
            throw new VbaException(VbaErrors.CannotCreateObject);
        }

        if (NativeMethods.CLSIDFromProgID(progId, out var clsid) < 0)
        {
            throw new VbaException(VbaErrors.CannotCreateObject);
        }

        var hr = NativeMethods.CoCreateInstance(in clsid, 0, NativeMethods.ClsctxServer, in NativeMethods.IidIDispatch, out var dispatch);
        return hr >= 0 && dispatch != 0 ? ObjectRefs.Owned(new ComObject(dispatch)) : throw new VbaException(VbaErrors.CannotCreateObject);
    }

    public object GetObject(string? pathName, string? className)
    {
        if (string.IsNullOrEmpty(pathName))
        {
            if (string.IsNullOrEmpty(className))
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            if (NativeMethods.CLSIDFromProgID(className, out var clsid) < 0)
            {
                throw new VbaException(VbaErrors.CannotCreateObject);
            }

            if (NativeMethods.GetActiveObject(in clsid, 0, out var unknown) < 0 || unknown == 0)
            {
                throw new VbaException(VbaErrors.CannotCreateObject);
            }

            try
            {
                return ObjectRefs.Owned(ComObject.FromPointer(unknown));
            }
            finally
            {
                NativeMethods.Release(unknown);
            }
        }

        var hr = NativeMethods.CoGetObject(pathName, 0, in NativeMethods.IidIDispatch, out var dispatch);
        return hr >= 0 && dispatch != 0 ? ObjectRefs.Owned(new ComObject(dispatch)) : throw new VbaException(VbaErrors.FileOrClassNameNotFound);
    }
}
