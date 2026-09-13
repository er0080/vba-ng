using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>Receives an event: the dispid the source raised and its arguments; a ByRef argument written back here reaches the source (Cancel = True).</summary>
public delegate void ComEventCallback(int dispId, Variant[] arguments);

/// <summary>
/// A COM event sink (ARCHITECTURE.md section 6, "Events"): an <c>IDispatch</c> implementation
/// built in unmanaged memory whose <c>Invoke</c> forwards to a managed handler, advised on a
/// source through <c>IConnectionPointContainer</c>. One sink serves one source interface of one
/// object; every advisory is undone by <see cref="Dispose"/>, so an unload leaves nothing
/// connected. Errors thrown by the handler go to <see cref="Error"/> and never reach the source.
/// </summary>
public sealed unsafe class EventSink : IDisposable
{
    private const int SOk = 0;
    private const int ENoInterface = unchecked((int)0x80004002);
    private const int ENotImpl = unchecked((int)0x80004001);
    private const int DispEUnknownName = unchecked((int)0x80020006);

    private static readonly Guid IidIConnectionPointContainer = new("B196B284-BAB4-101A-B69C-00AA00341D07");
    private static readonly nint Vtable = BuildVtable();

    private readonly ComEventCallback handler;
    private readonly GCHandle self;
    private nint block;
    private int references;
    private nint connectionPoint;
    private uint cookie;
    private bool disposed;

    public EventSink(Guid sourceInterface, ComEventCallback handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        SourceInterface = sourceInterface;
        this.handler = handler;
        self = GCHandle.Alloc(this);
        block = (nint)NativeMemory.Alloc((nuint)(2 * sizeof(nint)));
        *(nint*)block = Vtable;
        *(nint*)(block + sizeof(nint)) = GCHandle.ToIntPtr(self);
        references = 1;
    }

    /// <summary>The IID of the source dispinterface this sink implements.</summary>
    public Guid SourceInterface { get; }

    /// <summary>Receives exceptions the handler threw; the event still returns S_OK to the source.</summary>
    public Action<Exception>? Error { get; set; }

    /// <summary>The sink as an IUnknown pointer, for the source to AddRef.</summary>
    public nint Unknown => block;

    /// <summary>True while advised on a source.</summary>
    public bool IsAdvised => connectionPoint != 0;

    /// <summary>Connects the sink to a source object's connection point for <see cref="SourceInterface"/>.</summary>
    public void Advise(nint source)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        if (connectionPoint != 0)
        {
            throw new InvalidOperationException("The sink is already advised.");
        }

        var container = NativeMethods.QueryInterface(source, in IidIConnectionPointContainer);
        if (container == 0)
        {
            throw new VbaException(459);
        }

        try
        {
            nint point;
            var iid = SourceInterface;
            var found = ((delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)NativeMethods.Slot(container, 4))(container, &iid, &point);
            if (found < 0 || point == 0)
            {
                throw new VbaException(459);
            }

            uint token;
            var advised = ((delegate* unmanaged[Stdcall]<nint, nint, uint*, int>)NativeMethods.Slot(point, 5))(point, block, &token);
            if (advised < 0)
            {
                NativeMethods.Release(point);
                throw ComErrors.FromHResult(advised);
            }

            connectionPoint = point;
            cookie = token;
        }
        finally
        {
            NativeMethods.Release(container);
        }
    }

    /// <summary>Disconnects from the source; safe to call when not advised.</summary>
    public void Unadvise()
    {
        var point = connectionPoint;
        if (point == 0)
        {
            return;
        }

        connectionPoint = 0;
        try
        {
            _ = ((delegate* unmanaged[Stdcall]<nint, uint, int>)NativeMethods.Slot(point, 6))(point, cookie);
        }
        finally
        {
            NativeMethods.Release(point);
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        Unadvise();
        ReleaseNative();
    }

    /// <summary>Invoke as the source calls it: the arguments arrive right to left, ByRef ones are written back after the handler ran.</summary>
    private void Dispatch(int dispId, DISPPARAMS* parameters)
    {
        // The event's arguments, and whatever the handler leaves in a ByRef one, are temporaries of this
        // call: released after the write-back, on the thread that made them (ARCHITECTURE.md D18).
        using var frame = ObjectRefs.Frame();
        var count = parameters is null ? 0 : (int)parameters->cArgs;
        var arguments = new Variant[count];
        var natives = (VARIANT*)(parameters is null ? 0 : parameters->rgvarg);
        for (var i = 0; i < count; i++)
        {
            arguments[i] = VariantMarshal.FromNative(natives + (count - 1 - i));
        }

        try
        {
            handler(dispId, arguments);
        }
        catch (Exception ex)
        {
            Error?.Invoke(ex);
            return;
        }

        for (var i = 0; i < count; i++)
        {
            var native = natives + (count - 1 - i);
            if ((native->vt & Vt.ByRef) != 0)
            {
                VariantMarshal.WriteByRef(native, arguments[i]);
            }
        }
    }

    private void ReleaseNative()
    {
        if (block != 0 && Interlocked.Decrement(ref references) == 0)
        {
            FreeBlock();
        }
    }

    private void FreeBlock()
    {
        var memory = block;
        block = 0;
        if (memory != 0)
        {
            *(nint*)(memory + sizeof(nint)) = 0;
            NativeMemory.Free((void*)memory);
        }

        if (self.IsAllocated)
        {
            self.Free();
        }
    }

    private static EventSink? FromNative(nint self)
    {
        var handle = *(nint*)(self + sizeof(nint));
        return handle == 0 ? null : GCHandle.FromIntPtr(handle).Target as EventSink;
    }

    private static nint BuildVtable()
    {
        var table = (nint*)NativeMemory.Alloc((nuint)(7 * sizeof(nint)));
        table[0] = (nint)(delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)&QueryInterface;
        table[1] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&AddRef;
        table[2] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&Release;
        table[3] = (nint)(delegate* unmanaged[Stdcall]<nint, uint*, int>)&GetTypeInfoCount;
        table[4] = (nint)(delegate* unmanaged[Stdcall]<nint, uint, uint, nint*, int>)&GetTypeInfo;
        table[5] = (nint)(delegate* unmanaged[Stdcall]<nint, Guid*, nint, uint, uint, int*, int>)&GetIDsOfNames;
        table[6] = (nint)(delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, DISPPARAMS*, VARIANT*, EXCEPINFO*, uint*, int>)&Invoke;
        return (nint)table;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int QueryInterface(nint self, Guid* iid, nint* result)
    {
        var sink = FromNative(self);
        if (sink is not null && (*iid == NativeMethods.IidIUnknown || *iid == NativeMethods.IidIDispatch || *iid == sink.SourceInterface))
        {
            Interlocked.Increment(ref sink.references);
            *result = self;
            return SOk;
        }

        *result = 0;
        return ENoInterface;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint AddRef(nint self)
    {
        var sink = FromNative(self);
        return sink is null ? 1 : (uint)Interlocked.Increment(ref sink.references);
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint Release(nint self)
    {
        var sink = FromNative(self);
        if (sink is null)
        {
            return 0;
        }

        var remaining = Interlocked.Decrement(ref sink.references);
        if (remaining == 0)
        {
            sink.FreeBlock();
        }

        return (uint)remaining;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetTypeInfoCount(nint self, uint* count)
    {
        *count = 0;
        return SOk;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetTypeInfo(nint self, uint index, uint lcid, nint* typeInfo)
    {
        *typeInfo = 0;
        return ENotImpl;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetIDsOfNames(nint self, Guid* iid, nint names, uint count, uint lcid, int* dispIds) => DispEUnknownName;

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int Invoke(nint self, int dispId, Guid* iid, uint lcid, ushort flags, DISPPARAMS* parameters, VARIANT* result, EXCEPINFO* exception, uint* argumentError)
    {
        var sink = FromNative(self);
        if (sink is null)
        {
            return SOk;
        }

        if (result is not null)
        {
            result->vt = Vt.Empty;
            result->data = 0;
        }

        try
        {
            sink.Dispatch(dispId, parameters);
        }
        catch (Exception ex)
        {
            sink.Error?.Invoke(ex);
        }

        return SOk;
    }
}
