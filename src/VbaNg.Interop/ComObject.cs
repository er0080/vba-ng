using System.Runtime.InteropServices;

using VbaNg.Runtime;

namespace VbaNg.Interop;

/// <summary>
/// A COM object as VBA code holds it: an owned <c>IDispatch</c> reference invoked by dispid with
/// hand-marshaled VARIANTs (ARCHITECTURE.md section 6, D7). Names resolve through
/// <c>GetIDsOfNames</c> once per object and are cached; errors come back as the run-time errors
/// VBA raises for them. Released when disposed, or by the finalizer.
/// </summary>
public sealed unsafe class ComObject : IDispatchObject, IComEventSource, IDisposable, IReferenceCounted, IComInterface
{
    private const int MaxStackArguments = 8;

    private nint dispatch;
    private Dictionary<string, int>? dispIds;
    private string? typeName;
    private int references;

    /// <summary>A view over a pointer its owner (a Variant, a SAFEARRAY) holds: no reference of its own until the first <see cref="AddRef"/> takes one.</summary>
    private bool borrowed;

    private static long unreleased;

    /// <summary>Takes ownership of an IDispatch reference.</summary>
    internal ComObject(nint dispatch)
    {
        this.dispatch = dispatch;
    }

    /// <summary>
    /// The finalizer never releases the pointer: the object lives in the apartment of the thread
    /// that made it (Excel's are STA), and the finalizer thread is not that thread, so a Release
    /// from here corrupts the object's owner. A wrapper that reaches the finalizer still holding
    /// its pointer is a missed ownership rail (ARCHITECTURE.md D18); it is counted, not released.
    /// </summary>
    ~ComObject()
    {
        if (dispatch != 0 && !borrowed)
        {
            Interlocked.Increment(ref unreleased);
        }
    }

    /// <summary>How many wrappers the garbage collector found still holding a pointer: zero when every object a call returned was released by its statement or its owner.</summary>
    public static long Unreleased => Interlocked.Read(ref unreleased);

    /// <summary>References VBA code holds (ARCHITECTURE.md D18); a host that keeps a wrapper for itself adds one so VBA code cannot release it.</summary>
    public void AddRef()
    {
        if (borrowed)
        {
            // A view being stored: from now on the wrapper holds a COM reference of its own, released with its last VBA reference.
            borrowed = false;
            if (dispatch != 0)
            {
                NativeMethods.AddRef(dispatch);
            }
        }

        references++;
    }

    /// <summary>The last VBA reference goes: the IDispatch is released at once, as VBA releases a COM object when its last variable clears.</summary>
    public void Release()
    {
        if (references > 0 && --references == 0)
        {
            ReleasePointer();
        }
    }

    /// <summary>Wraps a runtime-callable wrapper the host already holds, such as Excel-DNA's Application object.</summary>
    public static ComObject FromRcw(object rcw)
    {
        ArgumentNullException.ThrowIfNull(rcw);
        return new ComObject(Marshal.GetIDispatchForObject(rcw));
    }

    /// <summary>
    /// A view over an IDispatch pointer a Variant or a SAFEARRAY element holds (D20): calls go
    /// through it while its owner keeps the pointer, and it takes a reference of its own only when
    /// something stores it (<see cref="AddRef"/>). Two views of one object are the same object to
    /// <c>Is</c> (<see cref="IsSameObject"/>).
    /// </summary>
    internal static ComObject View(nint dispatch) => new(dispatch) { borrowed = true };

    /// <summary>The IDispatch pointer a Variant holds for this object, without a new reference; 0 once released.</summary>
    public nint InterfacePointer => dispatch;

    /// <summary>Wraps an interface pointer without consuming it; raises error 430 when it has no IDispatch.</summary>
    public static ComObject FromPointer(nint unknown)
    {
        var dispatch = NativeMethods.QueryInterface(unknown, in NativeMethods.IidIDispatch);
        return dispatch != 0 ? new ComObject(dispatch) : throw new VbaException(VbaErrors.ClassDoesNotSupportAutomation);
    }

    /// <summary>The name <c>TypeName</c> reports: the coclass name when the object provides class information, else its type info's name.</summary>
    public string TypeName => typeName ??= ComTypeName.Of(Dispatch);

    private nint Dispatch => dispatch != 0 ? dispatch : throw VbaErrors.ObjectVariableNotSet();

    /// <summary>An AddRef'd copy of the IDispatch pointer for a VARIANT the callee will release.</summary>
    internal nint AddRefDispatch()
    {
        var pointer = Dispatch;
        NativeMethods.AddRef(pointer);
        return pointer;
    }

    /// <summary>The IDispatch pointer without a new reference, for a query that completes before the wrapper can go away.</summary>
    internal nint DispatchPointer() => Dispatch;

    /// <summary>Advises an event sink on this object (ARCHITECTURE.md section 6, "Events").</summary>
    public void Advise(EventSink sink)
    {
        ArgumentNullException.ThrowIfNull(sink);
        sink.Advise(Dispatch);
    }

    /// <summary>
    /// A WithEvents variable of a library type was assigned this object (MS-VBAL 5.2.3.1.4): a sink
    /// on the compile-time source interface forwards each handled event to the variable's owner by
    /// name. An object that describes itself fires its arguments as IDispatch prescribes; one that
    /// does not is an Office object and fires them in declared order (<see cref="EventSource.DeclaredOrder"/>).
    /// A handler's unhandled error is reported through the host's output, as it is for document modules.
    /// </summary>
    public IDisposable Advise(ComEventInterface events, IVbaEventSink sink, string variable)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(sink);
        ArgumentNullException.ThrowIfNull(variable);
        var names = new Dictionary<int, string>();
        foreach (var (dispId, name) in events.Events)
        {
            names[dispId] = name;
        }

        var reorder = EventSource.Describe(this) is null;
        var eventSink = new EventSink(events.InterfaceId, (dispId, arguments) =>
        {
            if (!names.TryGetValue(dispId, out var name))
            {
                return;
            }

            // The sink decodes rgvarg as the convention says; a declared-order source needs the reverse, and the
            // ByRef write-back happens by position, so the array is put back afterwards.
            if (reorder && arguments.Length > 1)
            {
                Array.Reverse(arguments);
            }

            try
            {
                sink.RaiseVbaEvent(variable, name, arguments);
            }
            finally
            {
                if (reorder && arguments.Length > 1)
                {
                    Array.Reverse(arguments);
                }
            }
        })
        {
            Error = ex =>
            {
                // End in a handler stops it and resets the project (docs/vba-quirks.md); anything else is reported.
                if (ex is EndStatementException end)
                {
                    VbaNg.Runtime.Hosting.ProjectReset.AfterEnd(end);
                    return;
                }

                Host.Current.Print("Run-time error in " + variable + " event: " + (ex is VbaException error
                    ? string.Create(System.Globalization.CultureInfo.InvariantCulture, $"'{error.Number}': {error.Description}")
                    : ex.GetType().Name + ": " + ex.Message));
            },
        };
        try
        {
            eventSink.Advise(Dispatch);
        }
        catch
        {
            eventSink.Dispose();
            throw;
        }

        return eventSink;
    }

    public int GetDispId(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        dispIds ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (dispIds.TryGetValue(name, out var cached))
        {
            return cached;
        }

        var target = Dispatch;
        int dispId;
        fixed (char* chars = name)
        fixed (Guid* iid = &NativeMethods.IidNull)
        {
            var names = stackalloc char*[1];
            names[0] = chars;
            var hr = ((delegate* unmanaged[Stdcall]<nint, Guid*, char**, uint, uint, int*, int>)NativeMethods.Slot(target, 5))(target, iid, names, 1, NativeMethods.LocaleUserDefault, &dispId);
            if (hr < 0)
            {
                throw hr is NativeMethods.DispEUnknownName or NativeMethods.DispEMemberNotFound
                    ? new VbaException(VbaErrors.ObjectDoesNotSupportMember)
                    : ComErrors.FromHResult(hr);
            }
        }

        dispIds[name] = dispId;
        return dispId;
    }

    public Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments) => InvokeNamed(dispId, kind, arguments, []);

    public Variant InvokeNamed(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, ReadOnlySpan<int> namedDispIds)
    {
        VARIANT result;
        var hr = InvokeCore(dispId, kind, arguments, namedDispIds, hasValue: false, Variant.Empty, &result);
        try
        {
            if (hr < 0)
            {
                throw ComErrors.FromHResult(hr);
            }

            return VariantMarshal.FromNative(&result);
        }
        finally
        {
            VariantMarshal.Clear(&result);
        }
    }

    public void PutNamed(int dispId, ReadOnlySpan<Variant> indices, ReadOnlySpan<int> namedDispIds, in Variant value, bool asReference)
    {
        VARIANT result;
        var hr = InvokeCore(dispId, asReference ? InvokeKind.PropertyPutRef : InvokeKind.PropertyPut, indices, namedDispIds, hasValue: true, value, &result);
        VariantMarshal.Clear(&result);
        if (hr < 0)
        {
            throw ComErrors.FromHResult(hr);
        }
    }

    /// <summary>
    /// One GetIDsOfNames for the member and the argument names a call passes to it (MS-VBAL
    /// 5.6.13.1). DISP_E_UNKNOWNNAME leaves DISPID_UNKNOWN in the slots it could not answer: the
    /// member's is error 438, an argument's is error 448, as VBA reports them.
    /// </summary>
    public int[] GetDispIds(string name, ReadOnlySpan<string> argumentNames)
    {
        ArgumentNullException.ThrowIfNull(name);
        var count = argumentNames.Length + 1;
        var names = new nint[count];
        var dispIds = new int[count];
        try
        {
            names[0] = Marshal.StringToCoTaskMemUni(name);
            for (var i = 0; i < argumentNames.Length; i++)
            {
                names[i + 1] = Marshal.StringToCoTaskMemUni(argumentNames[i]);
            }

            var target = Dispatch;
            int hr;
            fixed (nint* pointers = names)
            fixed (int* ids = dispIds)
            fixed (Guid* iid = &NativeMethods.IidNull)
            {
                hr = ((delegate* unmanaged[Stdcall]<nint, Guid*, char**, uint, uint, int*, int>)NativeMethods.Slot(target, 5))(target, iid, (char**)pointers, (uint)count, NativeMethods.LocaleUserDefault, ids);
            }

            if (hr is NativeMethods.DispEUnknownName or NativeMethods.DispEMemberNotFound)
            {
                throw new VbaException(dispIds[0] == NativeMethods.DispidUnknown ? VbaErrors.ObjectDoesNotSupportMember : VbaErrors.NamedArgumentNotFound);
            }

            if (hr < 0)
            {
                throw ComErrors.FromHResult(hr);
            }

            return dispIds;
        }
        finally
        {
            foreach (var pointer in names)
            {
                if (pointer != 0)
                {
                    Marshal.FreeCoTaskMem(pointer);
                }
            }
        }
    }

    public void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference)
    {
        VARIANT result;
        var hr = InvokeCore(dispId, asReference ? InvokeKind.PropertyPutRef : InvokeKind.PropertyPut, indices, hasValue: true, value, &result);
        VariantMarshal.Clear(&result);
        if (hr < 0)
        {
            throw ComErrors.FromHResult(hr);
        }
    }

    /// <summary>DISPID_NEWENUM as VBA's For Each asks for it, through IEnumVARIANT.</summary>
    public IEnumerator<Variant> Enumerate()
    {
        VARIANT result;
        var hr = InvokeCore(NativeMethods.DispidNewEnum, InvokeKind.Method | InvokeKind.PropertyGet, [], hasValue: false, Variant.Empty, &result);
        try
        {
            if (hr < 0)
            {
                throw hr is NativeMethods.DispEMemberNotFound or NativeMethods.DispEUnknownName
                    ? new VbaException(VbaErrors.ObjectDoesNotSupportMember)
                    : ComErrors.FromHResult(hr);
            }

            var type = (ushort)(result.vt & Vt.TypeMask);
            if (type is not (Vt.Unknown or Vt.Dispatch) || result.pointer == 0)
            {
                throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
            }

            var enumerator = NativeMethods.QueryInterface(result.pointer, in NativeMethods.IidIEnumVariant);
            return enumerator != 0 ? new ComEnumerator(enumerator) : throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }
        finally
        {
            VariantMarshal.Clear(&result);
        }
    }

    public bool IsSameObject(object? other)
    {
        if (other is not ComObject com)
        {
            return false;
        }

        if (ReferenceEquals(com, this))
        {
            return true;
        }

        var left = NativeMethods.QueryInterface(Dispatch, in NativeMethods.IidIUnknown);
        var right = NativeMethods.QueryInterface(com.Dispatch, in NativeMethods.IidIUnknown);
        try
        {
            return left == right;
        }
        finally
        {
            NativeMethods.Release(left);
            NativeMethods.Release(right);
        }
    }

    public void Dispose()
    {
        if (borrowed)
        {
            // A view drops the pointer it never owned.
            dispatch = 0;
        }
        else
        {
            ReleasePointer();
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// IDispatch::Invoke with the arguments marshaled right to left, as rgvarg expects, and the
    /// value of a property put as the named DISPID_PROPERTYPUT argument. A DISP_E_EXCEPTION
    /// becomes the object's own error; other failures return the HRESULT for the caller to map.
    /// </summary>
    private int InvokeCore(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, bool hasValue, in Variant putValue, VARIANT* result) =>
        InvokeCore(dispId, kind, arguments, [], hasValue, putValue, result);

    /// <summary>
    /// DISPPARAMS as OLE Automation lays them out: the named arguments first, in the order of
    /// their dispids (the property-put value counts as one, under DISPID_PROPERTYPUT), then the
    /// positional arguments in reverse. <paramref name="arguments"/> holds the positional ones
    /// followed by the named ones in <paramref name="namedDispIds"/> order.
    /// </summary>
    private int InvokeCore(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, ReadOnlySpan<int> namedDispIds, bool hasValue, in Variant putValue, VARIANT* result)
    {
        var target = Dispatch;
        var count = arguments.Length + (hasValue ? 1 : 0);
        var positional = arguments.Length - namedDispIds.Length;
        var stack = stackalloc VARIANT[MaxStackArguments];
        var heap = count > MaxStackArguments ? (VARIANT*)NativeMemory.Alloc((nuint)count, (nuint)sizeof(VARIANT)) : null;
        var natives = heap != null ? heap : stack;
        var namedIds = stackalloc int[namedDispIds.Length + 1];
        var marshaled = 0;
        var named = 0;
        try
        {
            if (hasValue)
            {
                VariantMarshal.ToNative(putValue, natives);
                marshaled++;
                namedIds[named++] = NativeMethods.DispidPropertyPut;
            }

            for (var i = 0; i < namedDispIds.Length; i++)
            {
                VariantMarshal.ToNative(arguments[positional + i], natives + marshaled);
                marshaled++;
                namedIds[named++] = namedDispIds[i];
            }

            for (var i = 0; i < positional; i++)
            {
                VariantMarshal.ToNative(arguments[positional - 1 - i], natives + marshaled);
                marshaled++;
            }

            var parameters = new DISPPARAMS
            {
                rgvarg = (nint)natives,
                rgdispidNamedArgs = named > 0 ? (nint)namedIds : 0,
                cArgs = (uint)count,
                cNamedArgs = (uint)named,
            };
            EXCEPINFO exception = default;
            uint argumentError;
            result->vt = Vt.Empty;
            result->data = 0;
            fixed (Guid* iid = &NativeMethods.IidNull)
            {
                var hr = ((delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, DISPPARAMS*, VARIANT*, EXCEPINFO*, uint*, int>)NativeMethods.Slot(target, 6))(
                    target, dispId, iid, NativeMethods.LocaleUserDefault, (ushort)kind, &parameters, result, &exception, &argumentError);
                if (hr == NativeMethods.DispEException)
                {
                    throw ComErrors.FromExceptionInfo(&exception);
                }

                return hr;
            }
        }
        finally
        {
            for (var i = 0; i < marshaled; i++)
            {
                VariantMarshal.Clear(natives + i);
            }

            if (heap is not null)
            {
                NativeMemory.Free(heap);
            }
        }
    }

    private void ReleasePointer()
    {
        var pointer = dispatch;
        if (pointer != 0)
        {
            dispatch = 0;
            NativeMethods.Release(pointer);
        }
    }
}

/// <summary>
/// IEnumVARIANT as For Each walks it: one element per Next call, released at the end. The element
/// the loop is on is the enumerator's own (a reference, a copy of a string) until the next Next or the end: the
/// header statement's temporaries end before the loop variable takes its reference, and across
/// processes a proxy no one holds is destroyed at once (ARCHITECTURE.md D18).
/// </summary>
internal sealed unsafe class ComEnumerator : IEnumerator<Variant>
{
    private nint enumerator;

    public ComEnumerator(nint enumerator)
    {
        this.enumerator = enumerator;
    }

    ~ComEnumerator()
    {
        // The finalizer thread is not the object's: the element's reference is left, as the statement's would be.
        Release();
    }

    public Variant Current { get; private set; } = Variant.Empty;

    object System.Collections.IEnumerator.Current => Current;

    public bool MoveNext()
    {
        if (enumerator == 0)
        {
            return false;
        }

        VARIANT item;
        item.vt = Vt.Empty;
        item.data = 0;
        uint fetched;
        var hr = ((delegate* unmanaged[Stdcall]<nint, uint, VARIANT*, uint*, int>)NativeMethods.Slot(enumerator, 3))(enumerator, 1, &item, &fetched);
        if (hr < 0)
        {
            throw ComErrors.FromHResult(hr);
        }

        ReleaseCurrent();
        if (hr != NativeMethods.SOk || fetched != 1)
        {
            return false;
        }

        try
        {
            Current = ObjectRefs.Own(VariantMarshal.FromNative(&item));
        }
        finally
        {
            VariantMarshal.Clear(&item);
        }

        return true;
    }

    public void Reset()
    {
        if (enumerator != 0)
        {
            ((delegate* unmanaged[Stdcall]<nint, int>)NativeMethods.Slot(enumerator, 5))(enumerator);
        }
    }

    public void Dispose()
    {
        ReleaseCurrent();
        Release();
        GC.SuppressFinalize(this);
    }

    private void ReleaseCurrent()
    {
        var current = Current;
        Current = Variant.Empty;
        ObjectRefs.Release(ref current);
    }

    private void Release()
    {
        var pointer = enumerator;
        if (pointer != 0)
        {
            enumerator = 0;
            NativeMethods.Release(pointer);
        }
    }
}
