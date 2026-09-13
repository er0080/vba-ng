using System.Collections;
using System.Runtime.InteropServices;

using VbaNg.Runtime.Hosting;
using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>
/// An object whose lifetime VBA code controls by reference count (ARCHITECTURE.md D18): project
/// class instances, <c>Collection</c>, and COM wrappers. Generated code adds and drops references
/// through <see cref="ObjectRefs"/>; the object acts when its count reaches zero.
/// </summary>
public interface IReferenceCounted
{
    void AddRef();

    void Release();
}

/// <summary>
/// What a variable of a project class type holds (ARCHITECTURE.md D18): the object's members by
/// dispid and its reference count. A class another class implements is emitted as a C# interface
/// extending this one, so a variable of the interface type carries the same contract.
/// </summary>
public interface IVbaClassInstance : IDispatchObject, IReferenceCounted;

/// <summary>
/// A class module with WithEvents variables (MS-VBAL 5.2.3.1.4): the source calls this when it
/// raises an event, naming the variable it was assigned to and the event; the handler's ByRef
/// parameters are written back into <paramref name="arguments"/>.
/// </summary>
public interface IVbaEventSink
{
    void RaiseVbaEvent(string source, string name, Variant[] arguments);
}

/// <summary>
/// The events a WithEvents variable of a library type handles (MS-VBAL 5.2.3.1.4; ARCHITECTURE.md
/// section 7, "Events"): the default source interface of its coclass, from the type library at
/// compile time, and the handled events by dispid. Generated code keeps one per variable.
/// </summary>
public sealed record ComEventInterface(Guid InterfaceId, string InterfaceName, IReadOnlyList<(int DispId, string Name)> Events);

/// <summary>
/// An object that raises COM events. The Interop wrapper of a COM object implements it, so that
/// Set on a WithEvents variable of a library type advises a sink on the object without the
/// runtime knowing COM (<see cref="ObjectRefs.Subscribe{T}(ref T, T, IVbaEventSink, string, ComEventInterface, ref IDisposable)"/>).
/// </summary>
public interface IComEventSource
{
    /// <summary>
    /// Advises on the source interface: each event in <paramref name="events"/> the object raises
    /// runs <c>sink.RaiseVbaEvent(variable, name, arguments)</c>, ByRef arguments written back to
    /// the object. Disposing the result unadvises. Error 459 when the object has no such source.
    /// </summary>
    IDisposable Advise(ComEventInterface events, IVbaEventSink sink, string variable);
}

/// <summary>
/// The base of every generated class module (MS-VBAL 4.2, 5.2.4.1.3): a reference-counted
/// object that runs <c>Class_Terminate</c> when its last reference goes away, and an
/// <see cref="IDispatchObject"/> whose members generated code exposes by name and by dispid,
/// so late binding, <c>CallByName</c>, default members (dispid 0), and <c>For Each</c> (dispid -4)
/// reach a project class through the same paths as a COM object.
/// </summary>
public abstract class VbaClassObject : RuntimeObject, IVbaClassInstance
{
    private bool terminated;
    private List<(WeakReference<IVbaEventSink> Sink, string Source)>? handlers;

    public abstract string TypeName { get; }

    /// <summary>The last reference went, from VBA code or from COM: Class_Terminate once, then the instance's storage.</summary>
    protected sealed override void OnLastRelease()
    {
        if (terminated)
        {
            return;
        }

        terminated = true;
        try
        {
            // A project being reset (End, its workbook closing) destroys its objects without Class_Terminate (docs/vba-quirks.md).
            if (!ProjectReset.IsResetting(GetType()))
            {
                ClassTerminate();
            }
        }
        finally
        {
            // Class_Terminate runs before the object's own references go, and they go even when it raises (Classes golden: Outer).
            ReleaseFields();
        }
    }

    /// <summary>A project reset destroys the instance whatever its count: its storage goes, Class_Terminate does not run (docs/vba-quirks.md).</summary>
    internal override void Destroy()
    {
        if (terminated)
        {
            return;
        }

        terminated = true;
        ReleaseFields();
    }

    public abstract int GetDispId(string name);

    public abstract Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments);

    public abstract void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference);

    /// <summary>For Each over the object: its NewEnum member (dispid -4); a class without one raises 438.</summary>
    public virtual IEnumerator<Variant> Enumerate() => throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);

    public bool IsSameObject(object? other) => ReferenceEquals(this, other);

    /// <summary>Class_Terminate (MS-VBAL 5.2.4.1.3); a generated class that declares one overrides.</summary>
    protected virtual void ClassTerminate()
    {
    }

    /// <summary>The release of the instance's own storage: its object references, strings, arrays, and records (ARCHITECTURE.md D18, D20); generated classes override.</summary>
    protected virtual void ReleaseFields()
    {
    }

    /// <summary>A WithEvents variable of <paramref name="sink"/> was assigned this object: its handlers run when this object raises an event.</summary>
    internal void Subscribe(IVbaEventSink sink, string source)
    {
        handlers ??= [];
        handlers.Add((new WeakReference<IVbaEventSink>(sink), source));
    }

    /// <summary>The WithEvents variable was cleared, or its owner terminated: the handlers stop running.</summary>
    internal void Unsubscribe(IVbaEventSink sink, string source)
    {
        handlers?.RemoveAll(h => !h.Sink.TryGetTarget(out var target) || (ReferenceEquals(target, sink) && string.Equals(h.Source, source, StringComparison.OrdinalIgnoreCase)));
    }

    /// <summary>
    /// RaiseEvent (MS-VBAL 5.4.2.10): every object that holds this one in a WithEvents variable
    /// runs its handler, in subscription order. A handler writes a ByRef parameter back into
    /// <paramref name="arguments"/>, which the raising procedure then sees.
    /// </summary>
    protected void RaiseVbaEvent(string name, Variant[] arguments)
    {
        if (handlers is null)
        {
            return;
        }

        foreach (var (weak, source) in handlers.ToArray())
        {
            if (weak.TryGetTarget(out var sink))
            {
                sink.RaiseVbaEvent(source, name, arguments);
            }
        }
    }

    /// <summary>The dispid of a member that does not exist: error 438 (MS-VBAL 5.6.12).</summary>
    protected static int NoMember() => throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);

    /// <summary>A member call whose kind the member cannot answer (a Let on a Sub, a call on a Let-only property): error 438.</summary>
    protected static Variant NotSupported() => throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);

    /// <summary>The argument at <paramref name="index"/> of a late-bound call, Missing when the caller left it out.</summary>
    protected static Variant Argument(ReadOnlySpan<Variant> arguments, int index) => index < arguments.Length ? arguments[index] : Variant.Missing;
}

/// <summary>
/// A class module whose instances keep their variables in a block of their own in native memory
/// (ARCHITECTURE.md section 5; ROADMAP.md M7 E4), laid out as the generated struct of them, so
/// VarPtr of a class variable is a stable address. The last reference's ReleaseFields leaves the
/// variables empty; the block itself goes with the object, so code still running on an object
/// whose last reference went reads empty variables rather than freed memory.
/// </summary>
/// <typeparam name="TFields">The generated struct of the class's variables.</typeparam>
public abstract unsafe class VbaClassObject<TFields> : VbaClassObject
    where TFields : unmanaged
{
    private readonly nint fields;

    protected VbaClassObject(TFields initial)
    {
        fields = (nint)NativeMemory.Alloc((nuint)sizeof(TFields));
        *(TFields*)fields = initial;
    }

    ~VbaClassObject() => NativeMemory.Free((void*)fields);

    internal ref TFields Fields => ref *(TFields*)fields;
}

/// <summary>The block of a class instance's variables, for generated code, which reaches each variable through it (ROADMAP.md M7 E4).</summary>
public static class InstanceBlocks
{
    public static ref TFields Of<TFields>(VbaClassObject<TFields> instance)
        where TFields : unmanaged
    {
        ArgumentNullException.ThrowIfNull(instance);
        return ref instance.Fields;
    }
}

/// <summary>
/// The reference-count bookkeeping generated code performs (ARCHITECTURE.md D18). Slots are
/// object-typed variables, fields, and Variants; owned temporaries are the results of New and
/// of calls, released when the statement that produced them ends, or by the procedure's frame
/// when an error cuts the statement short.
/// </summary>
public static class ObjectRefs
{
    [ThreadStatic]
    private static List<Temporary>? temporaries;

    [ThreadStatic]
    private static Temporary? pendingResult;

    private static List<Temporary> Temporaries => temporaries ??= new List<Temporary>(16);

    /// <summary>
    /// What a statement owns until it ends (ARCHITECTURE.md D18, D20): a reference-counted
    /// object, an array it destroys, a record a Function returned, the
    /// cell behind a late-bound ByRef argument, a BSTR, or a reference on the interface pointer a
    /// Variant holds for an object.
    /// </summary>
    private readonly record struct Temporary(object? Reference, nint Pointer, bool IsInterface = false, VarType ArrayElement = VarType.Empty)
    {
        public void Release()
        {
            if (IsInterface)
            {
                RuntimeObject.ReleaseInterface(Pointer);
                return;
            }

            if (ArrayElement != VarType.Empty)
            {
                var array = VbaArray.View(Pointer, ArrayElement);
                ObjectRefs.Release(ref array);
                return;
            }

            switch (Reference)
            {
                case null:
                    Bstr.Free(Pointer);
                    break;
                case Variant[] cell:
                    ObjectRefs.Release(ref cell[0]);
                    break;
                case VbaArray[] cell:
                    ObjectRefs.Release(ref cell[0]);
                    break;
                case VbaString[] cell:
                    ObjectRefs.Release(ref cell[0]);
                    break;
                case object?[] cell:
                    ReleaseReference(cell[0]);
                    break;
                case ICell cell:
                    cell.Release();
                    break;
                default:
                    ReleaseReference(Reference);
                    break;
            }
        }
    }

    /// <summary>A cell on the heap that owns what it holds until its statement ends.</summary>
    private interface ICell
    {
        void Release();
    }

    /// <summary>The cell of an object slot (<see cref="Cell{T}"/>).</summary>
    private sealed class SlotCell<T> : ICell
        where T : class
    {
        public ObjectSlot<T> Slot;

        public void Release() => ObjectRefs.Release(ref Slot);
    }

    /// <summary>What a temporary or a cell held loses what it owned: a counted object its reference, an array its elements, a record a Function returned its references.</summary>
    private static void ReleaseReference(object? reference)
    {
        switch (reference)
        {
            case IReferenceCounted counted:
                counted.Release();
                break;
            case IVbaRecord record:
                record.ReleaseReferences();
                break;
        }
    }

    /// <summary>Set slot = value (MS-VBAL 5.4.3.9): the new object gains a reference before the old one loses its own.</summary>
    public static void Assign<T>(ref T? slot, T? value)
        where T : class
    {
        if (ReferenceEquals(slot, value))
        {
            return;
        }

        (value as IReferenceCounted)?.AddRef();
        var old = slot;
        slot = value;
        (old as IReferenceCounted)?.Release();
    }

    /// <summary>Set slot = value into a typed object variable (MS-VBAL 5.4.3.9; ROADMAP.md M7 E3): the slot takes a reference on the new object's interface pointer before the old one loses its own, so Set x = x is safe.</summary>
    public static void Assign<T>(ref ObjectSlot<T> slot, T? value)
        where T : class
    {
        var pointer = HeldPointer(value);
        RuntimeObject.ReleaseInterface(slot.Exchange(pointer));
    }

    /// <summary>A slot holding a reference of its own on the value: a ByVal object parameter's storage for the call, and the temporary a ByRef object parameter works on.</summary>
    public static ObjectSlot<T> Holding<T>(T? value)
        where T : class
    {
        var slot = default(ObjectSlot<T>);
        Assign(ref slot, value);
        return slot;
    }

    /// <summary>A typed object variable going out of scope, or being cleared: its reference goes and it holds Nothing.</summary>
    public static void Release<T>(ref ObjectSlot<T> slot)
        where T : class => RuntimeObject.ReleaseInterface(slot.Exchange(0));

    /// <summary>The copy of a record takes a reference of its own for its object member (D18).</summary>
    public static void Retain<T>(ObjectSlot<T> slot)
        where T : class => RuntimeObject.AddRefInterface(slot.Pointer);

    /// <summary>Returning from a Function whose result is a typed object variable: the slot's reference moves to the caller's statement, as for any other result (ARCHITECTURE.md D18).</summary>
    public static T? Transfer<T>(ref ObjectSlot<T> slot)
        where T : class
    {
        var value = slot.Target;
        var pointer = slot.Exchange(0);
        if (pointer != 0)
        {
            pendingResult = new Temporary(null, pointer, IsInterface: true);
        }

        return value;
    }

    /// <summary>A typed object variable declared As New (MS-VBAL 5.2.3.1.1): any access creates the object when the variable holds Nothing.</summary>
    public static T AutoNew<T>(ref ObjectSlot<T> slot, Func<T> create)
        where T : class
    {
        if (slot.Target is { } existing)
        {
            return existing;
        }

        ArgumentNullException.ThrowIfNull(create);
        var created = create();
        Assign(ref slot, created);
        return created;
    }

    /// <summary>The cell for a ByRef object parameter of a late-bound call, which passes values: a slot holding a reference of its own, released when the current statement ends.</summary>
    public static ref ObjectSlot<T> Cell<T>(T? value)
        where T : class
    {
        var cell = new SlotCell<T> { Slot = Holding(value) };
        Temporaries.Add(new Temporary(cell, 0));
        return ref cell.Slot;
    }

    /// <summary>The interface pointer of an object, with a reference taken on it for a slot: a runtime object's block, made by that first reference if it had none, or a COM object's pointer.</summary>
    private static nint HeldPointer(object? value)
    {
        switch (value)
        {
            case null:
                return 0;
            case RuntimeObject runtime:
                return runtime.AddRefPointer();
            case IComInterface com:
                var pointer = com.InterfacePointer;
                RuntimeObject.AddRefInterface(pointer);
                return pointer;
            default:
                throw new InvalidOperationException(VbaErrors.Invariant($"A {value.GetType().Name} has no COM identity, so an object variable cannot hold it."));
        }
    }

    /// <summary>
    /// A store into a Variant slot: the slot takes what it must own (<see cref="Own(in Variant)"/>:
    /// a reference, a copy of a string, a copy of an array) and what it held before
    /// is released afterwards, so <c>s = s</c> is safe (D18, D20).
    /// </summary>
    public static void Assign(ref Variant slot, in Variant value)
    {
        if (slot.IsArray && slot.AsArray().IsLocked)
        {
            // A For Each is walking the array the slot holds, and VBA will not let it go (error 10).
            throw new VbaException(VbaErrors.ArrayFixedOrLocked);
        }

        var old = slot;
        slot = Own(value);
        ReleaseValue(old);
    }

    /// <summary>A store into a String slot (D20): the slot takes its own copy of the BSTR and frees the one it held, the copy first, so <c>s = s</c> is safe.</summary>
    public static void Assign(ref VbaString slot, VbaString value)
    {
        var copy = value.Copy();
        var old = slot;
        slot = copy;
        old.Free();
    }

    /// <summary>The copy a String store or a ByVal String parameter owns.</summary>
    public static VbaString Own(VbaString value) => value.Copy();

    /// <summary>A String local going out of scope, or a String slot being cleared: its BSTR is freed and the slot is null.</summary>
    public static void Release(ref VbaString slot)
    {
        var old = slot;
        slot = VbaString.Null;
        old.Free();
    }

    /// <summary>Returning a String result: the BSTR moves to the caller's statement rather than dying with the callee's frame (see <see cref="Transfer{T}"/>).</summary>
    public static VbaString Transfer(ref VbaString slot)
    {
        var value = slot;
        slot = VbaString.Null;
        if (!value.IsNull)
        {
            pendingResult = new Temporary(null, value.Pointer);
        }

        return value;
    }

    /// <summary>A ByVal object parameter holds its own reference for the length of the call.</summary>
    public static void Retain(object? value) => (value as IReferenceCounted)?.AddRef();

    /// <summary>
    /// The value a store must own: an object gains a reference, a string is copied into a new
    /// BSTR, an array is copied with elements of its own (MS-VBAL 5.5.1.2.5 Let-coercion to and
    /// from arrays), anything else passes through.
    /// For containers and copies that hold Variants outside <see cref="Assign(ref Variant, in Variant)"/>,
    /// and for a ByVal Variant parameter, which owns its own copy for the length of the call.
    /// </summary>
    public static Variant Own(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.String:
                return Variant.FromOwnedString(Bstr.Copy(value.StringPointer));
            case VarType.Object:
                RuntimeObject.AddRefInterface(value.InterfacePointer);
                return value;
            case VarType.Array:
                return Variant.FromArray(value.AsArray().Clone());
            default:
                return value;
        }
    }

    /// <summary>A local going out of scope, or a slot being cleared: the reference goes and the slot empties.</summary>
    public static void Release<T>(ref T? slot)
        where T : class
    {
        var old = slot;
        slot = null;
        (old as IReferenceCounted)?.Release();
    }

    /// <summary>A Variant slot being cleared: its object loses a reference, its string is freed, and its array's elements go (D18, D20).</summary>
    public static void Release(ref Variant slot)
    {
        var old = slot;
        slot = Variant.Empty;
        ReleaseValue(old);
    }

    /// <summary>A For Each enumerator going out of scope: what it held for the loop goes (a temporary array's elements, a reference on the collection).</summary>
    public static void Release(ref ForEachEnumerator? slot)
    {
        var old = slot;
        slot = null;
        old?.Finish();
    }

    /// <summary>A store into a declared array variable (MS-VBAL 5.5.1.2.5): the variable takes the copy, which must be its own, and the array it held is destroyed; one a For Each is walking raises error 10.</summary>
    public static void AssignArray(ref VbaArray slot, VbaArray value)
    {
        if (slot.IsLocked)
        {
            // The copy the store made goes with the refused store (Excel probe 2026-09-11: error 10, the array unchanged).
            value.Destroy();
            throw new VbaException(VbaErrors.ArrayFixedOrLocked);
        }

        var old = slot;
        slot = value;
        if (old.Descriptor != value.Descriptor)
        {
            old.Destroy();
        }
    }

    /// <summary>
    /// A host reading back what a callee left in a ByRef slot (an event handler's parameter, a
    /// UDF's): the value becomes a temporary of the current frame, released when the frame ends,
    /// and the slot empties. The value is handed back for the host to read meanwhile.
    /// </summary>
    public static Variant Adopt(ref Variant slot)
    {
        var value = slot;
        slot = Variant.Empty;
        if (value.InterfacePointer != 0)
        {
            Temporaries.Add(new Temporary(null, value.InterfacePointer, IsInterface: true));
        }
        else if (value.ArrayDescriptor != 0)
        {
            Temporaries.Add(ArrayTemporary(value.AsArray()));
        }
        else if (value.StringPointer != 0)
        {
            Temporaries.Add(new Temporary(null, value.StringPointer));
        }

        return value;
    }

    /// <summary>A record slot takes a copy (MS-VBAL 5.5.1.2.5): what the slot held before loses its references and strings, as the overwritten storage does in VBA (D20). The value is a copy of its own, never the slot's.</summary>
    public static void AssignRecord<T>(ref T slot, T value)
        where T : struct, IVbaRecord
    {
        var old = slot;
        slot = value;
        old.ReleaseReferences();
    }

    /// <summary>An array local going out of scope, or an array slot being cleared: the array is destroyed with what its elements own, and the slot is left unallocated. An array a For Each still walks is left to it.</summary>
    public static void Release(ref VbaArray array)
    {
        var old = array;
        array = VbaArray.Unallocated(old.ElementType);
        if (!old.IsLocked)
        {
            old.Destroy();
        }
    }

    /// <summary>
    /// A freshly created object, or an object a call returned: one reference owned by the current
    /// statement, dropped when the statement ends (<see cref="ReleaseTo"/>). Values that are not
    /// reference counted pass through untouched.
    /// </summary>
    public static T Owned<T>(T value)
        where T : class
    {
        if (value is IReferenceCounted counted)
        {
            counted.AddRef();
            Temporaries.Add(new Temporary(counted, 0));
        }

        return value;
    }

    public static Variant Owned(in Variant value)
    {
        var pointer = value.InterfacePointer;
        if (pointer != 0)
        {
            RuntimeObject.AddRefInterface(pointer);
            Temporaries.Add(new Temporary(null, pointer, IsInterface: true));
        }
        else if (value.ArrayDescriptor != 0)
        {
            Temporaries.Add(ArrayTemporary(value.AsArray()));
        }

        return value;
    }

    /// <summary>
    /// An interface pointer the caller already holds a reference on, such as QueryInterface's
    /// result for an object a COM call returned: the reference becomes the current statement's,
    /// released when it ends unless a store took one of its own (D18), and the Variant is its view.
    /// </summary>
    public static Variant OwnedInterface(nint dispatch)
    {
        if (dispatch == 0)
        {
            return Variant.Nothing;
        }

        Temporaries.Add(new Temporary(null, dispatch, IsInterface: true));
        return Variant.ViewInterface(dispatch);
    }

    /// <summary>
    /// A freshly created array (a Split, an Array(), a SAFEARRAY read from COM): the current
    /// statement destroys it when it ends, unless a For Each took it over (M7 C4).
    /// </summary>
    public static VbaArray Owned(VbaArray array)
    {
        if (array.IsAllocated)
        {
            Temporaries.Add(ArrayTemporary(array));
        }

        return array;
    }

    /// <summary>
    /// A For Each over an array the statement owns takes it for the loop's length: the array leaves the statement's
    /// temporaries. False when the array is a variable's, or a temporary below <paramref name="mark"/>, which an
    /// earlier frame's statement owns (the ParamArray a caller passed).
    /// </summary>
    internal static bool Detach(VbaArray array, int mark)
    {
        var list = temporaries;
        if (list is null)
        {
            return false;
        }

        for (var i = list.Count - 1; i >= Math.Max(mark, 0); i--)
        {
            if (list[i].ArrayElement != VarType.Empty && list[i].Pointer == array.Descriptor)
            {
                list.RemoveAt(i);
                return true;
            }
        }

        return false;
    }

    /// <summary>A BSTR the runtime just allocated for a value: the current statement frees it when it ends unless a store copied it first.</summary>
    internal static void OwnedString(nint bstr) => Temporaries.Add(new Temporary(null, bstr));

    /// <summary>The temporary stack's depth, taken before a statement or a procedure body runs.</summary>
    public static int Mark() => temporaries?.Count ?? 0;

    /// <summary>
    /// Drops the temporaries created since the mark, newest first, and hands a result a callee
    /// transferred (<see cref="Transfer{T}"/>) to the caller's statement.
    /// </summary>
    public static void ReleaseTo(int mark)
    {
        var list = temporaries;
        if (list is not null)
        {
            while (list.Count > mark)
            {
                var last = list[^1];
                list.RemoveAt(list.Count - 1);
                last.Release();
            }
        }

        if (pendingResult is { } result)
        {
            pendingResult = null;
            Temporaries.Add(result);
        }
    }

    /// <summary>
    /// A String a callback returns to a native caller (AddressOf; Declares golden): the BSTR leaves
    /// the statement it would have died with, and the runtime's live set, since the caller owns it
    /// from here, as VBA hands it over.
    /// </summary>
    public static nint HandOut(VbaString result)
    {
        var list = Temporaries;
        for (var i = list.Count - 1; i >= 0; i--)
        {
            if (list[i] is { Reference: null, IsInterface: false, ArrayElement: VarType.Empty } temporary && temporary.Pointer == result.Pointer)
            {
                list.RemoveAt(i);
                break;
            }
        }

        Bstr.Abandoned(result.Pointer);
        return result.Pointer;
    }

    /// <summary>
    /// A Variant a callback returns to a native caller (AddressOf; ROADMAP.md M7 E14; Declares
    /// golden): x64 returns a VARIANT through a hidden pointer the caller passes first, so the
    /// value is written there and the pointer handed back. What it holds leaves the statement it
    /// would have died with, and a String or an array the runtime's live set, since the caller
    /// owns it from here, as VBA hands it over.
    /// </summary>
    public static unsafe nint HandOut(in Variant result, nint target)
    {
        var payload = result.InterfacePointer != 0 ? result.InterfacePointer : result.ArrayDescriptor != 0 ? result.ArrayDescriptor : result.StringPointer;
        if (payload != 0)
        {
            if (pendingResult is { Reference: null } pending && pending.Pointer == payload)
            {
                pendingResult = null;
            }
            else
            {
                var list = Temporaries;
                for (var i = list.Count - 1; i >= 0; i--)
                {
                    if (list[i] is { Reference: null } temporary && temporary.Pointer == payload)
                    {
                        list.RemoveAt(i);
                        break;
                    }
                }
            }

            if (result.ArrayDescriptor != 0)
            {
                VbaArray.Abandoned(result.ArrayDescriptor);
            }
            else if (result.InterfacePointer == 0)
            {
                Bstr.Abandoned(result.StringPointer);
            }
        }

        *(Variant*)target = result;
        return target;
    }

    /// <summary>The count of temporaries the current thread holds; zero between host calls when every statement released its own.</summary>
    public static int PendingTemporaries => temporaries?.Count ?? 0;

    /// <summary>Releases the temporaries of a condition or a loop header after it was evaluated, passing the value through.</summary>
    public static T After<T>(int mark, T value)
    {
        ReleaseTo(mark);
        return value;
    }

    /// <summary>
    /// Returning from a Function or Property Get: the result variable's reference moves to the
    /// caller's statement rather than dying with the callee's frame. The frame's
    /// <see cref="ReleaseTo"/>, which runs after the return value is computed, completes the move.
    /// </summary>
    public static T? Transfer<T>(ref T? slot)
        where T : class
    {
        var value = slot;
        slot = null;
        switch (value)
        {
            case IReferenceCounted:
                pendingResult = new Temporary(value, 0);
                break;
        }

        return value;
    }

    /// <summary>Returning a record from a Function: what its fields own moves to the caller's statement, in a box the statement releases (D18; ROADMAP.md M7 C3).</summary>
    public static T TransferRecord<T>(ref T slot)
        where T : struct, IVbaRecord
    {
        var value = slot;
        pendingResult = new Temporary(value, 0);
        return value;
    }

    /// <summary>Returning an array from a Function: the array moves to the caller's statement, which destroys it unless a store copied it (D18; M7 C4).</summary>
    public static VbaArray Transfer(ref VbaArray slot)
    {
        var value = slot;
        slot = VbaArray.Unallocated(value.ElementType);
        if (value.IsAllocated)
        {
            pendingResult = ArrayTemporary(value);
        }

        return value;
    }

    public static Variant Transfer(ref Variant slot)
    {
        var value = slot;
        slot = Variant.Empty;
        if (value.InterfacePointer != 0)
        {
            pendingResult = new Temporary(null, value.InterfacePointer, IsInterface: true);
        }
        else if (value.ArrayDescriptor != 0)
        {
            pendingResult = ArrayTemporary(value.AsArray());
        }
        else if (value.StringPointer != 0)
        {
            pendingResult = new Temporary(null, value.StringPointer);
        }

        return value;
    }

    /// <summary>A variable declared As New (MS-VBAL 5.2.3.1.1): any access creates the object when the variable holds Nothing, Is Nothing included (Classes golden).</summary>
    public static T AutoNew<T>(ref T? slot, Func<T> create)
        where T : class
    {
        if (slot is not null)
        {
            return slot;
        }

        ArgumentNullException.ThrowIfNull(create);
        var created = create();
        Assign(ref slot, created);
        return created;
    }

    /// <summary>A host entry point (a command, a UDF, an event, a test): a frame whose temporaries go when it ends.</summary>
    public static ObjectFrame Frame() => new(Mark());

    /// <summary>Set on a WithEvents variable (MS-VBAL 5.2.3.1.4): the old source stops calling the sink's handlers and the new one starts.</summary>
    public static void Subscribe<T>(ref T? slot, T? value, IVbaEventSink sink, string name)
        where T : class
    {
        if (ReferenceEquals(slot, value))
        {
            return;
        }

        (slot as VbaClassObject)?.Unsubscribe(sink, name);
        Assign(ref slot, value);
        (value as VbaClassObject)?.Subscribe(sink, name);
    }

    /// <summary>
    /// Set on a WithEvents variable of a library type (MS-VBAL 5.2.3.1.4): the advisory on the old
    /// object is dropped and one on the new object taken, kept in <paramref name="advisory"/>
    /// until the next Set or the owner's end. An object that does not raise COM events (the
    /// in-process hosts have none) is assigned without one.
    /// </summary>
    public static void Subscribe<T>(ref T? slot, T? value, IVbaEventSink sink, string name, ComEventInterface events, ref IDisposable? advisory)
        where T : class
    {
        ArgumentNullException.ThrowIfNull(events);
        if (ReferenceEquals(slot, value))
        {
            return;
        }

        advisory?.Dispose();
        advisory = null;
        Assign(ref slot, value);
        if (value is IComEventSource source)
        {
            advisory = source.Advise(events, sink, name);
        }
    }

    /// <summary>
    /// The argument at a position of a late-bound call. A required parameter the caller left out
    /// raises 449, as VBA does, rather than running the procedure without it (Classes golden: v = s
    /// for an object whose default member takes a key); an optional one never reaches here then.
    /// </summary>
    public static Variant At(ReadOnlySpan<Variant> arguments, int index) =>
        index < arguments.Length ? arguments[index] : throw new VbaException(VbaErrors.ArgumentNotOptional);

    /// <summary>The arguments from a position on, for a ParamArray parameter.</summary>
    public static ReadOnlySpan<Variant> Rest(ReadOnlySpan<Variant> arguments, int index) =>
        index < arguments.Length ? arguments[index..] : [];

    /// <summary>
    /// A storage location for a ByRef parameter of a late-bound call, which passes values
    /// (ROADMAP.md backlog): the cell owns a copy of the value, since the callee may reassign the
    /// parameter, which releases what the cell held (D18, D20), and the current statement releases
    /// the cell when it ends.
    /// </summary>
    public static ref Variant Slot(in Variant value)
    {
        var cell = new Variant[1];
        cell[0] = Own(value);
        Temporaries.Add(new Temporary(cell, 0));
        return ref cell[0];
    }

    public static ref VbaString Slot(VbaString value)
    {
        var cell = new VbaString[1];
        cell[0] = value.Copy();
        Temporaries.Add(new Temporary(cell, 0));
        return ref cell[0];
    }

    /// <summary>The cell for an array parameter holds a copy of the array (MS-VBAL 5.5.1.2.5), which the statement destroys when it ends.</summary>
    public static ref VbaArray Slot(VbaArray value)
    {
        var cell = new VbaArray[1];
        cell[0] = value.Clone();
        Temporaries.Add(new Temporary(cell, 0));
        return ref cell[0];
    }

    /// <summary>The cell for any other parameter type: a scalar, or an object (the cell holds a reference of its own).</summary>
    public static ref T Slot<T>(T value)
    {
        var cell = new T[1];
        cell[0] = value;
        switch (value)
        {
            case IReferenceCounted counted:
                counted.AddRef();
                Temporaries.Add(new Temporary(cell, 0));
                break;
        }

        return ref cell[0];
    }

    /// <summary>The enumerator a class's NewEnum member returns, as For Each walks it (MS-VBAL 5.6.13.2).</summary>
    public static IEnumerator<Variant> Enumerator(in Variant value)
    {
        if (value.IsArray)
        {
            var array = value.AsArray();
            var elements = new Variant[array.Count];
            for (var i = 0; i < elements.Length; i++)
            {
                elements[i] = array.ElementAt(i);
            }

            return elements.AsEnumerable().GetEnumerator();
        }

        return value.IsObject
            ? value.AsObject() switch
            {
                IEnumerable<Variant> enumerable => enumerable.GetEnumerator(),
                IDispatchObject dispatch => dispatch.Enumerate(),
                null => throw VbaErrors.ObjectVariableNotSet(),
                _ => throw VbaErrors.TypeMismatch(),
            }
            : throw VbaErrors.TypeMismatch();
    }

    /// <summary>TypeOf value Is T (MS-VBAL 5.6.9.10) for a class of the project or the runtime.</summary>
    public static bool IsOfType<T>(in Variant value)
        where T : class =>
        value.IsObject && value.AsObject() is T;

    /// <summary>TypeOf value Is Object: any object reference other than Nothing.</summary>
    public static bool IsAnyObject(in Variant value) => value.IsObject && !value.IsNothing;

    /// <summary>TypeOf value Is a type library's type (MS-VBAL 5.6.9.10): the object answers QueryInterface for that type's interface, as VBA asks it; Nothing raises 91 (Objects golden).</summary>
    public static bool IsOfInterface(in Variant value, in Guid iid)
    {
        if (!value.IsObject)
        {
            return false;
        }

        var pointer = value.InterfacePointer;
        if (pointer == 0)
        {
            throw VbaErrors.ObjectVariableNotSet();
        }

        if (Marshal.QueryInterface(pointer, in iid, out var result) < 0 || result == 0)
        {
            return false;
        }

        Marshal.Release(result);
        return true;
    }

    /// <summary>The temporary that destroys an array when its statement ends.</summary>
    private static Temporary ArrayTemporary(VbaArray array) => new(null, array.Descriptor, ArrayElement: array.ElementType);

    /// <summary>What a slot held loses what it owned: an object reference, a BSTR, an array's elements (D18, D20).</summary>
    private static void ReleaseValue(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.String:
                Bstr.Free(value.StringPointer);
                break;
            case VarType.Object:
                RuntimeObject.ReleaseInterface(value.InterfacePointer);
                break;
            case VarType.Array:
                {
                    var array = value.AsArray();
                    Release(ref array);
                    break;
                }
        }
    }
}

/// <summary>The temporaries of a host-driven call, released when the call ends (<see cref="ObjectRefs.Frame"/>).</summary>
public readonly struct ObjectFrame(int mark) : IDisposable
{
    public void Dispose() => ObjectRefs.ReleaseTo(mark);
}

/// <summary>
/// The enumerator object <c>Collection.[_NewEnum]</c> hands to a class's <c>NewEnum</c> (dispid -4),
/// which <c>For Each</c> then walks; VBA types it IUnknown. A COM caller walks it through
/// IEnumVARIANT, its second face (<see cref="RuntimeObject"/>): Next copies items out, Skip passes
/// them, Reset starts again, and Clone goes on from the same place by itself (Declares golden).
/// </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1001:Types that own disposable fields should be disposable", Justification = "GetEnumerator hands the walk out and For Each disposes what it walks; this object may outlive the walk (Classes golden).")]
public sealed unsafe class VbaEnumerator : RuntimeObject, IVbaObject, IEnumerable<Variant>
{
    private readonly Func<IEnumerator<Variant>>? restart;
    private Walk walk;

    public VbaEnumerator(IEnumerator<Variant> enumerator)
        : this(enumerator, null)
    {
    }

    /// <summary>An enumerator that <paramref name="restart"/> begins again, for IEnumVARIANT's Reset and Clone.</summary>
    internal VbaEnumerator(IEnumerator<Variant> enumerator, Func<IEnumerator<Variant>>? restart)
    {
        this.restart = restart;
        walk = new Walk(enumerator);
    }

    public string TypeName => "Unknown";

    /// <summary>
    /// The enumerator has a COM identity like any object (ARCHITECTURE.md section 5, "Objects"). Its
    /// last reference going leaves the walk alone: For Each keeps the walk <see cref="GetEnumerator"/>
    /// handed out, which outlives this object when a class's NewEnum returns it (Classes golden).
    /// </summary>
    protected override void OnLastRelease()
    {
    }

    public IEnumerator<Variant> GetEnumerator() => walk;

    IEnumerator IEnumerable.GetEnumerator() => walk;

    /// <summary>IEnumVARIANT::Next: up to <paramref name="count"/> items copied into the caller's VARIANTs, which the caller owns from then on; the number copied.</summary>
    internal uint CopyNext(uint count, Variant* items)
    {
        uint copied = 0;
        while (copied < count && walk.MoveNext())
        {
            var item = walk.Current;
            var target = items + copied;
            *target = default;
            var hr = SafeArrayNative.VariantCopy(target, &item);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            copied++;
        }

        return copied;
    }

    /// <summary>IEnumVARIANT::Skip: the number of items passed, fewer than asked at the end.</summary>
    internal uint Skip(uint count)
    {
        uint skipped = 0;
        while (skipped < count && walk.MoveNext())
        {
            skipped++;
        }

        return skipped;
    }

    /// <summary>IEnumVARIANT::Reset: the walk begins again; false when nothing can begin it.</summary>
    internal bool Restart()
    {
        if (restart is null)
        {
            return false;
        }

        walk = new Walk(restart());
        return true;
    }

    /// <summary>IEnumVARIANT::Clone: an enumerator of its own at the same place; null when nothing can begin one.</summary>
    internal VbaEnumerator? Clone()
    {
        if (restart is null)
        {
            return null;
        }

        var clone = new VbaEnumerator(restart(), restart);
        for (var i = 0; i < walk.Position && clone.walk.MoveNext(); i++)
        {
        }

        return clone;
    }

    // The walk For Each and IEnumVARIANT share, counting its steps so a clone can start where it stands.
    private sealed class Walk(IEnumerator<Variant> inner) : IEnumerator<Variant>
    {
        public int Position { get; private set; }

        public Variant Current => inner.Current;

        object IEnumerator.Current => inner.Current;

        public bool MoveNext()
        {
            if (!inner.MoveNext())
            {
                return false;
            }

            Position++;
            return true;
        }

        public void Reset() => inner.Reset();

        public void Dispose() => inner.Dispose();
    }
}
