using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>
/// An object the runtime implements, with a COM identity of its own (ARCHITECTURE.md section 5,
/// "Class modules" and "Objects"; D18, D20): a class instance or a Collection. Its identity is a
/// block in native memory whose first word is an <c>IDispatch</c> vtable, followed by the
/// reference count, a handle to this object, and, for an enumerator, an <c>IEnumVARIANT</c>
/// vtable, the object's second face. The count is D18's: generated code's
/// <see cref="AddRef"/> and <see cref="Release"/> and a COM caller's go to the same number, and
/// the object acts on its last release whoever makes it. The handle keeps the object alive while
/// the count is above zero, so a pointer that only COM holds is enough. The block is made with
/// the first reference and freed after the last. The vtable's entries are
/// <c>[UnmanagedCallersOnly]</c> thunks here in the runtime, never in a project's collectible
/// assembly; a project reset destroys the project's remaining instances and cuts their blocks off,
/// so a COM caller still holding one gets RPC_E_DISCONNECTED rather than a call into unloaded code.
/// </summary>
public abstract unsafe class RuntimeObject : IReferenceCounted
{
    private const int SOk = 0;
    private const int SFalse = 1;
    private const int ENotImpl = unchecked((int)0x80004001);
    private const int ENoInterface = unchecked((int)0x80004002);
    private const int EPointer = unchecked((int)0x80004003);
    private const int EUnexpected = unchecked((int)0x8000FFFF);
    private const int DispEMemberNotFound = unchecked((int)0x80020003);
    private const int DispEUnknownName = unchecked((int)0x80020006);
    private const int RpcEDisconnected = unchecked((int)0x80010108);
    private const int BlockSize = 32;
    private const int CountOffset = 8;
    private const int HandleOffset = 16;
    private const int EnumOffset = 24;

    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly Guid IidIDispatch = new("00020400-0000-0000-C000-000000000046");
    private static readonly Guid IidIEnumVariant = new("00020404-0000-0000-C000-000000000046");
    private static readonly ConcurrentDictionary<nint, byte> Live = new();
    private static readonly nint Vtable = BuildVtable();
    private static readonly nint EnumVtable = BuildEnumVtable();

    private nint block;
    private bool lastReleasing;
    private bool constructing;

    /// <summary>References VBA code and COM callers hold; 0 before the first and after the last.</summary>
    public int References => block == 0 ? 0 : Volatile.Read(ref Count(block));

    /// <summary>Blocks alive in the process: objects with a reference, not yet cut off by a project reset.</summary>
    public static int LiveBlocks => Live.Count;

    public void AddRef()
    {
        if (block == 0)
        {
            Allocate();
        }

        Interlocked.Increment(ref Count(block));
    }

    /// <summary>A reference goes; the last one runs <see cref="OnLastRelease"/> and frees the block. An object never referenced has nothing to release.</summary>
    public void Release()
    {
        var current = block;
        if (current == 0)
        {
            return;
        }

        int count;
        do
        {
            count = Volatile.Read(ref Count(current));
            if (count <= 0)
            {
                return;
            }
        }
        while (Interlocked.CompareExchange(ref Count(current), count - 1, count) != count);

        if (count != 1 || lastReleasing || constructing)
        {
            // Not the last reference, a reference taken and dropped by the last release's own work (Class_Terminate passing Me
            // along), or one taken and dropped while the object is being built (Class_Initialize passing Me along).
            return;
        }

        lastReleasing = true;
        try
        {
            OnLastRelease();
        }
        finally
        {
            lastReleasing = false;
            if (block == current && Volatile.Read(ref Count(current)) == 0)
            {
                FreeBlock();
            }
        }
    }

    /// <summary>
    /// The object is being built: generated code runs Class_Initialize between this and <see cref="EndConstruction"/>. The
    /// object counts as referenced meanwhile, as New holds it in VBA, so a reference taken and dropped (TypeName(Me)) does
    /// not end it (Classes golden).
    /// </summary>
    protected void BeginConstruction() => constructing = true;

    /// <summary>The object is built: an identity no one kept a reference to goes, and the reference New hands out takes a new one.</summary>
    protected void EndConstruction()
    {
        constructing = false;
        if (block != 0 && Volatile.Read(ref Count(block)) == 0)
        {
            FreeBlock();
        }
    }

    /// <summary>The object as an <c>IDispatch</c> pointer COM may keep, with a reference of its own for the callee to release.</summary>
    public nint AddRefPointer()
    {
        AddRef();
        return block;
    }

    /// <summary>The runtime object behind a pointer COM handed back, or null when the pointer is not one of the runtime's blocks or was cut off by a reset.</summary>
    public static RuntimeObject? FromPointer(nint pointer) => pointer != 0 && *(nint*)pointer == Vtable ? Target(pointer) : null;

    /// <summary>The object's IDispatch pointer without a new reference, what a Variant holds for it; 0 while nothing references it.</summary>
    internal nint InterfacePointer => block;

    /// <summary>IUnknown::AddRef on an interface pointer a Variant holds (D18, D20): a runtime object's count directly, any other object through its vtable.</summary>
    internal static void AddRefInterface(nint pointer)
    {
        if (pointer == 0)
        {
            return;
        }

        if (*(nint*)pointer == Vtable && Target(pointer) is { } target)
        {
            target.AddRef();
            return;
        }

        _ = ((delegate* unmanaged[Stdcall]<nint, uint>)(*(nint**)pointer)[1])(pointer);
    }

    /// <summary>IUnknown::Release on an interface pointer a Variant held, on the thread that made it, as the statement or the storage that owned it ends (D18).</summary>
    internal static void ReleaseInterface(nint pointer)
    {
        if (pointer == 0)
        {
            return;
        }

        if (*(nint*)pointer == Vtable && Target(pointer) is { } target)
        {
            target.Release();
            return;
        }

        _ = ((delegate* unmanaged[Stdcall]<nint, uint>)(*(nint**)pointer)[2])(pointer);
    }

    /// <summary>The object's last reference went: a class instance runs Class_Terminate and releases its storage; a Collection releases its items.</summary>
    protected abstract void OnLastRelease();

    /// <summary>What a COM caller reaches through IDispatch on an object that is no dispatch object to VBA code (Collection, Err): its members, as late binding answers them (ROADMAP.md M7 F); null for none.</summary>
    internal virtual IDispatchObject? ComMembers => null;

    /// <summary>A project reset destroys the object whatever its count (docs/vba-quirks.md): a class instance releases its storage without Class_Terminate.</summary>
    internal virtual void Destroy()
    {
    }

    /// <summary>
    /// The reset of <paramref name="project"/> destroys every instance of its classes still alive
    /// (in a reference cycle, or held by a COM caller), as VBA destroys every object of a project it
    /// resets: first each releases its storage, which may bring others to their last release, then
    /// the blocks still referenced are cut off from their objects.
    /// </summary>
    internal static void DestroyInstances(Assembly project)
    {
        var doomed = new List<(nint Block, RuntimeObject Target)>();
        foreach (var memory in Live.Keys)
        {
            if (Target(memory) is { } target && target.GetType().Assembly == project)
            {
                doomed.Add((memory, target));
            }
        }

        foreach (var (_, target) in doomed)
        {
            target.Destroy();
        }

        foreach (var (memory, _) in doomed)
        {
            // A block the destruction already freed is no longer in the table; the rest are cut off and freed with their last reference.
            if (Live.TryRemove(memory, out _))
            {
                var handle = *(nint*)(memory + HandleOffset);
                *(nint*)(memory + HandleOffset) = 0;
                if (handle != 0)
                {
                    GCHandle.FromIntPtr(handle).Free();
                }
            }
        }
    }

    private static ref int Count(nint memory) => ref *(int*)(memory + CountOffset);

    private static RuntimeObject? Target(nint memory)
    {
        var handle = *(nint*)(memory + HandleOffset);
        return handle == 0 ? null : GCHandle.FromIntPtr(handle).Target as RuntimeObject;
    }

    private void Allocate()
    {
        var memory = (nint)NativeMemory.AllocZeroed(BlockSize);
        *(nint*)memory = Vtable;
        *(nint*)(memory + HandleOffset) = GCHandle.ToIntPtr(GCHandle.Alloc(this));
        if (this is VbaEnumerator)
        {
            *(nint*)(memory + EnumOffset) = EnumVtable;
        }

        block = memory;
        Live[memory] = 0;
    }

    private void FreeBlock()
    {
        var memory = block;
        block = 0;
        Live.TryRemove(memory, out _);
        var handle = *(nint*)(memory + HandleOffset);
        if (handle != 0)
        {
            GCHandle.FromIntPtr(handle).Free();
        }

        NativeMemory.Free((void*)memory);
    }

    private static nint BuildVtable()
    {
        var table = (nint*)NativeMemory.Alloc((nuint)(7 * sizeof(nint)));
        table[0] = (nint)(delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)&QueryInterface;
        table[1] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&ComAddRef;
        table[2] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&ComRelease;
        table[3] = (nint)(delegate* unmanaged[Stdcall]<nint, uint*, int>)&GetTypeInfoCount;
        table[4] = (nint)(delegate* unmanaged[Stdcall]<nint, uint, uint, nint*, int>)&GetTypeInfo;
        table[5] = (nint)(delegate* unmanaged[Stdcall]<nint, Guid*, char**, uint, uint, int*, int>)&GetIDsOfNames;
        table[6] = (nint)(delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, nint, nint, nint, uint*, int>)&Invoke;
        return (nint)table;
    }

    private static nint BuildEnumVtable()
    {
        var table = (nint*)NativeMemory.Alloc((nuint)(7 * sizeof(nint)));
        table[0] = (nint)(delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)&EnumQueryInterface;
        table[1] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&EnumAddRef;
        table[2] = (nint)(delegate* unmanaged[Stdcall]<nint, uint>)&EnumRelease;
        table[3] = (nint)(delegate* unmanaged[Stdcall]<nint, uint, Variant*, uint*, int>)&EnumNext;
        table[4] = (nint)(delegate* unmanaged[Stdcall]<nint, uint, int>)&EnumSkip;
        table[5] = (nint)(delegate* unmanaged[Stdcall]<nint, int>)&EnumReset;
        table[6] = (nint)(delegate* unmanaged[Stdcall]<nint, nint*, int>)&EnumClone;
        return (nint)table;
    }

    // The thunks never let an exception out: a managed exception crossing into a COM caller would take the process down.

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int QueryInterface(nint self, Guid* iid, nint* result) => QueryCore(self, iid, result);

    /// <summary>One identity whichever face is asked: IUnknown and IDispatch are the block itself, IEnumVARIANT an enumerator's fourth word.</summary>
    private static int QueryCore(nint memory, Guid* iid, nint* result)
    {
        if (result is null || iid is null)
        {
            return EPointer;
        }

        *result = 0;
        try
        {
            if (Target(memory) is not { } target)
            {
                return RpcEDisconnected;
            }

            nint face;
            if (*iid == IidIUnknown || *iid == IidIDispatch)
            {
                face = memory;
            }
            else if (*iid == IidIEnumVariant && target is VbaEnumerator)
            {
                face = memory + EnumOffset;
            }
            else
            {
                return ENoInterface;
            }

            target.AddRef();
            *result = face;
            return SOk;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint ComAddRef(nint self) => AddRefCore(self);

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint ComRelease(nint self) => ReleaseCore(self);

    private static uint AddRefCore(nint memory)
    {
        try
        {
            if (Target(memory) is { } target)
            {
                target.AddRef();
                return (uint)target.References;
            }

            return (uint)Interlocked.Increment(ref Count(memory));
        }
        catch (Exception)
        {
            return 0;
        }
    }

    private static uint ReleaseCore(nint memory)
    {
        try
        {
            if (Target(memory) is { } target)
            {
                target.Release();
                return (uint)target.References;
            }

            // A block a project reset cut off: only its count is left, and the memory goes with the last reference.
            var remaining = Interlocked.Decrement(ref Count(memory));
            if (remaining == 0)
            {
                NativeMemory.Free((void*)memory);
            }

            return (uint)Math.Max(remaining, 0);
        }
        catch (Exception)
        {
            return 0;
        }
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetTypeInfoCount(nint self, uint* count)
    {
        if (count is null)
        {
            return EPointer;
        }

        *count = 0;
        return SOk;
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetTypeInfo(nint self, uint index, uint lcid, nint* typeInfo)
    {
        if (typeInfo is not null)
        {
            *typeInfo = 0;
        }

        return ENotImpl;
    }

    /// <summary>GetIDsOfNames through the object's own dispid table (<see cref="IDispatchObject.GetDispId"/>); an unknown name or argument is DISP_E_UNKNOWNNAME with every id -1.</summary>
    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int GetIDsOfNames(nint self, Guid* iid, char** names, uint count, uint lcid, int* dispIds)
    {
        if (names is null || dispIds is null || count == 0)
        {
            return count == 0 ? SOk : EPointer;
        }

        for (var i = 0; i < count; i++)
        {
            dispIds[i] = -1;
        }

        try
        {
            var target = Target(self);
            var dispatch = target as IDispatchObject ?? target?.ComMembers;
            if (dispatch is null)
            {
                return target is null ? RpcEDisconnected : DispEUnknownName;
            }

            var name = new string(names[0]);
            int[] ids;
            try
            {
                if (count == 1)
                {
                    ids = [dispatch.GetDispId(name)];
                }
                else
                {
                    var arguments = new string[count - 1];
                    for (var i = 1; i < count; i++)
                    {
                        arguments[i - 1] = new string(names[i]);
                    }

                    ids = dispatch.GetDispIds(name, arguments);
                }
            }
            catch (VbaException)
            {
                return DispEUnknownName;
            }

            for (var i = 0; i < count && i < ids.Length; i++)
            {
                dispIds[i] = ids[i];
            }

            return SOk;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    /// <summary>IDispatch::Invoke: the Interop provider reads the DISPPARAMS, runs the member, and writes the result or the EXCEPINFO (<see cref="IComProvider.Invoke"/>).</summary>
    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int Invoke(nint self, int dispId, Guid* iid, uint lcid, ushort flags, nint parameters, nint result, nint exception, uint* argumentError)
    {
        try
        {
            return Target(self) switch
            {
                null => RpcEDisconnected,
                IDispatchObject dispatch => Com.Provider.Invoke(dispatch, dispId, flags, parameters, result, exception, (nint)argumentError),
                { ComMembers: { } members } => Com.Provider.Invoke(members, dispId, flags, parameters, result, exception, (nint)argumentError),
                _ => DispEMemberNotFound,
            };
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    // IEnumVARIANT, an enumerator's second face: self is the block's fourth word, so the block is EnumOffset below it.

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int EnumQueryInterface(nint self, Guid* iid, nint* result) => QueryCore(self - EnumOffset, iid, result);

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint EnumAddRef(nint self) => AddRefCore(self - EnumOffset);

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static uint EnumRelease(nint self) => ReleaseCore(self - EnumOffset);

    /// <summary>IEnumVARIANT::Next: up to <paramref name="count"/> items, each a copy the caller owns; S_FALSE when fewer remained.</summary>
    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int EnumNext(nint self, uint count, Variant* items, uint* fetched)
    {
        if (fetched is not null)
        {
            *fetched = 0;
        }

        if (items is null || (fetched is null && count != 1))
        {
            return EPointer;
        }

        try
        {
            if (Target(self - EnumOffset) is not VbaEnumerator target)
            {
                return RpcEDisconnected;
            }

            var copied = target.CopyNext(count, items);
            if (fetched is not null)
            {
                *fetched = copied;
            }

            return copied == count ? SOk : SFalse;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int EnumSkip(nint self, uint count)
    {
        try
        {
            return Target(self - EnumOffset) is not VbaEnumerator target ? RpcEDisconnected
                : target.Skip(count) == count ? SOk : SFalse;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int EnumReset(nint self)
    {
        try
        {
            return Target(self - EnumOffset) is not VbaEnumerator target ? RpcEDisconnected
                : target.Restart() ? SOk : ENotImpl;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }

    [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
    private static int EnumClone(nint self, nint* result)
    {
        if (result is null)
        {
            return EPointer;
        }

        *result = 0;
        try
        {
            if (Target(self - EnumOffset) is not VbaEnumerator target)
            {
                return RpcEDisconnected;
            }

            if (target.Clone() is not { } clone)
            {
                return ENotImpl;
            }

            *result = clone.AddRefPointer() + EnumOffset;
            return SOk;
        }
        catch (Exception)
        {
            return EUnexpected;
        }
    }
}

/// <summary>
/// The IDispatch face Collection and Err show a COM caller (ROADMAP.md M7 F): each member name
/// at its dispid, the default member at 0, a call or a read going through LateBound as a
/// late-bound call in VBA code goes, and a put through its Let or Set, so a COM caller and VBA
/// code reach the same members with the same errors. DISPID_NEWENUM (-4) hands a Collection's
/// caller its enumerator, which answers IEnumVARIANT.
/// </summary>
internal sealed class LateBoundMembers(RuntimeObject target, string typeName, (int DispId, string Name)[] members) : IDispatchObject
{
    private const int DispIdNewEnum = -4;

    public string TypeName => typeName;

    public int GetDispId(string name)
    {
        foreach (var (dispId, member) in members)
        {
            if (member.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return dispId;
            }
        }

        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
    }

    public Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments) =>
        dispId == DispIdNewEnum && target is Collection collection
            ? Variant.FromObject(collection.NewEnum())
            : LateBound.Get(Variant.FromObject(target), NameOf(dispId), arguments);

    public void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference)
    {
        if (asReference)
        {
            LateBound.Set(Variant.FromObject(target), NameOf(dispId), indices, value);
            return;
        }

        LateBound.Let(Variant.FromObject(target), NameOf(dispId), indices, value);
    }

    public IEnumerator<Variant> Enumerate() =>
        target is IEnumerable<Variant> items ? items.GetEnumerator() : throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);

    public bool IsSameObject(object? other) => ReferenceEquals(target, other);

    private string NameOf(int dispId)
    {
        foreach (var (id, member) in members)
        {
            if (id == dispId)
            {
                return member;
            }
        }

        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
    }
}
