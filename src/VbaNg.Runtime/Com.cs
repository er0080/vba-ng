using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>How a member is invoked; the values are the DISPATCH_* flags of <c>IDispatch::Invoke</c>.</summary>
[Flags]
public enum InvokeKind
{
    Method = 1,
    PropertyGet = 2,
    PropertyPut = 4,
    PropertyPutRef = 8,
}

/// <summary>
/// An object whose members are reached by dispid, the way VBA reaches COM objects through
/// <c>IDispatch</c> (ARCHITECTURE.md section 6, D7). The Interop project implements it for COM
/// objects; the runtime only knows this contract, so generated code and the late-bound helpers
/// never depend on COM themselves. Dispids follow OLE Automation: 0 is the default member,
/// -4 the enumerator.
/// </summary>
public interface IDispatchObject : IVbaObject
{
    /// <summary>The dispid of a member by name, case-insensitively; raises error 438 when the object has no such member.</summary>
    int GetDispId(string name);

    /// <summary>Calls a method or reads a property. Arguments are passed by value; the result is Empty for a call that returns nothing.</summary>
    Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments);

    /// <summary>
    /// The dispids of a member and of the named arguments a call passes to it, from one
    /// GetIDsOfNames (MS-VBAL 5.6.13.1): element 0 is the member's, the rest follow
    /// <paramref name="argumentNames"/>. Error 438 when the object has no such member, 448 when
    /// the member has no such argument. Objects that cannot answer for names raise 448.
    /// </summary>
    int[] GetDispIds(string name, ReadOnlySpan<string> argumentNames) => throw new VbaException(VbaErrors.NamedArgumentNotFound);

    /// <summary>
    /// <see cref="Invoke"/> with named arguments: <paramref name="arguments"/> holds the positional
    /// ones and then the named ones in <paramref name="namedDispIds"/> order.
    /// </summary>
    Variant InvokeNamed(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, ReadOnlySpan<int> namedDispIds) =>
        namedDispIds.IsEmpty ? Invoke(dispId, kind, arguments) : throw new VbaException(VbaErrors.NamedArgumentNotFound);

    /// <summary><see cref="Put"/> with named index arguments, laid out as in <see cref="InvokeNamed"/>.</summary>
    void PutNamed(int dispId, ReadOnlySpan<Variant> indices, ReadOnlySpan<int> namedDispIds, in Variant value, bool asReference)
    {
        if (!namedDispIds.IsEmpty)
        {
            throw new VbaException(VbaErrors.NamedArgumentNotFound);
        }

        Put(dispId, indices, value, asReference);
    }

    /// <summary>Assigns a property, with index arguments for indexed properties; <paramref name="asReference"/> is the Set form (DISPATCH_PROPERTYPUTREF).</summary>
    void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference);

    /// <summary>The object's enumerator for <c>For Each</c>; raises error 438 when it has none.</summary>
    IEnumerator<Variant> Enumerate();

    /// <summary>COM identity for the <c>Is</c> operator: true when both wrap the same underlying object.</summary>
    bool IsSameObject(object? other);
}

/// <summary>Creates and locates COM objects for <c>CreateObject</c>, <c>GetObject</c>, and <c>New</c> on COM classes (ARCHITECTURE.md section 6, "Creation").</summary>
public interface IComProvider
{
    /// <summary>CreateObject(class, [servername]); raises error 429 when the class cannot be created.</summary>
    object CreateObject(string progId, string? serverName);

    /// <summary>GetObject([pathname], [class]); raises error 429 when no running instance exists and 432 when the path cannot be bound.</summary>
    object GetObject(string? pathName, string? className);

    /// <summary>New on a COM class: an instance by CLSID; raises error 429 when the class cannot be created.</summary>
    object CreateInstance(Guid classId);

    /// <summary>
    /// The instance behind a library's application object (a coclass marked appobject, such as
    /// Excel's Global), whose members are in scope unqualified. The Excel add-in answers with the
    /// running Application; elsewhere a running instance is found or one is created.
    /// </summary>
    object GetAppObject(Guid libraryId, Guid classId);

    /// <summary>
    /// IDispatch::Invoke from COM on an object the runtime implements (<see cref="RuntimeObject"/>;
    /// ARCHITECTURE.md section 5, "Class modules"): the provider reads the DISPPARAMS at
    /// <paramref name="parameters"/> into Variants, runs the member, writes the result VARIANT at
    /// <paramref name="result"/>, and turns a run-time error into DISP_E_EXCEPTION with the
    /// EXCEPINFO at <paramref name="exception"/>. Returns the HRESULT; without the Interop
    /// provider there is no VARIANT marshaling, so E_NOTIMPL.
    /// </summary>
    int Invoke(IDispatchObject target, int dispId, ushort flags, nint parameters, nint result, nint exception, nint argumentError) => unchecked((int)0x80004001);

    /// <summary>
    /// A wrapper over an IDispatch pointer a Variant holds, for a call through it (D20: a Variant
    /// holds an object as its interface pointer). The wrapper takes no reference of its own until
    /// something stores it; without the Interop provider there is none, so error 430.
    /// </summary>
    IDispatchObject Wrap(nint dispatch) => throw new VbaException(VbaErrors.ClassDoesNotSupportAutomation);
}

/// <summary>
/// A wrapper over a COM interface pointer the runtime did not make (the Interop layer's COM
/// objects): the pointer a Variant holds for it, without a new reference (D20; ARCHITECTURE.md
/// section 6, "Variant"). Zero once the wrapper released its pointer, which reads as Nothing.
/// </summary>
public interface IComInterface
{
    nint InterfacePointer { get; }
}

/// <summary>The COM provider in effect: the Interop project's, installed by the host, or one that raises error 429 for everything.</summary>
public static class Com
{
    private static IComProvider provider = NoComProvider.Instance;

    public static IComProvider Provider
    {
        get => provider;
        set => provider = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>New on a COM class, from generated code.</summary>
    public static IDispatchObject CreateInstance(string classId) =>
        (IDispatchObject)provider.CreateInstance(Guid.Parse(classId));

    /// <summary>A library's application object, from generated code; cached by the generated module, which holds the reference returned here for as long as it is loaded (D18).</summary>
    public static IDispatchObject AppObject(string libraryId, string classId)
    {
        var application = (IDispatchObject)provider.GetAppObject(Guid.Parse(libraryId), Guid.Parse(classId));
        ObjectRefs.Retain(application);
        return application;
    }

    private sealed class NoComProvider : IComProvider
    {
        public static readonly NoComProvider Instance = new();

        public object CreateObject(string progId, string? serverName) => throw new VbaException(VbaErrors.CannotCreateObject);

        public object GetObject(string? pathName, string? className) => throw new VbaException(VbaErrors.CannotCreateObject);

        public object CreateInstance(Guid classId) => throw new VbaException(VbaErrors.CannotCreateObject);

        public object GetAppObject(Guid libraryId, Guid classId) => throw new VbaException(VbaErrors.CannotCreateObject);
    }
}
