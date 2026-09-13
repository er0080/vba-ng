using ExcelDna.Integration;

using VbaNg.Interop;
using VbaNg.Runtime;

namespace VbaNg.AddIn;

/// <summary>
/// The running Application as generated code reaches it (ROADMAP.md WP4). Every member goes to
/// the real object but the three Excel itself must answer inside a UDF: <c>Volatile</c>, which
/// COM accepts and silently ignores, and <c>Caller</c> and <c>ThisCell</c>, of which
/// <c>ThisCell</c> raises 1004 during a recalculation. Inside a UDF they are answered from the
/// call Excel is making, through the C API (xlfVolatile, xlfCaller).
/// Caller and ThisCell differ for an array formula, where Caller is the whole calling range and
/// ThisCell the one cell being evaluated; both answer from xlfCaller here, which is the same
/// cell for the single-cell formulas the alpha covers.
/// </summary>
internal sealed class ExcelApplication(ComObject application) : IDispatchObject, IComEventSource, IReferenceCounted, IComInterface
{
    private const int VolatileDispId = 788;
    private const int CallerDispId = 317;
    private const int ThisCellDispId = 1962;

    [ThreadStatic]
    private static int functionDepth;

    /// <summary>The running Application itself; every member but the three answered here is handed to it.</summary>
    private readonly ComObject inner = application;

    /// <summary>True while a UDF runs on this thread, when Excel answers for the three members.</summary>
    private static bool InFunction => functionDepth > 0;

    /// <summary>Marks the thread as inside a UDF until the returned scope is disposed.</summary>
    public static IDisposable RunningFunction() => new FunctionScope();

    public string TypeName => inner.TypeName;

    public nint InterfacePointer => inner.InterfacePointer;

    public void AddRef() => inner.AddRef();

    public void Release() => inner.Release();

    public int GetDispId(string name) => inner.GetDispId(name);

    public int[] GetDispIds(string name, ReadOnlySpan<string> argumentNames) => inner.GetDispIds(name, argumentNames);

    public IDisposable Advise(ComEventInterface events, IVbaEventSink sink, string variable) => inner.Advise(events, sink, variable);

    public IEnumerator<Variant> Enumerate() => inner.Enumerate();

    public bool IsSameObject(object? other) => inner.IsSameObject(other is ExcelApplication wrapper ? wrapper.inner : other);

    public void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference) =>
        inner.Put(dispId, indices, value, asReference);

    public void PutNamed(int dispId, ReadOnlySpan<Variant> indices, ReadOnlySpan<int> namedDispIds, in Variant value, bool asReference) =>
        inner.PutNamed(dispId, indices, namedDispIds, value, asReference);

    public Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments) =>
        Answer(dispId, kind, arguments, out var answer) ? answer : inner.Invoke(dispId, kind, arguments);

    public Variant InvokeNamed(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, ReadOnlySpan<int> namedDispIds) =>
        namedDispIds.IsEmpty && Answer(dispId, kind, arguments, out var answer)
            ? answer
            : inner.InvokeNamed(dispId, kind, arguments, namedDispIds);

    /// <summary>What Excel answers for the three members while a UDF runs; false leaves the call to COM.</summary>
    private static bool Answer(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments, out Variant answer)
    {
        answer = Variant.Empty;
        if (!InFunction)
        {
            return false;
        }

        switch (dispId)
        {
            case VolatileDispId when kind == InvokeKind.Method:
                // Application.Volatile [True]: the argument is optional, as in the type library.
                XlCall.Excel(XlCall.xlfVolatile, arguments.IsEmpty || Coerce.ToBoolean(arguments[0]));
                return true;
            case CallerDispId when kind == InvokeKind.PropertyGet && arguments.IsEmpty:
            case ThisCellDispId when kind == InvokeKind.PropertyGet:
                answer = CallerRange();
                return true;
            default:
                return false;
        }
    }

    /// <summary>The Range the formula being evaluated sits in, from the C API rather than the object model.</summary>
    private static Variant CallerRange() =>
        XlCall.Excel(XlCall.xlfCaller) is ExcelReference reference && HostCommands.Workbooks is { } workbooks
            ? workbooks.RangeOf(reference)
            : Variant.Empty;

    private sealed class FunctionScope : IDisposable
    {
        public FunctionScope() => functionDepth++;

        public void Dispose() => functionDepth--;
    }
}
