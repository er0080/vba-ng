using ExcelDna.Integration;

using VbaNg.Interop;
using VbaNg.Runtime;

namespace VbaNg.AddIn;

/// <summary>
/// The COM provider inside Excel: the application object of the Excel library is the running
/// <c>Application</c> the add-in lives in, so <c>ActiveSheet</c>, <c>Range</c>, and the other
/// unqualified members of Excel's global interface reach this instance (ARCHITECTURE.md section 6).
/// </summary>
internal sealed class ExcelComProvider : ComProvider
{
    private ComObject? application;
    private ExcelApplication? appObject;

    /// <summary>The running Application as a COM object, created on first use; the add-in holds a reference of its own, so VBA code cannot release it (ARCHITECTURE.md D18).</summary>
    public ComObject Application => application ??= Held(ComObject.FromRcw(ExcelDnaUtil.Application));

    private static ComObject Held(ComObject wrapper)
    {
        wrapper.AddRef();
        return wrapper;
    }

    public override object GetAppObject(Guid libraryId, Guid classId)
    {
        if (libraryId == TypeLibraryCache.ExcelLibraryId)
        {
            // Generated code reaches Application through the wrapper, which answers the members
            // Excel itself must answer inside a UDF (ROADMAP.md WP4).
            return appObject ??= new ExcelApplication(Application);
        }

        return base.GetAppObject(libraryId, classId);
    }

    /// <summary>
    /// A view over an IDispatch pointer: the running Application becomes the wrapper again, so
    /// that <c>Application.Volatile</c> and the rest reach it even though the object model answers
    /// the <c>Application</c> property (dispid 148) with the real object (ROADMAP.md WP4).
    /// </summary>
    public override IDispatchObject Wrap(nint dispatch) =>
        application is { } app && appObject is { } wrapper && dispatch == app.InterfacePointer
            ? wrapper
            : base.Wrap(dispatch);

    /// <summary>Drops the cached Application wrapper; the next use creates a fresh one.</summary>
    public void Release()
    {
        appObject = null;
        application?.Dispose();
        application = null;
    }
}
