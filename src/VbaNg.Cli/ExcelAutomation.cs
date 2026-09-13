using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace VbaNg.Cli;

/// <summary>
/// COM automation of the running Excel instance. This is the CLI's only channel to the add-in
/// (ARCHITECTURE.md D6): find Excel, call <c>Application.Run</c>, and retry while Excel reports
/// itself busy.
/// </summary>
[SupportedOSPlatform("windows")]
internal static partial class ExcelAutomation
{
    private const int RpcServerCallRetryLater = unchecked((int)0x8001010A);
    private const int RpcCallRejected = unchecked((int)0x80010001);
    private const int BusyRetryCount = 40;
    private const uint ObjectIdNativeOm = 0xFFFFFFF0;
    private static readonly TimeSpan BusyRetryInterval = TimeSpan.FromMilliseconds(250);
    private static readonly Guid DispatchInterfaceId = new("00020400-0000-0000-C000-000000000046");

    [LibraryImport("ole32.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int CLSIDFromProgID(string progId, out Guid clsid);

    [LibraryImport("oleaut32.dll")]
    private static partial int GetActiveObject(in Guid clsid, nint reserved, out nint unknown);

    [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial nint FindWindowEx(nint parent, nint childAfter, string? className, string? windowName);

    [LibraryImport("oleacc.dll")]
    private static partial int AccessibleObjectFromWindow(nint hwnd, uint objectId, in Guid interfaceId, out nint dispatch);

    /// <summary>The running Excel application object, or null when no reachable instance exists.</summary>
    public static object? GetRunningExcel() => FromRunningObjectTable() ?? FromWorkbookWindow();

    /// <summary>Calls <c>Application.Run(macro, arguments...)</c>, retrying while Excel is busy.</summary>
    public static object? Run(object application, string macro, params object[] arguments)
    {
        ArgumentNullException.ThrowIfNull(application);
        ArgumentException.ThrowIfNullOrWhiteSpace(macro);
        ArgumentNullException.ThrowIfNull(arguments);

        var invokeArguments = new object[arguments.Length + 1];
        invokeArguments[0] = macro;
        arguments.CopyTo(invokeArguments, 1);

        for (var attempt = 0; ; attempt++)
        {
            try
            {
                return application.GetType().InvokeMember(
                    "Run",
                    BindingFlags.InvokeMethod,
                    binder: null,
                    application,
                    invokeArguments,
                    CultureInfo.InvariantCulture);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is COMException inner)
            {
                if (!IsBusy(inner) || attempt >= BusyRetryCount)
                {
                    ExceptionDispatchInfo.Capture(inner).Throw();
                }
            }
            catch (COMException ex) when (IsBusy(ex) && attempt < BusyRetryCount)
            {
            }

            Thread.Sleep(BusyRetryInterval);
        }
    }

    private static object? FromRunningObjectTable()
    {
        Marshal.ThrowExceptionForHR(CLSIDFromProgID("Excel.Application", out var clsid));
        if (GetActiveObject(in clsid, 0, out var unknown) != 0)
        {
            return null;
        }

        try
        {
            return Marshal.GetObjectForIUnknown(unknown);
        }
        finally
        {
            Marshal.Release(unknown);
        }
    }

    /// <summary>
    /// Excel registers in the Running Object Table only after its main window first loses focus
    /// (Microsoft KB 238610), so also walk the window tree XLMAIN > XLDESK > EXCEL7 and reach the
    /// native object model through the accessibility API. Needs at least one open workbook window.
    /// </summary>
    private static object? FromWorkbookWindow()
    {
        var main = nint.Zero;
        while ((main = FindWindowEx(nint.Zero, main, "XLMAIN", null)) != nint.Zero)
        {
            var desk = FindWindowEx(main, nint.Zero, "XLDESK", null);
            if (desk == nint.Zero)
            {
                continue;
            }

            var book = FindWindowEx(desk, nint.Zero, "EXCEL7", null);
            if (book == nint.Zero)
            {
                continue;
            }

            if (AccessibleObjectFromWindow(book, ObjectIdNativeOm, in DispatchInterfaceId, out var dispatch) != 0)
            {
                continue;
            }

            object window;
            try
            {
                window = Marshal.GetObjectForIUnknown(dispatch);
            }
            finally
            {
                Marshal.Release(dispatch);
            }

            try
            {
                return window.GetType().InvokeMember(
                    "Application",
                    BindingFlags.GetProperty,
                    binder: null,
                    window,
                    args: null,
                    CultureInfo.InvariantCulture);
            }
            finally
            {
                Marshal.FinalReleaseComObject(window);
            }
        }

        return null;
    }

    private static bool IsBusy(COMException exception) =>
        exception.HResult is RpcServerCallRetryLater or RpcCallRejected;
}
