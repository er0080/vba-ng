using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace VbaNg.E2E;

/// <summary>
/// A throwaway hidden Excel with the vba-ng add-in loaded (CLAUDE.md R17): started fresh through
/// COM, never the user's instance, given a scratch workbook so host commands have something to
/// run in, quit on dispose and killed if it will not quit. Every member must be called on the
/// STA thread that created it (see <see cref="Sta"/>).
/// </summary>
internal sealed partial class ExcelInstance : IDisposable
{
    private readonly object application;
    private readonly int processId;
    private volatile bool killed;
    private bool disposed;

    private ExcelInstance(object application, int processId)
    {
        this.application = application;
        this.processId = processId;
    }

    public int ProcessId => processId;

    /// <summary>Opens a workbook and returns it; the add-in's WorkbookOpen handling runs inside the call.</summary>
    public object OpenWorkbook(string path)
    {
        var workbooks = Get(application, "Workbooks")!;
        try
        {
            return Call(workbooks, "Open", path)!;
        }
        finally
        {
            Release(workbooks);
        }
    }

    /// <summary>The workbook Excel has active: the scratch one from <see cref="Start"/> until another is opened.</summary>
    public object ActiveWorkbook() => Get(application, "ActiveWorkbook") ?? throw new InvalidOperationException("Excel has no active workbook.");

    /// <summary>The Name of a workbook, the qualifier Application.Run wants in front of a macro.</summary>
    public static string Name(object workbook) => (string)Get(workbook, "Name")!;

    /// <summary>Imports a module file into a workbook's VBA project (VBE object model access must be trusted on the machine).</summary>
    public static void ImportModule(object workbook, string path)
    {
        var project = Get(workbook, "VBProject")!;
        try
        {
            var components = Get(project, "VBComponents")!;
            try
            {
                Release(Call(components, "Import", path));
            }
            finally
            {
                Release(components);
            }
        }
        finally
        {
            Release(project);
        }
    }

    /// <summary>Saves a workbook under a path in a file format (52: macro-enabled), so its Path is that folder.</summary>
    public static void SaveAs(object workbook, string path, int format) => Call(workbook, "SaveAs", path, format);

    /// <summary>Saves a workbook where it already is, so an imported module reaches the file.</summary>
    public static void Save(object workbook) => Call(workbook, "Save");

    /// <summary>The Excel processes running now, by id.</summary>
    public static HashSet<int> RunningInstances()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                ids.Add(process.Id);
            }
        }

        return ids;
    }

    /// <summary>
    /// Waits, at most <paramref name="timeout"/>, for the Excel processes that were not running in
    /// <paramref name="before"/> to exit: instances a library under test started with CreateObject and quit
    /// itself, which Excel takes a while to shut down. None is killed.
    /// </summary>
    public static void WaitForOtherInstances(HashSet<int> before, TimeSpan timeout)
    {
        ArgumentNullException.ThrowIfNull(before);
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline && RunningInstances().Any(id => !before.Contains(id)))
        {
            Thread.Sleep(1000);
        }
    }

    /// <summary>The Value of a cell on a sheet of a workbook, as .NET's COM interop returns it.</summary>
    public static object? CellValue(object workbook, string sheet, string address)
    {
        var sheets = Get(workbook, "Worksheets")!;
        try
        {
            var worksheet = Get(sheets, "Item", sheet)!;
            try
            {
                var range = Get(worksheet, "Range", address)!;
                try
                {
                    return Get(range, "Value");
                }
                finally
                {
                    Release(range);
                }
            }
            finally
            {
                Release(worksheet);
            }
        }
        finally
        {
            Release(sheets);
        }
    }

    public static void SetCellValue(object workbook, string sheet, string address, object value)
    {
        var sheets = Get(workbook, "Worksheets")!;
        try
        {
            var worksheet = Get(sheets, "Item", sheet)!;
            try
            {
                var range = Get(worksheet, "Range", address)!;
                try
                {
                    Set(range, "Value", value);
                }
                finally
                {
                    Release(range);
                }
            }
            finally
            {
                Release(worksheet);
            }
        }
        finally
        {
            Release(sheets);
        }
    }

    /// <summary>Presses an ActiveX command button by setting its Value, which fires its Click event as VBA documents.</summary>
    public static void ClickActiveXButton(object workbook, string sheet, string name)
    {
        var sheets = Get(workbook, "Worksheets")!;
        try
        {
            var worksheet = Get(sheets, "Item", sheet)!;
            try
            {
                var oleObject = Get(worksheet, "OLEObjects", name)!;
                try
                {
                    var button = Get(oleObject, "Object")!;
                    try
                    {
                        Set(button, "Value", true);
                    }
                    finally
                    {
                        Release(button);
                    }
                }
                finally
                {
                    Release(oleObject);
                }
            }
            finally
            {
                Release(worksheet);
            }
        }
        finally
        {
            Release(sheets);
        }
    }

    /// <summary>Runs a macro by name through Application.Run, retrying while Excel has not registered it yet.</summary>
    public object? RunMacro(string name, int attempts = 40)
    {
        for (var attempt = 1; ; attempt++)
        {
            try
            {
                return Call(application, "Run", name);
            }
            catch (COMException) when (attempt < attempts)
            {
                Thread.Sleep(250);
            }
        }
    }

    public static void CloseWorkbook(object workbook)
    {
        try
        {
            Call(workbook, "Close", false);
        }
        finally
        {
            Release(workbook);
        }
    }

    [LibraryImport("user32.dll")]
    private static partial uint GetWindowThreadProcessId(nint hwnd, out uint processId);

    private const uint BmClick = 0x00F5;

    [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial nint FindWindowEx(nint parent, nint childAfter, string? className, string? windowName);

    [LibraryImport("user32.dll", EntryPoint = "GetWindowTextW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int GetWindowText(nint hwnd, [Out] char[] text, int maxCount);

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    private static partial nint SendMessage(nint hwnd, uint message, nint wParam, nint lParam);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool IsWindowVisible(nint hwnd);

    /// <summary>
    /// Answers a dialog this Excel shows, as a person would: waits for a visible window titled <paramref name="title"/> in
    /// Excel's process, returns the texts of its controls, and clicks the button captioned <paramref name="button"/> (an
    /// accelerator's &amp; ignored). Win32 only, so it runs on another thread while the STA thread is blocked in the call
    /// that raised the dialog; when none appears within the timeout, the result is null and Excel is killed so that call
    /// returns, unless <paramref name="killIfAbsent"/> is false because no dialog is the outcome being tested.
    /// </summary>
    public string[]? AnswerDialog(string title, string button, TimeSpan timeout, bool killIfAbsent = true)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            for (var window = FindWindowEx(0, 0, null, title); window != 0; window = FindWindowEx(0, window, null, title))
            {
                _ = GetWindowThreadProcessId(window, out var owner);
                if (owner != (uint)processId || !IsWindowVisible(window))
                {
                    continue;
                }

                var texts = new List<string>();
                nint target = 0;
                for (var child = FindWindowEx(window, 0, null, null); child != 0; child = FindWindowEx(window, child, null, null))
                {
                    var text = WindowText(child);
                    texts.Add(text);
                    if (text.Replace("&", string.Empty, StringComparison.Ordinal) == button)
                    {
                        target = child;
                    }
                }

                if (target != 0)
                {
                    _ = SendMessage(target, BmClick, 0, 0);
                    return [.. texts];
                }
            }

            Thread.Sleep(200);
        }

        if (killIfAbsent)
        {
            Kill();
        }

        return null;
    }

    private static string WindowText(nint hwnd)
    {
        var buffer = new char[1024];
        var length = GetWindowText(hwnd, buffer, buffer.Length);
        return new string(buffer, 0, Math.Max(length, 0));
    }


    /// <summary>Starts a hidden Excel, adds a workbook, and registers the add-in from its build output.</summary>
    public static ExcelInstance Start(string addInPath)
    {
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
        var application = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Excel.Application could not be created.");
        try
        {
            Set(application, "Visible", false);
            Set(application, "DisplayAlerts", false);
            var hwnd = Convert.ToInt64(Get(application, "Hwnd"), CultureInfo.InvariantCulture);
            _ = GetWindowThreadProcessId((nint)hwnd, out var processId);

            var workbooks = Get(application, "Workbooks")!;
            try
            {
                Release(Call(workbooks, "Add"));
            }
            finally
            {
                Release(workbooks);
            }

            if (Call(application, "RegisterXLL", addInPath) is not true)
            {
                throw new InvalidOperationException("Excel refused to register " + addInPath);
            }

            return new ExcelInstance(application, (int)processId);
        }
        catch
        {
            try
            {
                Call(application, "Quit");
            }
            catch (COMException)
            {
            }

            Release(application);
            throw;
        }
    }

    /// <summary>Calls a host command through Application.Run and returns its JSON; Excel is killed when the command exceeds the timeout.</summary>
    public string RunHostCommand(string command, TimeSpan timeout, params object[] arguments)
    {
        using var watchdog = new Timer(_ => Kill(), null, timeout, Timeout.InfiniteTimeSpan);
        var invokeArguments = new object[arguments.Length + 1];
        invokeArguments[0] = command;
        arguments.CopyTo(invokeArguments, 1);
        try
        {
            return Invoke(application, "Run", BindingFlags.InvokeMethod, invokeArguments) as string
                ?? throw new InvalidOperationException(command + " returned nothing.");
        }
        catch (COMException) when (killed)
        {
            throw new TimeoutException(command + " did not return within " + timeout + "; Excel was killed.");
        }
        finally
        {
            watchdog.Change(Timeout.Infinite, Timeout.Infinite);
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        if (!killed)
        {
            try
            {
                Call(application, "Quit");
            }
            catch (COMException)
            {
            }
        }

        Release(application);
        GC.Collect();
        GC.WaitForPendingFinalizers();

        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.WaitForExit(TimeSpan.FromSeconds(15)))
            {
                process.Kill();
                process.WaitForExit();
            }
        }
        catch (ArgumentException)
        {
            // Already gone.
        }
        catch (InvalidOperationException)
        {
            // Exited between the check and the kill.
        }
    }

    private void Kill()
    {
        killed = true;
        try
        {
            using var process = Process.GetProcessById(processId);
            process.Kill();
        }
        catch (ArgumentException)
        {
        }
        catch (InvalidOperationException)
        {
        }
    }

    private static object? Get(object target, string name) => Invoke(target, name, BindingFlags.GetProperty, null);

    private static object? Get(object target, string name, params object?[] arguments) => Invoke(target, name, BindingFlags.GetProperty, arguments);

    private static void Set(object target, string name, object? value) => Invoke(target, name, BindingFlags.SetProperty, [value]);

    private static object? Call(object target, string name, params object?[] arguments) => Invoke(target, name, BindingFlags.InvokeMethod, arguments);

    public static void Release(object? target)
    {
        if (target is not null && Marshal.IsComObject(target))
        {
            Marshal.FinalReleaseComObject(target);
        }
    }

    private static object? Invoke(object target, string name, BindingFlags flags, object?[]? arguments)
    {
        try
        {
            return target.GetType().InvokeMember(name, flags, binder: null, target, arguments, CultureInfo.InvariantCulture);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }
}

/// <summary>Runs COM work on a fresh STA thread (CLAUDE.md R17) and rethrows what it threw.</summary>
internal static class Sta
{
    public static void Run(Action action)
    {
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                failure = ExceptionDispatchInfo.Capture(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        failure?.Throw();
    }
}
