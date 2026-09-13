using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

using VbaNg.Golden.Harness;

namespace VbaNg.Golden.Regen;

internal enum RunOutcome
{
    Completed,
    TimedOut,

    /// <summary>Excel refused to run the procedure: the project does not compile (a name that does not resolve, for example), which a hidden Excel reports instead of showing the dialog.</summary>
    Rejected,
}

/// <summary>
/// A throwaway hidden Excel (CLAUDE.md R17): started fresh, never the user's instance, quit on
/// dispose and killed if it will not quit. Recorder modules are injected through the VBE object
/// model, which requires the "Trust access to the VBA project object model" setting.
/// </summary>
internal sealed partial class ExcelSession : IDisposable
{
    private const int StandardModule = 1; // vbext_ct_StdModule
    private const int CannotRunMacro = unchecked((int)0x800A03EC); // "Cannot run the macro": the project failed to compile

    private readonly object application;
    private readonly object workbook;
    private readonly string workbookName;
    private readonly int processId;
    private object? project;
    private volatile bool killed;
    private bool disposed;

    private ExcelSession(object application, object workbook, string workbookName, int processId)
    {
        this.application = application;
        this.workbook = workbook;
        this.workbookName = workbookName;
        this.processId = processId;
    }

    [LibraryImport("user32.dll")]
    private static partial uint GetWindowThreadProcessId(nint hwnd, out uint processId);

    /// <summary>Excel's version and build, for example 16.0.20326.</summary>
    public string Version => string.Create(CultureInfo.InvariantCulture, $"{Com.Get(application, "Version")}.{Com.Get(application, "Build")}");

    public static ExcelSession Start()
    {
        var type = Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
        var application = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Excel.Application could not be created.");
        try
        {
            Com.Set(application, "Visible", false);
            Com.Set(application, "DisplayAlerts", false);
            var hwnd = Convert.ToInt64(Com.Get(application, "Hwnd"), CultureInfo.InvariantCulture);
            _ = GetWindowThreadProcessId((nint)hwnd, out var processId);

            var workbooks = Com.Get(application, "Workbooks")!;
            object workbook;
            try
            {
                workbook = Com.Call(workbooks, "Add")!;
            }
            finally
            {
                Com.Release(workbooks);
            }

            var name = (string)Com.Get(workbook, "Name")!;
            return new ExcelSession(application, workbook, name, (int)processId);
        }
        catch
        {
            try
            {
                Com.Call(application, "Quit");
            }
            catch (COMException)
            {
            }

            Com.Release(application);
            throw;
        }
    }

    /// <summary>Opens the workbook's VBA project; false with Excel's reason when access is not trusted.</summary>
    public bool TryOpenProject(out string reason)
    {
        try
        {
            project = Com.Get(workbook, "VBProject");
            reason = string.Empty;
            return project is not null;
        }
        catch (COMException ex)
        {
            reason = ex.Message.Trim();
            return false;
        }
    }

    /// <summary>
    /// Adds the module and imports the area's class files, runs the entry procedure, and removes
    /// them all. Excel is killed when the run exceeds the timeout, which is what a compile error
    /// or a modal dialog looks like from here; the session is unusable afterwards.
    /// </summary>
    public RunOutcome RunModule(string moduleName, string moduleText, IReadOnlyList<string> classFiles, TimeSpan timeout)
    {
        if (project is null)
        {
            throw new InvalidOperationException("Call TryOpenProject first.");
        }

        var components = Com.Get(project, "VBComponents")!;
        object? component = null;
        var imported = new List<object>();
        // The watchdog covers the imports and AddFromString too: the VBE shows a modal dialog for a syntax error it finds at insertion, before anything runs.
        using var watchdog = new Timer(_ => Kill(), null, timeout, Timeout.InfiniteTimeSpan);
        try
        {
            foreach (var classFile in classFiles)
            {
                imported.Add(Com.Call(components, "Import", classFile)!);
            }

            component = Com.Call(components, "Add", StandardModule)!;
            Com.Set(component, "Name", moduleName);
            var codeModule = Com.Get(component, "CodeModule")!;
            try
            {
                Com.Call(codeModule, "AddFromString", moduleText);
            }
            finally
            {
                Com.Release(codeModule);
            }

            try
            {
                Com.Call(application, "Run", $"'{workbookName}'!{moduleName}.{RecorderModule.RunProcedure}");
            }
            catch (COMException) when (killed)
            {
                return RunOutcome.TimedOut;
            }
            catch (COMException ex) when (ex.ErrorCode == CannotRunMacro)
            {
                watchdog.Change(Timeout.Infinite, Timeout.Infinite);
                return RunOutcome.Rejected;
            }

            watchdog.Change(Timeout.Infinite, Timeout.Infinite);
            return killed ? RunOutcome.TimedOut : RunOutcome.Completed;
        }
        finally
        {
            foreach (var added in imported.Prepend(component).OfType<object>().Reverse())
            {
                if (!killed)
                {
                    try
                    {
                        Com.Call(components, "Remove", added);
                    }
                    catch (COMException)
                    {
                    }
                }

                Com.Release(added);
            }

            Com.Release(components);
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
                Com.Call(workbook, "Close", false);
                Com.Call(application, "Quit");
            }
            catch (COMException)
            {
            }
        }

        Com.Release(project);
        Com.Release(workbook);
        Com.Release(application);
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
}
