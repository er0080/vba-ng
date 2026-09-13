using ExcelDna.Integration;

using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

/// <summary>Add-in entry points. Excel-DNA calls <see cref="AutoOpen"/> when the .xll loads and <see cref="AutoClose"/> when it unloads.</summary>
public sealed class AddIn : IExcelAddIn
{
    /// <summary>Project folder to run at startup; set by the VS Code launch configuration.</summary>
    public const string StartupProjectVariable = "VBANG_PROJECT";

    /// <summary>Procedure (Module.Procedure) to run at startup; set by the VS Code launch configuration.</summary>
    public const string StartupProcedureVariable = "VBANG_RUN";

    public void AutoOpen()
    {
        Host.Current = ExcelHostServices.Instance;
        var provider = new ExcelComProvider();
        Com.Provider = provider;

        // Workbook binding (ARCHITECTURE.md section 7) starts once the object model answers; behind the
        // splash screen Application may not be ready, so the first attempt is queued as a macro.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                var workbooks = new WorkbookProjects(HostCommands.Projects, text => ExcelHostServices.Instance.Print(text));
                workbooks.Start(provider.Application);
                HostCommands.Workbooks = workbooks;
            }
            catch (Exception ex) when (ex is VbaException or InvalidOperationException or System.Runtime.InteropServices.COMException)
            {
                ExcelHostServices.Instance.Print("Workbook binding could not start: " + ex.Message);
            }
        });

        var projectDir = Environment.GetEnvironmentVariable(StartupProjectVariable);
        var procedure = Environment.GetEnvironmentVariable(StartupProcedureVariable);
        if (!string.IsNullOrWhiteSpace(projectDir) && !string.IsNullOrWhiteSpace(procedure))
        {
            // AutoOpen runs behind the splash screen, before Excel's window exists.
            StartupRunner.Start(projectDir, procedure);
        }
    }

    public void AutoClose()
    {
        HostCommands.Shutdown();
        (Com.Provider as ExcelComProvider)?.Release();
    }
}
