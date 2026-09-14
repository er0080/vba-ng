using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

using ExcelDna.Integration;

using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

/// <summary>
/// Hidden XLL functions the CLI calls through <c>Application.Run</c> (ARCHITECTURE.md D6 and
/// section 8, "Host commands"). Each returns a <see cref="RunResponse"/> as JSON.
/// </summary>
public static class HostCommands
{
    internal static readonly ProjectHost Projects = new();
    private static string? lastStartupRun;

    /// <summary>The workbook lifecycle, started by <see cref="AddIn.AutoOpen"/>.</summary>
    internal static WorkbookProjects? Workbooks { get; set; }

    [ExcelFunction(
        Name = "vbang.Run",
        IsHidden = true,
        IsMacroType = true,
        Description = "vba-ng host command: loads out/<Project>.dll from a project folder and runs Module.Procedure; ui \"ui\" keeps dialogs interactive.")]
    public static string Run(string projectDir, string procedure, string ui)
    {
        var capture = new CapturingHostServices();
        using (ExcelHostServices.Instance.Capture(capture, interactive: ui == "ui"))
        {
            using var running = ExcelHostServices.Instance.Running(projectDir);
            try
            {
                var project = Projects.Load(projectDir);
                Workbooks?.Refresh(project);
                project.Run(procedure);
                return Transport(RunResponse.Success(capture.Text), projectDir);
            }
            catch (Exception ex)
            {
                // An error the procedure left unhandled reaches the project's log as a button's does, and the reply names it
                // as VBA's dialog would; the output printed before it is the reply's.
                var output = capture.Text;
                ExcelHostServices.Instance.ReportUnhandled(procedure, ex, dialog: ui == "ui");
                var error = UnhandledError.Of(ex);
                var message = error is null
                    ? $"vba-ng could not run {procedure}: {ex.GetType().Name}: {ex.Message}"
                    : string.Create(CultureInfo.InvariantCulture, $"Run-time error '{error.Number}': {error.Description}");
                return Transport(RunResponse.Failure(output, message, ex.StackTrace), projectDir);
            }
        }
    }

    /// <summary>A response too long for Application.Run goes to out/response.json; when even that fails, the inline form is returned and may be cut short by Excel.</summary>
    private static string Transport(RunResponse response, string projectDir)
    {
        try
        {
            return response.ToTransportJson(projectDir);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return response.ToJson();
        }
    }

    /// <summary>
    /// Runs the project's <c>'@Test</c> procedures (ARCHITECTURE.md section 8). The response is
    /// Ok when every test passed; otherwise its output is still the report and its error the
    /// summary line, with no detail. A detail means the command itself failed, for example
    /// because the project is not built.
    /// </summary>
    [ExcelFunction(
        Name = "vbang.Test",
        IsHidden = true,
        IsMacroType = true,
        Description = "vba-ng host command: runs the '@Test procedures of a built project and writes JUnit XML to junitPath when it is not empty; ui \"ui\" keeps dialogs interactive.")]
    public static string Test(string projectDir, string junitPath, string ui)
    {
        var capture = new CapturingHostServices();
        using (ExcelHostServices.Instance.Capture(capture, interactive: ui == "ui"))
        {
            using var running = ExcelHostServices.Instance.Running(projectDir);
            try
            {
                var result = TestRunner.Run(Projects.Load(projectDir));
                if (!string.IsNullOrWhiteSpace(junitPath))
                {
                    File.WriteAllText(junitPath, result.ToJUnitXml());
                }

                var response = result.AllPassed
                    ? RunResponse.Success(result.ToText())
                    : RunResponse.Failure(result.ToText(), result.Summary, null);
                return Transport(response, projectDir);
            }
            catch (Exception ex)
            {
                return Transport(RunResponse.Failure(capture.Text, ex.GetType().Name + ": " + ex.Message, ex.StackTrace ?? string.Empty), projectDir);
            }
        }
    }

    /// <summary>
    /// Which collation the process runs with (ARCHITECTURE.md section 5): under Windows NLS the
    /// line feed inside CRLF is found by a culture-aware search, under ICU it is not, so the probe
    /// tells the two apart without depending on how the switch was set.
    /// </summary>
    private static string Collation => "\r\n".IndexOf("\n", StringComparison.CurrentCulture) == 1 ? "NLS" : "ICU";

    [ExcelFunction(
        Name = "vbang.Status",
        IsHidden = true,
        Description = "vba-ng host command: reports the add-in version, loaded projects, and the startup run.")]
    public static string Status()
    {
        var version = typeof(HostCommands).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown";
        var projects = string.Join(", ", Projects.Projects.Select(p => p.Name + " (" + p.ProjectDir + ")"));
        var workbooks = string.Join(", ", Workbooks?.Summary ?? []);
        var recent = string.Join(" | ", ExcelHostServices.Instance.Recent);
        var text = string.Create(
            CultureInfo.InvariantCulture,
            $"vba-ng {version}; collation: {Collation}; loaded projects: {(projects.Length == 0 ? "none" : projects)}; bound workbooks: {(workbooks.Length == 0 ? "none" : workbooks)}; reloads: {Workbooks?.Reloads ?? 0}; startup run: {lastStartupRun ?? "none"}; recent output: {(recent.Length == 0 ? "none" : recent)}");
        return RunResponse.Success(text).ToJson();
    }

    /// <summary>
    /// Runs the procedure named by the startup environment variables (see <see cref="AddIn"/>).
    /// Output reaches the attached debugger through <see cref="ExcelHostServices"/>; the outcome is
    /// reported by <see cref="Status"/>.
    /// </summary>
    internal static void RunFromStartup(string projectDir, string procedure, string readinessSummary)
    {
        var response = RunResponse.FromJson(RunResponse.ResolveTransport(Run(projectDir, procedure, string.Empty)));
        var outcome = response.Ok
            ? "ok (" + procedure + ")"
            : "error (" + procedure + "): " + response.Error;
        lastStartupRun = outcome + "; " + readinessSummary;

        // Run has already printed the error itself.
        if (!response.Ok)
        {
            if (response.Detail is not null)
            {
                ExcelHostServices.Instance.Print(response.Detail);
            }
        }
    }

    /// <summary>Records that the startup procedure could not be run because Excel never became ready.</summary>
    internal static void RecordStartupFailure(string reason)
    {
        lastStartupRun = "not run: " + reason;
        ExcelHostServices.Instance.Print("Startup run skipped: " + reason);
    }

    internal static void Shutdown()
    {
        Workbooks?.Dispose();
        Workbooks = null;
        Projects.UnloadAll();
    }
}
