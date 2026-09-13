using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

using VbaNg.Compiler;
using VbaNg.Import;
using VbaNg.Interop;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.E2E;

/// <summary>
/// ROADMAP.md WP3, the M6 criterion 2: the corpus libraries run, not just compile. Each spec
/// workbook of the corpus (VBA-JSON, VBA-Dictionary, VBA-UTC) is imported with the MS-OVBA
/// importer, which brings the library, its specs, and the VBA-TDD SpecSuite the workbook
/// carries; a <c>'@Test</c> bridge runs the suite under <c>vbang.Test</c> in a hidden Excel and
/// fails on any spec that fails, naming it. A build stopped only by marshaling vba-ng has not
/// recorded (Declare parameters of user-defined type, the pointer family) is skipped; any other
/// build error fails the test. VBA-Better-Array ships no spec workbook: its sources and its own
/// TestRunner build from src into a workbook-bound project, and its report is compared with VBA's.
/// Needs VBANG_E2E and VBANG_CORPUS, like the scorecard.
/// </summary>
[Trait("Category", "E2E")]
public sealed class CorpusSpecTests : IDisposable
{
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(300);
    private static readonly TimeSpan OthersTimeout = TimeSpan.FromMinutes(2);
    private readonly string workDir = Path.Combine(Path.GetTempPath(), "vbang-e2e", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try
        {
            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Theory]
    [InlineData("VBA-JSON", "specs/VBA-JSON - Specs.xlsm")]
    [InlineData("VBA-Dictionary", "specs/VBA-Dictionary - Specs.xlsm")]
    [InlineData("VBA-UTC", "specs/VBA-UTC - Specs.xlsm")]
    public void SpecWorkbook_ImportsBuildsAndPasses(string library, string workbook)
    {
        var addIn = RequireAddIn();
        var source = RequireCorpusFile(library, workbook);

        var folder = Path.Combine(workDir, library);
        Directory.CreateDirectory(folder);
        var project = VbaProjectReader.FromWorkbook(source);
        Assert.NotNull(project);
        var projectDir = Path.Combine(folder, library + ProjectPaths.FolderSuffix);
        var imported = ProjectWriter.Write(project, projectDir);
        Assert.Contains(imported.Files, f => f.Equals("Specs.bas", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(imported.Files, f => f.Equals("SpecSuite.cls", StringComparison.OrdinalIgnoreCase));
        AddBridge(projectDir, library);

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        if (!build.Success)
        {
            var errors = build.Diagnostics.Where(d => d.IsError).ToList();
            var unexplained = errors.Where(d => !NeedsMarshaling(d)).ToList();
            Assert.True(unexplained.Count == 0, "Build errors beyond Declare marshaling:" + Environment.NewLine + string.Join(Environment.NewLine, unexplained));
            Assert.Skip($"{library} needs marshaling vba-ng has not recorded: {errors.Count} diagnostic(s) on Declare parameter types and the pointer family, for example: {errors[0].Message}");
        }

        var junitPath = Path.Combine(folder, "specs.xml");
        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Test", CommandTimeout, projectDir, junitPath, string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Output + Environment.NewLine + response.Error);
        Assert.Contains("1 test: 1 passed, 0 failed, 0 errors.", response.Output, StringComparison.Ordinal);
    }

    /// <summary>
    /// VBA-Better-Array's own test runner, whose suite does not pass whole in VBA either (a test compares text
    /// that depends on the machine's formats), so its report under vbang is compared with the one VBA gives on
    /// the same machine: the same total, the same failures. Under vbang the sources and tests build as one
    /// project bound to a workbook by folder convention (ARCHITECTURE.md D4), so ThisWorkbook is that workbook
    /// as in the author's dev workbook, and a <c>'@Test</c> bridge prints the report; under VBA the same sources
    /// are imported into a saved workbook. Both runners list the TestModule_ files and call each test through
    /// Application.Run (ARCHITECTURE.md D23). Some tests start Excel instances of their own and quit them; each
    /// run waits for those to go.
    /// </summary>
    [Fact]
    public void BetterArray_TestRunner_MatchesVba()
    {
        var addIn = RequireAddIn();
        var sourceDir = Path.GetDirectoryName(RequireCorpusFile("VBA-Better-Array", "src/TestRunner.bas"))!;
        var folder = Path.Combine(workDir, "VBA-Better-Array");
        var projectDir = Path.Combine(folder, "BetterArray" + ProjectPaths.FolderSuffix);
        var srcDir = Path.Combine(folder, "src");
        Directory.CreateDirectory(projectDir);
        Directory.CreateDirectory(srcDir);
        foreach (var file in Directory.EnumerateFiles(sourceDir).Where(f => f.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".cls", StringComparison.OrdinalIgnoreCase)))
        {
            File.Copy(file, Path.Combine(projectDir, Path.GetFileName(file)));
            File.Copy(file, Path.Combine(srcDir, Path.GetFileName(file)));
        }

        File.WriteAllText(Path.Combine(projectDir, "ThisWorkbook.cls"), string.Join("\r\n", [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"ThisWorkbook\"",
            "Attribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True",
            "Option Explicit",
            string.Empty,
        ]));
        File.WriteAllText(ProjectPaths.ManifestPath(projectDir), "{ \"references\": [ { \"name\": \"Excel\" }, { \"name\": \"vbang\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "VbangTests.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"VbangTests\"",
            "Option Explicit",
            string.Empty,
            "'@Test",
            "Public Sub RunnerReport()",
            "    Debug.Print TestRunner.RunAllTests_Report()",
            "End Sub",
            string.Empty,
        ]));
        var workbookPath = Path.Combine(folder, "BetterArray.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Hello", "Hello.xlsx"), workbookPath);

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics.Where(d => d.IsError)));

        var junitPath = Path.Combine(folder, "tests.xml");
        var json = string.Empty;
        var before = ExcelInstance.RunningInstances();
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            json = excel.RunHostCommand("vbang.Test", CommandTimeout, projectDir, junitPath, string.Empty);
            ExcelInstance.CloseWorkbook(workbook);
        });
        ExcelInstance.WaitForOtherInstances(before, OthersTimeout);

        var vba = string.Empty;
        before = ExcelInstance.RunningInstances();
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.ActiveWorkbook();
            ExcelInstance.SaveAs(workbook, Path.Combine(folder, "BetterArrayVba.xlsm"), 52);
            foreach (var file in Directory.EnumerateFiles(srcDir))
            {
                ExcelInstance.ImportModule(workbook, file);
            }

            vba = excel.RunMacro("'" + ExcelInstance.Name(workbook) + "'!TestRunner.RunAllTests_Report") as string ?? string.Empty;
            ExcelInstance.Release(workbook);
        });
        ExcelInstance.WaitForOtherInstances(before, OthersTimeout);

        var response = RunResponse.FromJson(RunResponse.ResolveTransport(json));
        Assert.True(response.Ok, response.Output + Environment.NewLine + response.Error);
        var report = string.Join("\n", XDocument.Load(junitPath).Descendants("testcase")
            .Where(test => (string?)test.Attribute("name") == "RunnerReport")
            .Select(test => (string?)test.Element("system-out") ?? string.Empty));
        File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "betterarray-report.txt"), "VBA:\n" + vba + "\n\nvbang:\n" + report);
        Assert.StartsWith("[VBATests] Total: ", RunnerTotal(vba), StringComparison.Ordinal);
        var vbaFailures = RunnerFailures(vba);
        var vbangFailures = RunnerFailures(report);
        Assert.True(
            vbaFailures.SequenceEqual(vbangFailures),
            "failing under vbang only: " + string.Join(" | ", vbangFailures.Except(vbaFailures)) + Environment.NewLine + "failing under VBA only: " + string.Join(" | ", vbaFailures.Except(vbangFailures)));
        Assert.Equal(RunnerTotal(vba), RunnerTotal(report));
    }

    /// <summary>The runner's report lines, trimmed. Under vbang they are the bridge test's system-out in the JUnit file, since the text report shows only a failing test's output.</summary>
    private static IEnumerable<string> ReportLines(string text) =>
        text.Split('\n').Select(line => line.Trim());

    private static string RunnerTotal(string report) =>
        ReportLines(report).FirstOrDefault(line => line.StartsWith("[VBATests] Total: ", StringComparison.Ordinal)) ?? "(no total)";

    /// <summary>The report's failure entries (<c>12. Module.Test :: detail</c>), without their numbers, in order.</summary>
    private static string[] RunnerFailures(string report) =>
        [.. ReportLines(report)
            .Select(line => (Dot: line.IndexOf(". ", StringComparison.Ordinal), Line: line))
            .Where(entry => entry.Dot > 0 && entry.Line[..entry.Dot].All(char.IsAsciiDigit) && entry.Line.Contains(" :: ", StringComparison.Ordinal))
            .Select(entry => entry.Line[(entry.Dot + 2)..])
            .Order(StringComparer.Ordinal)];

    /// <summary>
    /// The bridge between the workbook's SpecSuite and <c>vbang test</c>: runs the suite through
    /// the InlineRunner the workbook carries, so every failed expectation is printed, and fails
    /// the test when any spec failed, naming them.
    /// </summary>
    private static void AddBridge(string projectDir, string library)
    {
        var bridge = new StringBuilder();
        bridge.Append("Attribute VB_Name = \"VbangSpecs\"\r\n");
        bridge.Append("Option Explicit\r\n\r\n");
        bridge.Append("'@Test\r\n");
        bridge.Append("Public Sub AllSpecsPass()\r\n");
        bridge.Append("    Dim suite As SpecSuite\r\n");
        if (library == "VBA-UTC")
        {
            // VBA-UTC's specs ask for the machine's UTC offset through an InputBox, which a run with no one at the
            // screen answers with its default, 0; the bridge answers as the person running them would.
            var minutes = (int)TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).TotalMinutes;
            File.AppendAllText(Path.Combine(projectDir, "Specs.bas"), "\r\nPublic Sub VbangSetOffset(ByVal minutes As Long)\r\n    pOffsetMinutes = minutes\r\n    pOffsetLoaded = True\r\nEnd Sub\r\n");
            bridge.Append(CultureInfo.InvariantCulture, $"    Call Specs.VbangSetOffset({minutes})\r\n");
        }

        bridge.Append("    Set suite = Specs.Specs()\r\n");
        bridge.Append("    InlineRunner.RunSuite suite, ShowFailureDetails:=True, ShowPassed:=False, ShowSuiteDetails:=True\r\n");
        bridge.Append("    Dim spec As SpecDefinition\r\n");
        bridge.Append("    Dim failed As String\r\n");
        bridge.Append("    For Each spec In suite.SpecsCol\r\n");
        bridge.Append("        If spec.Result = SpecResult.Fail Then failed = failed & \"; \" & spec.Description\r\n");
        bridge.Append("    Next\r\n");
        bridge.Append("    Assert.AreEqual \"\", failed, \"failed specs\"\r\n");
        bridge.Append("End Sub\r\n");
        File.WriteAllText(Path.Combine(projectDir, "VbangSpecs.bas"), bridge.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        // The Assert module comes with the vbang reference (ARCHITECTURE.md section 8).
        var manifestPath = ProjectPaths.ManifestPath(projectDir);
        var manifest = File.ReadAllText(manifestPath);
        manifest = manifest.Replace("\"references\": [", "\"references\": [\n    { \"name\": \"vbang\" },", StringComparison.Ordinal);
        File.WriteAllText(manifestPath, manifest, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>A diagnostic about marshaling vba-ng has not recorded: Declare parameters beyond the machine types, and the pointer family called in a shape that does not bind as an address.</summary>
    private static bool NeedsMarshaling(Diagnostic diagnostic) =>
        diagnostic.Id == DiagnosticIds.NotSupported
        && (diagnostic.Message.Contains("passed ByRef to a Declare", StringComparison.Ordinal)
            || diagnostic.Message.Contains("passed ByVal to a Declare", StringComparison.Ordinal)
            || diagnostic.Message.Contains("binds as an address", StringComparison.Ordinal)
            || diagnostic.Message.Contains("needs its own marshaling", StringComparison.Ordinal));

    private static string RequireAddIn()
    {
        Assert.SkipWhen(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VBANG_E2E")), "Set VBANG_E2E=1 to run the Excel-driving tests.");
        var configuration = typeof(CorpusSpecTests).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? "Debug";
        var path = Path.Combine(RepositoryRoot(), "src", "VbaNg.AddIn", "bin", configuration, "net10.0-windows", "VbaNg.AddIn-AddIn64.xll");
        Assert.True(File.Exists(path), "Add-in not built: " + path);
        return path;
    }

    /// <summary>
    /// ROADMAP.md WP6, the M8 criterion 4: a real workbook D-H names goes through the importer and
    /// builds from what the workbook carries, with no source edits. The spec suites above drive
    /// their workbooks in Excel; this is the gate that these two still compile, which five
    /// divergences had to be fixed for (WP6). Excel is not needed, only the corpus.
    /// </summary>
    [Theory]
    [InlineData("stdVBA", "testBuilder.xlsm")]
    [InlineData("VBA-Web", "specs/VBA-Web - Specs.xlsm")]
    public void Workbook_ImportsAndBuildsFromTheWorkbook(string library, string workbook)
    {
        var source = RequireCorpusFile(library, workbook);
        var projectDir = Path.Combine(workDir, library, "Book" + ProjectPaths.FolderSuffix);
        var project = VbaProjectReader.FromWorkbook(source);
        Assert.NotNull(project);
        ProjectWriter.Write(project, projectDir, documentKinds: WorkbookCodeNames.Read(source));

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));

        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics.Where(d => d.IsError)));
    }

    private static string RequireCorpusFile(string library, string relativePath)
    {
        var setting = Environment.GetEnvironmentVariable("VBANG_CORPUS");
        Assert.SkipWhen(string.IsNullOrWhiteSpace(setting), "Set VBANG_CORPUS to the corpus folder (or several, separated by ';') to run the spec suites.");
        foreach (var root in setting!.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var path = Path.Combine(root, library, relativePath);
            if (File.Exists(path))
            {
                return path;
            }
        }

        Assert.Skip($"{library}/{relativePath} is not under VBANG_CORPUS.");
        return string.Empty;
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "VbaNg.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Repository root (VbaNg.slnx) not found above " + AppContext.BaseDirectory);
    }
}
