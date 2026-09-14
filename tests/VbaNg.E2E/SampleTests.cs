using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;

using VbaNg.Compiler;
using VbaNg.Interop;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.E2E;

/// <summary>
/// The M3 exit criterion (ROADMAP.md): the sample procedural project passes under <c>vbang test</c>
/// in a real Excel. Each test starts a hidden throwaway Excel with the add-in from the build
/// output, runs host commands in it, and quits it (CLAUDE.md R17). Skipped unless VBANG_E2E is
/// set, so plain <c>dotnet test</c> never needs Excel (CLAUDE.md R9).
/// </summary>
[Trait("Category", "E2E")]
public sealed class SampleTests : IDisposable
{
    private const string GateVariable = "VBANG_E2E";
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(120);
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

    [Fact]
    public void HostTest_ProceduralSample_PassesInExcel()
    {
        var addIn = RequireAddIn();
        var projectDir = BuildSample("Procedural");
        var junitPath = Path.Combine(workDir, "procedural.xml");

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Test", CommandTimeout, projectDir, junitPath, string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Output + Environment.NewLine + response.Error);
        Assert.Contains("passed  Tests.MovePoint_AddsTheOffset", response.Output, StringComparison.Ordinal);
        Assert.StartsWith("12 tests: 12 passed, 0 failed, 0 errors.", response.Output.Split('\n')[^1], StringComparison.Ordinal);
        var report = XDocument.Load(junitPath);
        Assert.Equal("12", report.Root!.Attribute("tests")!.Value);
        Assert.Equal("0", report.Root.Attribute("failures")!.Value);
        Assert.Equal("0", report.Root.Attribute("errors")!.Value);
    }

    /// <summary>The M5 exit criterion (ROADMAP.md): the class-heavy sample passes in a real Excel.</summary>
    [Fact]
    public void HostTest_ClassesSample_PassesInExcel()
    {
        var addIn = RequireAddIn();
        var projectDir = BuildSample("Classes");

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Test", CommandTimeout, projectDir, string.Empty, string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Output + Environment.NewLine + response.Error);
        Assert.Contains("passed  Tests.Events_RunTheHandlerAndCanCancel", response.Output, StringComparison.Ordinal);
        Assert.Contains("passed  Tests.Terminate_RunsWhenTheLastReferenceGoes", response.Output, StringComparison.Ordinal);
        Assert.StartsWith("10 tests: 10 passed, 0 failed, 0 errors.", response.Output.Split('\n')[^1], StringComparison.Ordinal);
    }

    [Fact]
    public void HostRun_ProceduralSample_PrintsInExcel()
    {
        var addIn = RequireAddIn();
        var projectDir = BuildSample("Procedural");

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Run", CommandTimeout, projectDir, "Main.Main", string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Error);
        Assert.Equal(["Distance from origin: 5", "first: triangle", "Next id: 1"], response.Output.Split(Environment.NewLine));
    }

    /// <summary>
    /// Excel's object model through early binding (ARCHITECTURE.md section 6): the project
    /// references the Excel type library, its unqualified globals reach the running Application,
    /// and cells are read and written through Range, Cells, ActiveSheet, and default members.
    /// </summary>
    [Fact]
    public void HostRun_ExcelGlobals_ReadAndWriteCells()
    {
        var addIn = RequireAddIn();
        var projectDir = Path.Combine(workDir, "Cells" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Option Explicit",
            "",
            "Public Sub Fill()",
            "    Dim ws As Worksheet",
            "    Set ws = ActiveSheet",
            "    ws.Range(\"A1\").Value = 42",
            "    Range(\"A2\").Value = \"hello\"",
            "    Cells(3, 1).Value = Range(\"A1\").Value * 2",
            "    Range(\"B1\") = \"direct\"",
            "    ActiveSheet.Range(\"B2\") = \"late\"",
            "    Dim x As Long",
            "    x = Range(\"A1\")",
            "    Debug.Print ws.Name; Range(\"A1\").Value; Cells(3, 1).Value; x",
            "    Debug.Print TypeName(ActiveSheet) & \" \" & TypeName(Range(\"A1\")) & \" \" & Application.Name",
            "    Debug.Print Range(\"A2\") & \"/\" & Range(\"B1\").Value & \"/\" & Cells(2, 2).Text",
            "    Debug.Print Range(\"A1\").Row + Range(\"A1\").Column; Worksheets.Count; xlUp",
            "    ' Foreign names (MS-VBAL 3.3.5.3): [A1] is Application.Evaluate, [A1] = v assigns its default member, ws.[A1] is the sheet's Evaluate.",
            "    [B3].Value = \"bracket\"",
            "    [B4] = [B3].Value & \"!\"",
            "    Debug.Print [B3].Value & \"/\" & ws.[B4].Value & \"/\" & [B3:B4].Count & \"/\" & ActiveSheet.[B4].Value",
            "End Sub",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Run", CommandTimeout, projectDir, "Main.Fill", string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Error + Environment.NewLine + response.Detail);
        Assert.Equal(
            ["Sheet1 42  84  42 ", "Worksheet Range Microsoft Excel", "hello/direct/late", " 2  1 -4162 ", "bracket/bracket!/2/bracket!"],
            response.Output.Split(Environment.NewLine));
    }

    /// <summary>
    /// Application.Run of the project's own procedures (ARCHITECTURE.md D23): the same modules run
    /// under VBA, imported into a throwaway workbook, and under vbang, and report the same outcome
    /// for each way of naming a procedure, passing it arguments, and failing (docs/vba-quirks.md,
    /// "Host, Application.Run"). The error a procedure leaves unhandled is not among them: VBA
    /// shows a modal dialog for it.
    /// </summary>
    [Fact]
    public void HostRun_ApplicationRun_MatchesVba()
    {
        var addIn = RequireAddIn();
        var projectDir = Path.Combine(workDir, "RunByName" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "Targets.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Targets\"",
            "Option Explicit",
            "Public Counter As Long",
            "Public Sub Bump()",
            "    Counter = Counter + 1",
            "End Sub",
            "Private Sub Hidden()",
            "    Counter = Counter + 100",
            "End Sub",
            "Public Sub Dup()",
            "    Counter = Counter + 10",
            "End Sub",
            "Public Function Twice(ByVal x As Long) As Long",
            "    Twice = x * 2",
            "End Function",
            "Public Sub Append(ByRef s As String)",
            "    s = s & \"!\"",
            "End Sub",
            "Public Sub SetVariant(v As Variant)",
            "    v = \"changed\"",
            "End Sub",
            "Public Function Kind(ByVal v As Variant) As String",
            "    Kind = TypeName(v)",
            "End Function",
            "Public Function Opt(Optional ByVal x As Variant) As String",
            "    Opt = \"missing=\" & IsMissing(x)",
            "End Function",
            "Public Function OptDefault(Optional ByVal x As Long = 7) As Long",
            "    OptDefault = x",
            "End Function",
            "Public Function Coerced(ByVal x As Long) As String",
            "    Coerced = TypeName(x) & x",
            "End Function",
            "Public Function Total(ParamArray p() As Variant) As Long",
            "    Total = UBound(p) - LBound(p) + 1",
            "End Function",
            "Public Function GetMissing(Optional v As Variant) As Variant",
            "    GetMissing = v",
            "End Function",
            "Public Function MakeCollection() As Collection",
            "    Set MakeCollection = New Collection",
            "    MakeCollection.Add 1",
            "End Function",
            string.Empty,
        ]));
        File.WriteAllText(Path.Combine(projectDir, "Other.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Other\"",
            "Option Explicit",
            "Public Sub Dup()",
            "    Targets.Counter = Targets.Counter + 20",
            "End Sub",
            string.Empty,
        ]));
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Option Explicit",
            "Private Out As String",
            "Private Sub Note(ByVal text As String)",
            "    Out = Out & text & vbLf",
            "End Sub",
            "Private Sub NoteErr(ByVal label As String)",
            "    Note label & \": \" & Err.Number & \" \" & Err.Source & \" \" & Err.Description",
            "    Err.Clear",
            "End Sub",
            "Public Function Report() As String",
            "    Dim r As Variant, s As String, v As Variant, i As Integer, vm As Variant, c As Object",
            "    On Error Resume Next",
            "    Out = \"\"",
            "    Targets.Counter = 0",
            "    Application.Run \"Bump\"",
            "    Application.Run \"Targets.Bump\"",
            "    Application.Run \"targets.BUMP\"",
            "    Application.Run \"Hidden\"",
            "    Application.Run \"Targets.Hidden\"",
            "    Note \"counter \" & Targets.Counter",
            "    Application.Run \"Dup\"",
            "    NoteErr \"ambiguous\"",
            "    Application.Run \"Other.Dup\"",
            "    Note \"counter \" & Targets.Counter",
            "    r = Application.Run(\"Twice\", 21)",
            "    Note \"twice \" & r & \" \" & TypeName(r)",
            "    r = Application.Run(\"Bump\")",
            "    Note \"sub \" & TypeName(r)",
            "    s = \"a\"",
            "    Application.Run \"Append\", s",
            "    v = \"orig\"",
            "    Application.Run \"SetVariant\", v",
            "    Note \"byref \" & s & \" \" & v",
            "    i = 5",
            "    Note \"kinds \" & Application.Run(\"Kind\", i) & \" \" & Application.Run(\"Kind\", CCur(5)) & \" \" & Application.Run(\"Kind\", CDec(5)) & \" \" & Application.Run(\"Kind\", Null) & \" \" & Application.Run(\"Kind\", Nothing) & \" \" & Application.Run(\"Kind\", Array(1, 2))",
            "    Note \"optional \" & Application.Run(\"Opt\") & \" \" & Application.Run(\"OptDefault\")",
            "    r = Application.Run(\"Twice\")",
            "    NoteErr \"too few\"",
            "    r = Application.Run(\"Twice\", 1, 2)",
            "    NoteErr \"too many\"",
            "    r = Application.Run(\"Twice\", \"x\")",
            "    NoteErr \"mismatch\"",
            "    Note \"coerced \" & Application.Run(\"Coerced\", 2.5) & \" \" & Application.Run(\"Twice\", \"12\")",
            "    Application.Run \"NoSuchProc\"",
            "    NoteErr \"missing\"",
            "    Application.Run \"Targets.NoSuch\"",
            "    NoteErr \"missing qualified\"",
            "    Note \"paramarray \" & Application.Run(\"Total\", 1, 2, 3) & \" \" & Application.Run(\"Total\")",
            "    vm = Targets.GetMissing()",
            "    Note \"trailing missing \" & Application.Run(\"Twice\", 3, vm)",
            "    r = Application.Run(\"Twice\", vm, 3)",
            "    NoteErr \"leading missing\"",
            "    r = Application.Run(\"Kind\", vm)",
            "    NoteErr \"only missing\"",
            "    Set c = Application.Run(\"MakeCollection\")",
            "    Note \"object \" & c.Count",
            "    Note \"global \" & Run(\"Twice\", 4)",
            "    Report = Out",
            "End Function",
            "Public Sub Go()",
            "    Debug.Print Report();",
            "End Sub",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        var vba = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.ActiveWorkbook();
            ExcelInstance.ImportModule(workbook, Path.Combine(projectDir, "Targets.bas"));
            ExcelInstance.ImportModule(workbook, Path.Combine(projectDir, "Other.bas"));
            ExcelInstance.ImportModule(workbook, Path.Combine(projectDir, "Main.bas"));

            vba = excel.RunMacro("'" + ExcelInstance.Name(workbook) + "'!Main.Report") as string ?? string.Empty;
            ExcelInstance.Release(workbook);
        });

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Run", CommandTimeout, projectDir, "Main.Go", string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Error + Environment.NewLine + response.Detail);
        Assert.StartsWith("counter 203\n", vba, StringComparison.Ordinal);
        Assert.Equal(ReportLines(vba), ReportLines(response.Output));
    }

    private static string[] ReportLines(string text) => text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n', StringSplitOptions.RemoveEmptyEntries);

    /// <summary>
    /// The M4 exit criterion (ROADMAP.md): a workbook with a button and sheet events runs its
    /// project unchanged. Opening samples/Sheets/Sheets.xlsx binds Sheets.vbang by folder
    /// convention (D4): Workbook_Open and Auto_Open run, Worksheet_Change fires when a cell
    /// changes, the Form control's OnAction resolves to the registered command, and closing the
    /// workbook runs Workbook_BeforeClose and unbinds everything.
    /// </summary>
    [Fact]
    public void Workbook_Lifecycle_BindsDocumentModulesAndEvents()
    {
        var addIn = RequireAddIn();
        var sample = Path.Combine(RepositoryRoot(), "samples", "Sheets");
        var sampleDir = Path.Combine(workDir, "Sheets");
        Directory.CreateDirectory(sampleDir);
        var workbookPath = Path.Combine(sampleDir, "Sheets.xlsx");
        File.Copy(Path.Combine(sample, "Sheets.xlsx"), workbookPath);
        var projectDir = Path.Combine(sampleDir, "Sheets" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        foreach (var file in Directory.GetFiles(Path.Combine(sample, "Sheets" + ProjectPaths.FolderSuffix)).Where(f => f.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".cls", StringComparison.OrdinalIgnoreCase) || f.EndsWith("vbang.json", StringComparison.OrdinalIgnoreCase)))
        {
            File.Copy(file, Path.Combine(projectDir, Path.GetFileName(file)));
        }

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        string? a1 = null, e1 = null, b1 = null, d1 = null, c1 = null, a3 = null, f1 = null, f2 = null, f3 = null, statusBefore = null, statusAfter = null, reloadStatus = null;
        object? a5 = null, a6 = null, a7 = null, a8 = null, a9 = null, a10 = null, a11 = null, a12 = null, a10Again = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            a1 = ExcelInstance.CellValue(workbook, "Sheet1", "A1") as string;
            e1 = ExcelInstance.CellValue(workbook, "Sheet1", "E1") as string;
            statusBefore = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;

            ExcelInstance.SetCellValue(workbook, "Sheet1", "A2", 5);
            b1 = ExcelInstance.CellValue(workbook, "Sheet1", "B1") as string;
            // WithEvents on library types: the Watcher class advised Sheet1, ThisWorkbook advised the workbook itself (a two-argument event).
            f1 = ExcelInstance.CellValue(workbook, "Sheet1", "F1") as string;
            f2 = ExcelInstance.CellValue(workbook, "Sheet1", "F2") as string;
            // WP4: WithEvents on Application in a document module, whose events come from a different
            // source interface (AppEvents) than the workbook's and reach every open workbook.
            f3 = ExcelInstance.CellValue(workbook, "Sheet1", "F3") as string;

            try
            {
                excel.RunMacro("Module1.ButtonClick", attempts: 12);
            }
            catch (COMException ex)
            {
                d1 = "not run: " + ex.Message + " / " + RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            }

            d1 ??= ExcelInstance.CellValue(workbook, "Sheet1", "D1") as string;

            // The ActiveX button: setting its Value fires Click, which Sheet1's handler answers in C1.
            try
            {
                ExcelInstance.ClickActiveXButton(workbook, "Sheet1", "CommandButton1");
                c1 = ExcelInstance.CellValue(workbook, "Sheet1", "C1") as string;
            }
            catch (COMException ex)
            {
                c1 = "not clicked: " + ex.Message;
            }

            // UDFs: a typed parameter gets the value, a Variant parameter the Range, a run-time error shows #VALUE!, Module.Name works too.
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A5", "=Twice(21)");
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A6", "=Describe(A2)");
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A7", "=Failing()");
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A8", "=Module1.Twice(2)");
            a5 = ExcelInstance.CellValue(workbook, "Sheet1", "A5");
            a6 = ExcelInstance.CellValue(workbook, "Sheet1", "A6");
            a7 = ExcelInstance.CellValue(workbook, "Sheet1", "A7");
            a8 = ExcelInstance.CellValue(workbook, "Sheet1", "A8");

            // WP4: Application.Volatile, Caller, and ThisCell inside a UDF, which COM cannot answer
            // while Excel recalculates; the host answers them from the call Excel is making.
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A10", "=VolatileCount()");
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A11", "=CallerAddress()");
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A12", "=ThisCellAddress()");
            a10 = ExcelInstance.CellValue(workbook, "Sheet1", "A10");
            a11 = ExcelInstance.CellValue(workbook, "Sheet1", "A11");
            a12 = ExcelInstance.CellValue(workbook, "Sheet1", "A12");

            // A volatile function evaluates again when a cell it does not read changes; a plain one does not.
            ExcelInstance.SetCellValue(workbook, "Sheet1", "A4", 1);
            a10Again = ExcelInstance.CellValue(workbook, "Sheet1", "A10");

            // Hot reload: a function added to the source is rebuilt by the watcher and formulas find it.
            File.AppendAllText(Path.Combine(projectDir, "Module1.bas"), "\r\nPublic Function Thrice(x As Double) As Double\r\n    Thrice = x * 3\r\nEnd Function\r\n");
            reloadStatus = WaitForStatus(excel, status => status.Contains("reloads: 1", StringComparison.Ordinal), TimeSpan.FromSeconds(90));
            for (var attempt = 0; attempt < 20; attempt++)
            {
                ExcelInstance.SetCellValue(workbook, "Sheet1", "A9", "=Thrice(2)");
                a9 = ExcelInstance.CellValue(workbook, "Sheet1", "A9");
                if (a9 is double)
                {
                    break;
                }

                Thread.Sleep(500);
            }

            a3 = ExcelInstance.CellValue(workbook, "Sheet1", "A3") as string;
            ExcelInstance.CloseWorkbook(workbook);
            statusAfter = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
        });

        Assert.True(a1 == "opened", $"A1 was '{a1}'; status: {statusBefore}");
        Assert.True(e1 == "auto", $"E1 was '{e1}'; status: {statusBefore}");
        Assert.True(statusBefore!.Contains("bound workbooks: Sheets.xlsx -> " + projectDir, StringComparison.OrdinalIgnoreCase), "status before close: " + statusBefore);
        Assert.True(b1 == "changed $A$2 to 5", $"B1 was '{b1}'; status: {statusBefore}");
        Assert.True(f1 == "watched $A$2", $"F1 was '{f1}'; status: {statusBefore}");
        Assert.True(f2 == "Sheet1!A2", $"F2 was '{f2}'; status: {statusBefore}");
        Assert.True(f3 == "app Sheet1", $"F3 was '{f3}'; status: {statusBefore}");
        Assert.True(d1 == "button", $"D1 was '{d1}'; status: {statusBefore}");
        Assert.True(c1 == "clicked", $"C1 was '{c1}'; status: {statusAfter}");
        Assert.True(a5 is double v5 && v5 == 42, $"A5 was '{a5}'; status: {statusAfter}");
        Assert.True(a6 as string == "Range A2 5", $"A6 was '{a6}'; status: {statusAfter}");
        Assert.True(a7 is int e7 && e7 == -2146826273, $"A7 was '{a7}' rather than #VALUE!; status: {statusAfter}");
        Assert.True(a8 is double v8 && v8 == 4, $"A8 was '{a8}'; status: {statusAfter}");
        // One assertion, so a run reports all three members at once rather than stopping at the first.
        var volatileTwice = a10 is double firstCount && a10Again is double laterCount && laterCount > firstCount;
        Assert.True(
            volatileTwice && a11 as string == "A11" && a12 as string == "A12",
            $"Volatile, Caller, and ThisCell in a UDF: A10 was '{a10}' then '{a10Again}' (Volatile wants the second larger), A11 was '{a11}' (wants A11), A12 was '{a12}' (wants A12); status: {statusAfter}");
        Assert.True(a9 is double v9 && v9 == 6, $"A9 was '{a9}'; reload status: {reloadStatus}");
        Assert.True(statusBefore!.Contains("collation: NLS", StringComparison.Ordinal), "collation: " + statusBefore);
        Assert.True(a3 is null, $"A3 was '{a3}' before closing");
        Assert.True(statusAfter!.Contains("bound workbooks: none", StringComparison.Ordinal), "status after close: " + statusAfter);
    }

    /// <summary>
    /// The quickstart of README.md, step by step: build samples/Quickstart, open its workbook so
    /// the add-in binds the project next to it (D4), and run the macro the Form control's OnAction
    /// names. B2 carries the greeting and the same line reaches the output vbang status reports.
    /// </summary>
    [Fact]
    public void Quickstart_ButtonMacro_WritesTheCellAndPrints()
    {
        var addIn = RequireAddIn();
        var sample = Path.Combine(RepositoryRoot(), "samples", "Quickstart");
        var sampleDir = Path.Combine(workDir, "Quickstart");
        Directory.CreateDirectory(sampleDir);
        var workbookPath = Path.Combine(sampleDir, "Quickstart.xlsx");
        File.Copy(Path.Combine(sample, "Quickstart.xlsx"), workbookPath);
        var projectDir = Path.Combine(sampleDir, "Quickstart" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        foreach (var file in Directory.GetFiles(Path.Combine(sample, "Quickstart" + ProjectPaths.FolderSuffix)))
        {
            File.Copy(file, Path.Combine(projectDir, Path.GetFileName(file)));
        }

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        string? b2 = null, status = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            excel.RunMacro("Greeting.SayHello", attempts: 12);
            b2 = ExcelInstance.CellValue(workbook, "Sheet1", "B2") as string;
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
        });

        Assert.True(b2 == "Hello, world", $"B2 was '{b2}'; status: {status}");
        var printed = status![(status!.IndexOf("recent output:", StringComparison.Ordinal) + "recent output:".Length)..];
        Assert.True(printed.Contains("Hello, world", StringComparison.Ordinal), "status: " + status);
    }

    /// <summary>Polls vbang.Status until it satisfies the condition or the timeout passes; returns the last status either way.</summary>
    private static string WaitForStatus(ExcelInstance excel, Func<string, bool> ready, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        string status;
        do
        {
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            if (ready(status))
            {
                return status;
            }

            Thread.Sleep(500);
        }
        while (DateTime.UtcNow < deadline);

        return status;
    }

    /// <summary>
    /// The CLI finds Excel through the Running Object Table or the workbook window, so it would
    /// reach another Excel before the throwaway one; this test only runs when none is open.
    /// </summary>
    [Fact]
    public void Cli_TestAndRun_ReachTheRunningExcel()
    {
        var addIn = RequireAddIn();
        var cli = RequireCli();
        Assert.SkipWhen(Process.GetProcessesByName("EXCEL").Length > 0, "Another Excel is running; the CLI would find it instead of the throwaway instance (CLAUDE.md R17).");
        var projectDir = CopySample("Procedural");
        var failingDir = WriteFailingProject();
        var bigDir = WriteBigOutputProject();
        var junitPath = Path.Combine(workDir, "cli", "procedural.xml");

        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);

            var test = RunCli(cli, "test", "--project", projectDir, "--junit", junitPath);
            Assert.True(test.ExitCode == 0, test.ToString());
            Assert.Contains("passed  Tests.Classify_UsesGoSub", test.Output, StringComparison.Ordinal);
            Assert.Contains("12 tests: 12 passed, 0 failed, 0 errors.", test.Output, StringComparison.Ordinal);
            Assert.True(File.Exists(junitPath), "JUnit report not written: " + junitPath);

            var failing = RunCli(cli, "test", "--project", failingDir);
            Assert.True(failing.ExitCode == 5, failing.ToString());
            Assert.Contains("FAILED  Tests.Fails", failing.Output, StringComparison.Ordinal);
            Assert.Contains("Assert.AreEqual failed. Expected:<3>. Actual:<2>. sum", failing.Output, StringComparison.Ordinal);

            var run = RunCli(cli, "run", "Main.Main", "--project", projectDir);
            Assert.True(run.ExitCode == 0, run.ToString());
            Assert.Contains("Distance from origin: 5", run.Output, StringComparison.Ordinal);

            // What the run printed reached the project's log, which vbang logs prints (ROADMAP.md WP4).
            var logs = RunCli(cli, "logs", "--project", projectDir);
            Assert.True(logs.ExitCode == 0, logs.ToString());
            Assert.Contains("Distance from origin: 5", logs.Output, StringComparison.Ordinal);

            var json = RunCli(cli, "test", "--project", failingDir, "--json");
            Assert.True(json.ExitCode == 5, json.ToString());
            Assert.False(RunResponse.FromJson(json.Output).Ok);

            var big = RunCli(cli, "run", "Main.Main", "--project", bigDir);
            Assert.True(big.ExitCode == 0, big.ToString());
            Assert.Equal(600, big.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length);
            Assert.True(File.Exists(ProjectPaths.ResponsePath(bigDir)), "the long response should have gone through out/response.json");
        });
    }

    /// <summary>
    /// ROADMAP.md WP4: an error a button's macro leaves unhandled shows VBA's run-time error
    /// dialog (its title, "Run-time error '5':", a blank line, the description, and End), and the
    /// project's out/output.log has what the macro printed and the error. The dialog is answered
    /// from another thread, as a person clicks End.
    /// </summary>
    [Fact]
    public void Workbook_UnhandledError_ShowsVbasDialogAndReachesTheLog()
    {
        var addIn = RequireAddIn();
        var folder = Path.Combine(workDir, "Errors");
        Directory.CreateDirectory(folder);
        var workbookPath = Path.Combine(folder, "Errors.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Hello", "Hello.xlsx"), workbookPath);
        var projectDir = Path.Combine(folder, "Errors" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), "Attribute VB_Name = \"Module1\"\r\nPublic Sub Fails()\r\n    Debug.Print \"before\"\r\n    Err.Raise 5, , \"custom\"\r\nEnd Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        string[]? dialog = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            var answer = Task.Run(() => excel.AnswerDialog("Microsoft Visual Basic", "End", TimeSpan.FromSeconds(60)));
            excel.RunMacro("Module1.Fails", attempts: 12);
            dialog = answer.Result;
            ExcelInstance.CloseWorkbook(workbook);
        });

        Assert.NotNull(dialog);
        Assert.Contains("Run-time error '5':\n\ncustom", dialog);
        var log = File.ReadAllText(ProjectPaths.OutputLogPath(projectDir));
        Assert.Contains(" before", log, StringComparison.Ordinal);
        Assert.Contains("Run-time error in Module1.Fails: '5': custom", log, StringComparison.Ordinal);
    }

    /// <summary>
    /// ROADMAP.md WP4, ARCHITECTURE.md D17: a workbook that still has its VBA project is not
    /// bound to the project folder next to it, so no macro runs twice, and vba-ng says why,
    /// naming Save As .xlsx. The workbook is made here, with a module imported into
    /// its VBA project, so the test owns it and does not reach into the Import tests fixture.
    /// This Excel is hidden, so the notice goes to the log without a dialog: a modal box nobody
    /// can see blocked the automation call that opened the workbook until Excel was killed.
    /// </summary>
    [Fact]
    public void Workbook_WithItsVbaProject_IsNotBound()
    {
        var addIn = RequireAddIn();
        var folder = Path.Combine(workDir, "Legacy");
        Directory.CreateDirectory(folder);
        var workbookPath = Path.Combine(folder, "Legacy.xlsm");
        var projectDir = Path.Combine(folder, "Legacy" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), "Attribute VB_Name = \"Module1\"\r\nPublic Sub Hello()\r\nEnd Sub\r\n");
        var nativeModule = Path.Combine(folder, "Native.bas");
        File.WriteAllText(nativeModule, "Attribute VB_Name = \"Native\"\r\nPublic Sub NativeHello()\r\nEnd Sub\r\n");
        Assert.True(ProjectCompiler.Build(projectDir).Success);

        string[]? dialog = null;
        string? status = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var made = excel.ActiveWorkbook();
            ExcelInstance.SaveAs(made, workbookPath, 52);
            ExcelInstance.ImportModule(made, nativeModule);
            ExcelInstance.Save(made);
            ExcelInstance.CloseWorkbook(made);

            // A dialog, if one opened, would open inside Open; the watcher clicks it so the call returns either way.
            var answer = Task.Run(() => excel.AnswerDialog("vba-ng", "OK", TimeSpan.FromSeconds(10), killIfAbsent: false));
            var workbook = excel.OpenWorkbook(workbookPath);
            dialog = answer.Result;
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.CloseWorkbook(workbook);
        });

        Assert.True(dialog is null, "a dialog opened in a hidden Excel: " + string.Join(" | ", dialog ?? []));
        Assert.Contains("Save a copy as an Excel Workbook (.xlsx)", status, StringComparison.Ordinal);
        Assert.Contains("bound workbooks: none", status, StringComparison.Ordinal);
    }

    /// <summary>
    /// WP4: <c>Application.OnTime</c> reaches a registered command, and the workbook-qualified
    /// <c>Application.Run "Book.xlsx!Module.Proc"</c> resolves in process when the prefix names
    /// the project's own workbook (D23). Naming another loaded project's workbook does not: the
    /// name falls through to Excel, which has no VBA project to resolve it against, so a
    /// workbook calling into another workbook by name is a gap the alpha carries (ROADMAP.md WP4).
    /// </summary>
    [Fact]
    public void Commands_OnTimeAndQualifiedRun_ReachRegisteredCommands()
    {
        var addIn = RequireAddIn();
        var sampleDir = Path.Combine(workDir, "Timed");
        Directory.CreateDirectory(sampleDir);
        var workbookPath = Path.Combine(sampleDir, "Timed.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Sheets", "Sheets.xlsx"), workbookPath);
        var projectDir = Path.Combine(sampleDir, "Timed" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "Sched.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Sched\"",
            "Option Explicit",
            "Public Sub StartTimer()",
            "    Application.OnTime Now + TimeSerial(0, 0, 1), \"Landed\"",
            "End Sub",
            "Public Sub Landed()",
            "    Worksheets(\"Sheet1\").Range(\"H1\").Value = \"ontime\"",
            "End Sub",
            "Public Sub Marked()",
            "    Worksheets(\"Sheet1\").Range(\"H2\").Value = \"qualified\"",
            "End Sub",
            "Public Sub CallQualified()",
            "    Application.Run \"Timed.xlsx!Sched.Marked\"",
            "End Sub",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        object? h1 = null, h2 = null;
        string? qualifiedError = null, status = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            excel.RunMacro("Sched.StartTimer");

            // OnTime fires when Excel is next idle, so wait for the second it was asked for.
            for (var attempt = 0; attempt < 40 && h1 as string != "ontime"; attempt++)
            {
                Thread.Sleep(250);
                h1 = ExcelInstance.CellValue(workbook, "Sheet1", "H1");
            }

            // The workbook-qualified form from inside the project, which the runtime resolves
            // in process once the prefix names its own workbook (D23).
            try
            {
                excel.RunMacro("Sched.CallQualified");
            }
            catch (COMException ex)
            {
                qualifiedError = ex.Message;
            }

            h2 = ExcelInstance.CellValue(workbook, "Sheet1", "H2");
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.Release(workbook);
        });

        // One assertion each, so a run says which of the two reached its command.
        Assert.True(h1 as string == "ontime", $"Application.OnTime: H1 was '{h1}'; status: {status}");
        Assert.True(h2 as string == "qualified", $"Application.Run \"Book.xlsx!Module.Proc\": H2 was '{h2}', error '{qualifiedError}'; status: {status}");
    }

    /// <summary>
    /// WP4: two loaded projects claiming one name follow the section 8 policy. XLL registration
    /// is global to the Excel session while VBA scopes macros to a workbook, so the project that
    /// loads first keeps the bare name, the second is reachable as <c>Project.Name</c> and as
    /// <c>Module.Name</c>, and the collision is logged rather than left silent.
    /// </summary>
    [Fact]
    public void Commands_TwoProjectsClaimingOneName_WarnAndQualify()
    {
        var addIn = RequireAddIn();
        var marker = Path.Combine(workDir, "claimed.txt");
        var firstWorkbook = MakeClaimingProject("First", "ModA", marker, "first");
        var secondWorkbook = MakeClaimingProject("Second", "ModB", marker, "second");

        string? status = null, bareError = null, projectQualifiedError = null, moduleQualifiedError = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var first = excel.OpenWorkbook(firstWorkbook);
            var second = excel.OpenWorkbook(secondWorkbook);

            // The bare name belongs to the project that loaded first; the second answers to both
            // of its qualified names.
            try { excel.RunMacro("Common"); } catch (COMException ex) { bareError = ex.Message; }
            try { excel.RunMacro("Second.Common"); } catch (COMException ex) { projectQualifiedError = ex.Message; }
            try { excel.RunMacro("ModB.Common"); } catch (COMException ex) { moduleQualifiedError = ex.Message; }

            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.Release(second);
            ExcelInstance.Release(first);
        });

        var claimed = File.Exists(marker)
            ? File.ReadAllText(marker).Replace("\r\n", " ", StringComparison.Ordinal).Trim()
            : "(no file)";

        // One assertion, so a run reports every name at once rather than stopping at the first.
        Assert.True(
            claimed == "first second second",
            $"Two projects, one name: the marker held '{claimed}' (wants 'first second second'), errors '{bareError}' / '{projectQualifiedError}' / '{moduleQualifiedError}'; status: {status}");
        Assert.Contains("already registered by another project", status, StringComparison.Ordinal);
    }

    /// <summary>
    /// WP7: <c>vbang init</c> starts a project beside a workbook (ARCHITECTURE.md section 8, D11):
    /// the manifest with the references an Excel project carries, a <c>.gitignore</c> for the build
    /// output, and with <c>--vscode</c> the build task and the attach configuration. No Excel is
    /// needed; what it writes has to build, which is the whole point of the command.
    /// </summary>
    [Fact]
    public void Cli_Init_StartsAProjectThatBuilds()
    {
        var cli = RequireCli();
        var dir = Path.Combine(workDir, "Init");
        Directory.CreateDirectory(dir);
        var workbookPath = Path.Combine(dir, "Budget.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Quickstart", "Quickstart.xlsx"), workbookPath);

        var result = RunCli(cli, "init", workbookPath, "--vscode");

        Assert.Equal(0, result.ExitCode);
        var projectDir = Path.Combine(dir, "Budget" + ProjectPaths.FolderSuffix);
        var manifest = File.ReadAllText(Path.Combine(projectDir, "vbang.json"));
        Assert.Contains("\"name\": \"Budget\"", manifest, StringComparison.Ordinal);
        Assert.Contains("{ \"name\": \"Excel\" }", manifest, StringComparison.Ordinal);
        Assert.Contains("{ \"name\": \"Office\" }", manifest, StringComparison.Ordinal);
        Assert.Equal("out/\n", File.ReadAllText(Path.Combine(projectDir, ".gitignore")).ReplaceLineEndings("\n"));
        Assert.Contains("$msCompile", File.ReadAllText(Path.Combine(dir, ".vscode", "tasks.json")), StringComparison.Ordinal);
        Assert.Contains("EXCEL.EXE", File.ReadAllText(Path.Combine(dir, ".vscode", "launch.json")), StringComparison.Ordinal);

        // A project is never overwritten, so a second init on the same workbook is a usage error.
        Assert.Equal(2, RunCli(cli, "init", workbookPath).ExitCode);

        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Module1\"",
            "Option Explicit",
            "Public Function Gross(ByVal net As Currency) As Currency",
            "    Gross = net * 1.2",
            "End Function",
            string.Empty,
        ]));

        Assert.SkipUnless(OfficeTypeLibrariesRegistered(), "The project init starts references Excel and Office, whose type libraries are not registered here.");
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
    }

    /// <summary>
    /// WP7: <c>vbang report</c> writes the bundle a bug report carries. What it leaves out is the
    /// point as much as what it holds: never the workbook, so the bundle can go on a public issue
    /// without the data going with it.
    /// </summary>
    [Fact]
    public void Cli_Report_WritesABundleWithoutTheWorkbook()
    {
        var cli = RequireCli();
        var dir = Path.Combine(workDir, "Report");
        Directory.CreateDirectory(dir);
        var workbookPath = Path.Combine(dir, "Budget.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Quickstart", "Quickstart.xlsx"), workbookPath);
        Assert.Equal(0, RunCli(cli, "init", workbookPath).ExitCode);
        var projectDir = Path.Combine(dir, "Budget" + ProjectPaths.FolderSuffix);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Module1\"",
            "Option Explicit",
            "Public Function Gross(ByVal net As Currency) As Currency",
            "    Gross = net * 1.2",
            "End Function",
            string.Empty,
        ]));

        // Run where the report lands, and with no Excel of its own: an unreachable Excel is a line
        // in the bundle, not a failure, because a tester reporting a bug may have closed it.
        var result = RunCliIn(cli, dir, "report", projectDir);

        Assert.Equal(0, result.ExitCode);
        var bundle = Assert.Single(Directory.GetFiles(dir, "vbang-report-Budget-*.zip"));
        using var zip = System.IO.Compression.ZipFile.OpenRead(bundle);
        var entries = zip.Entries.Select(e => e.FullName).ToList();
        Assert.Contains("versions.txt", entries, StringComparer.Ordinal);
        Assert.Contains("build.log", entries, StringComparer.Ordinal);
        Assert.Contains("vbang.json", entries, StringComparer.Ordinal);
        Assert.Contains("source/Module1.bas", entries, StringComparer.Ordinal);
        if (OfficeTypeLibrariesRegistered())
        {
            // Generated C# exists only when the build succeeds, which needs the references init writes.
            Assert.Contains("out/gen/Module1.cs", entries, StringComparer.Ordinal);
        }

        Assert.DoesNotContain(entries, e => e.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || e.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

        using var versions = new StreamReader(zip.GetEntry("versions.txt")!.Open());
        var text = versions.ReadToEnd();
        Assert.Contains("vbang: ", text, StringComparison.Ordinal);
        Assert.Contains("culture: ", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// An import never overwrites: not a project folder, whose sources may have been edited since the last import, and not
    /// the macro-free copy, which may have been worked in. Both are checked before anything is written. Needs no Excel.
    /// </summary>
    [Fact]
    public void Cli_Import_NeverOverwritesAProjectOrACopy()
    {
        var cli = RequireCli();
        var dir = Path.Combine(workDir, "Reimport");
        Directory.CreateDirectory(dir);
        var macroBook = Path.Combine(dir, "Legacy.xlsm");
        File.Copy(Path.Combine(RepositoryRoot(), "tests", "VbaNg.Import.Tests", "Fixtures", "Legacy.xlsm"), macroBook);
        Assert.Equal(0, RunCli(cli, "import", macroBook).ExitCode);
        var edited = Path.Combine(dir, "Legacy" + ProjectPaths.FolderSuffix, "Sales.bas");
        File.AppendAllText(edited, "' edited after the import\r\n");

        var again = RunCli(cli, "import", macroBook);

        Assert.True(again.ExitCode == 2, again.ToString());
        Assert.Contains("already holds a project", again.Error, StringComparison.Ordinal);
        Assert.EndsWith("' edited after the import\r\n", File.ReadAllText(edited), StringComparison.Ordinal);

        var copy = Path.Combine(dir, "Legacy.xlsx");
        File.WriteAllText(copy, "worked in");
        var elsewhere = Path.Combine(dir, "Other" + ProjectPaths.FolderSuffix);

        var toXlsx = RunCli(cli, "import", macroBook, elsewhere, "--to-xlsx");

        Assert.True(toXlsx.ExitCode == 2, toXlsx.ToString());
        Assert.Contains("already exists", toXlsx.Error, StringComparison.Ordinal);
        Assert.Equal("worked in", File.ReadAllText(copy));
        Assert.False(Directory.Exists(elsewhere), "the import wrote sources before it refused");
    }

    /// <summary>
    /// Excel enables macros in a workbook automation opens, so --to-xlsx would run the Workbook_Open of the workbook it
    /// reads unless it disables them first. The workbook here writes a marker file when it opens.
    /// </summary>
    [Fact]
    public void Import_ToXlsx_RunsNoMacroOfTheWorkbook()
    {
        var addIn = RequireAddIn();
        var cli = RequireCli();
        var dir = Path.Combine(workDir, "NoMacros");
        Directory.CreateDirectory(dir);
        var macroBook = Path.Combine(dir, "Opens.xlsm");
        var marker = Path.Combine(dir, "opened.txt");
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var made = excel.ActiveWorkbook();
            ExcelInstance.SaveAs(made, macroBook, 52);
            ExcelInstance.AddWorkbookCode(made, $"Private Sub Workbook_Open()\r\n    Dim f As Integer\r\n    f = FreeFile\r\n    Open \"{marker}\" For Output As #f\r\n    Close #f\r\nEnd Sub\r\n");
            ExcelInstance.Save(made);
            ExcelInstance.CloseWorkbook(made);
        });

        var imported = RunCli(cli, "import", macroBook, "--to-xlsx");

        Assert.True(imported.ExitCode == 0, imported.ToString());
        Assert.True(File.Exists(Path.Combine(dir, "Opens.xlsx")), imported.ToString());
        Assert.False(File.Exists(marker), "--to-xlsx ran the workbook's Workbook_Open");
    }

    /// <summary>
    /// WP6: the M4 exit criterion shown through the importer. The fixture workbook is imported with
    /// <c>--to-xlsx</c>, the project that comes out is built, and the macro-free copy is opened in
    /// Excel, where the workbook's own events, both kinds of button, and a UDF all work with no
    /// VBA project left in the file.
    /// </summary>
    [Fact]
    public void Import_LegacyFixture_RoundTripsThroughExcel()
    {
        var addIn = RequireAddIn();
        var cli = RequireCli();
        var dir = Path.Combine(workDir, "RoundTrip");
        Directory.CreateDirectory(dir);
        var macroBook = Path.Combine(dir, "Legacy.xlsm");
        File.Copy(Path.Combine(RepositoryRoot(), "tests", "VbaNg.Import.Tests", "Fixtures", "Legacy.xlsm"), macroBook);

        var imported = RunCli(cli, "import", macroBook, "--to-xlsx");
        Assert.Equal(0, imported.ExitCode);

        var workbookPath = Path.Combine(dir, "Legacy.xlsx");
        var projectDir = Path.Combine(dir, "Legacy" + ProjectPaths.FolderSuffix);
        Assert.True(File.Exists(workbookPath), imported.ToString());
        Assert.True(Directory.Exists(projectDir), imported.ToString());

        // The copy is macro-free, and it kept the controls: a Form control lives in the sheet's VML
        // and an ActiveX control in its own part.
        var parts = System.IO.Compression.ZipFile.OpenRead(workbookPath);
        using (parts)
        {
            var names = parts.Entries.Select(e => e.FullName).ToList();
            Assert.DoesNotContain(names, n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(names, n => n.StartsWith("xl/drawings/vmlDrawing", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(names, n => n.StartsWith("xl/activeX/", StringComparison.OrdinalIgnoreCase));
        }

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        // The UserForm is the one thing that did not come over whole, and it is a warning (D-D).
        Assert.Contains(build.Diagnostics, d => !d.IsError && d.FilePath.EndsWith("Dialog.frm", StringComparison.OrdinalIgnoreCase));

        object? a1 = null, b1 = null, c1 = null, d1 = null, e1 = null;
        string? status = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);

            // Workbook_Open, which the add-in runs itself because the workbook's own event fired
            // before the sink existed (ARCHITECTURE.md section 7).
            a1 = ExcelInstance.CellValue(workbook, "Sheet1", "A1");

            ExcelInstance.SetCellValue(workbook, "Sheet1", "A2", 5);
            b1 = ExcelInstance.CellValue(workbook, "Sheet1", "B1");

            ExcelInstance.ClickActiveXButton(workbook, "Sheet1", "CommandButton1");
            c1 = ExcelInstance.CellValue(workbook, "Sheet1", "C1");

            // The name the Form control's OnAction carries, which registration answers.
            excel.RunMacro("Sales.ButtonClick");
            d1 = ExcelInstance.CellValue(workbook, "Sheet1", "D1");

            ExcelInstance.SetCellValue(workbook, "Sheet1", "E1", "=Gross(100)");
            e1 = ExcelInstance.CellValue(workbook, "Sheet1", "E1");

            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.Release(workbook);
        });

        // One assertion, so a run reports every part of the round trip at once.
        Assert.True(
            a1 as string == "opened" && b1 as string == "changed" && c1 as string == "activex" && d1 as string == "form button" && e1 is double and 120,
            $"Round trip: A1 '{a1}' (opened), B1 '{b1}' (changed), C1 '{c1}' (activex), D1 '{d1}' (form button), E1 '{e1}' (120); status: {status}");
    }

    /// <summary>
    /// WP4: a class instance of the project handed to a COM method and read back out of it.
    /// ROADMAP.md WP4 records this as raising 13, so rather than assert a number the same source
    /// runs as real VBA and under vbang and the two reports are compared (R1); the probe traps
    /// its own errors, so whatever VBA does is what vbang has to do.
    /// </summary>
    [Fact]
    public void HostRun_ClassInstanceThroughCom_MatchesVba()
    {
        var addIn = RequireAddIn();
        var projectDir = Path.Combine(workDir, "ComObjects" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "Thing.cls"), string.Join("\r\n", [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"Thing\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
            "Public Value As Long",
            string.Empty,
        ]));
        File.WriteAllText(Path.Combine(projectDir, "Probe.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Probe\"",
            "Option Explicit",
            "Public Function Report() As String",
            "    Dim out As String, d As Object, t As Thing, back As Object",
            "    On Error Resume Next",
            "    Set d = CreateObject(\"Scripting.Dictionary\")",
            "    Set t = New Thing",
            "    t.Value = 7",
            "    d.Add \"k\", t",
            "    out = \"add: \" & Err.Number & \" \" & Err.Description",
            "    Err.Clear",
            "    Set back = d(\"k\")",
            "    out = out & vbLf & \"back: \" & Err.Number & \" \" & TypeName(back)",
            "    Err.Clear",
            "    out = out & vbLf & \"value: \" & back.Value & \" \" & Err.Number",
            "    Err.Clear",
            "    out = out & vbLf & \"same: \" & (back Is t) & \" \" & Err.Number",
            "    Report = out",
            "End Function",
            "Public Sub Go()",
            "    Debug.Print Report();",
            "End Sub",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        var vba = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.ActiveWorkbook();
            ExcelInstance.ImportModule(workbook, Path.Combine(projectDir, "Thing.cls"));
            ExcelInstance.ImportModule(workbook, Path.Combine(projectDir, "Probe.bas"));
            vba = excel.RunMacro("'" + ExcelInstance.Name(workbook) + "'!Probe.Report") as string ?? string.Empty;
            ExcelInstance.Release(workbook);
        });

        var json = string.Empty;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            json = excel.RunHostCommand("vbang.Run", CommandTimeout, projectDir, "Probe.Go", string.Empty);
        });

        var response = RunResponse.FromJson(json);
        Assert.True(response.Ok, response.Error + Environment.NewLine + response.Detail);

        // Not a match of two failures: VBA really did hold the instance and hand it back.
        Assert.Contains("back: 0 Thing", vba, StringComparison.Ordinal);
        Assert.Contains("value: 7 0", vba, StringComparison.Ordinal);
        Assert.Equal(ReportLines(vba), ReportLines(response.Output));
    }

    /// <summary>
    /// WP4: <c>End</c> closes the files the project opened and returns its module state to its
    /// initial values, and <c>Stop</c> breaks only under a debugger, so a command that runs past
    /// one finishes. Both stop the golden recorder itself (ROADMAP.md WP5), so they are measured
    /// in an Excel of their own here.
    /// </summary>
    [Fact]
    public void Program_EndAndStop_MatchVba()
    {
        var addIn = RequireAddIn();
        var dir = Path.Combine(workDir, "Program");
        Directory.CreateDirectory(dir);
        var workbookPath = Path.Combine(dir, "Program.xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Sheets", "Sheets.xlsx"), workbookPath);
        var projectDir = Path.Combine(dir, "Program" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        var opened = Path.Combine(dir, "opened.txt");
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");
        File.WriteAllText(Path.Combine(projectDir, "Prog.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Prog\"",
            "Option Explicit",
            "Public Counter As Long",
            "Public Sub Prime()",
            "    Counter = 42",
            "End Sub",
            "Public Sub EndNow()",
            "    Dim f As Integer",
            "    f = FreeFile",
            "    Open \"" + opened + "\" For Output As #f",
            "    Print #f, \"open\"",
            "    End",
            "End Sub",
            "Public Sub PastStop()",
            "    Stop",
            "    Counter = 7",
            "End Sub",
            "Public Function Report() As String",
            "    Report = \"counter \" & Counter",
            "End Function",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        string? primed = null, afterEnd = null, afterStop = null, status = null;
        var fileClosed = false;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            excel.RunMacro("Prog.Prime");
            primed = excel.RunMacro("Prog.Report") as string;

            excel.RunMacro("Prog.EndNow");
            afterEnd = excel.RunMacro("Prog.Report") as string;

            // Excel is still up, so the handle is the project's: VBA's End closes it.
            try
            {
                using var probe = new FileStream(opened, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                fileClosed = true;
            }
            catch (IOException)
            {
                fileClosed = false;
            }

            excel.RunMacro("Prog.PastStop");
            afterStop = excel.RunMacro("Prog.Report") as string;
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.Release(workbook);
        });

        // One assertion, so a run reports all three at once rather than stopping at the first.
        Assert.True(
            primed == "counter 42" && afterEnd == "counter 0" && fileClosed && afterStop == "counter 7",
            $"End and Stop: primed '{primed}' (wants counter 42), after End '{afterEnd}' (wants counter 0), the file End left open was closed: {fileClosed} (wants True), after Stop '{afterStop}' (wants counter 7); status: {status}");
    }

    /// <summary>
    /// A project whose one Sub appends a word to a file, so which of two same-named commands ran
    /// is visible from outside Excel. The workbook is named after the folder, which is what binds
    /// them (D4).
    /// </summary>
    private string MakeClaimingProject(string name, string module, string marker, string word)
    {
        var dir = Path.Combine(workDir, name);
        Directory.CreateDirectory(dir);
        var workbookPath = Path.Combine(dir, name + ".xlsx");
        File.Copy(Path.Combine(RepositoryRoot(), "samples", "Sheets", "Sheets.xlsx"), workbookPath);
        var projectDir = Path.Combine(dir, name + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"Excel\" } ] }");

        // VBA has no escapes in a string literal, so the path goes in as it is.
        File.WriteAllText(Path.Combine(projectDir, module + ".bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"" + module + "\"",
            "Option Explicit",
            "Public Sub Common()",
            "    Dim f As Integer",
            "    f = FreeFile",
            "    Open \"" + marker + "\" For Append As #f",
            "    Print #f, \"" + word + "\"",
            "    Close #f",
            "End Sub",
            string.Empty,
        ]));

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        return workbookPath;
    }

    private static string RequireAddIn()
    {
        Assert.SkipWhen(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(GateVariable)), $"Set {GateVariable}=1 to run the Excel-driving tests.");
        var path = Path.Combine(RepositoryRoot(), "src", "VbaNg.AddIn", "bin", Configuration(), "net10.0-windows", "VbaNg.AddIn-AddIn64.xll");
        Assert.True(File.Exists(path), "Add-in not built: " + path);
        return path;
    }

    /// <summary>
    /// Whether Excel's and Office's type libraries are registered, which a project <c>vbang init</c>
    /// starts needs in order to build. A machine without Office, such as a CI runner, has neither.
    /// </summary>
    private static bool OfficeTypeLibrariesRegistered() =>
        TypeLibraryCache.Resolve("Excel", null, null) is not null && TypeLibraryCache.Resolve("Office", null, null) is not null;

    private static string RequireCli()
    {
        var path = Path.Combine(RepositoryRoot(), "src", "VbaNg.Cli", "bin", Configuration(), "net10.0-windows", "vbang.exe");
        Assert.True(File.Exists(path), "CLI not built: " + path);
        return path;
    }

    private static string Configuration() =>
        typeof(SampleTests).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? "Debug";

    private string BuildSample(string name)
    {
        var projectDir = CopySample(name);
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        return projectDir;
    }

    private string CopySample(string name)
    {
        var source = Path.Combine(RepositoryRoot(), "samples", name, name + ProjectPaths.FolderSuffix);
        var target = Path.Combine(workDir, name + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(target);
        foreach (var file in Directory.GetFiles(source, "*.bas").Concat(Directory.GetFiles(source, "*.cls")).Concat(Directory.GetFiles(source, "vbang.json")))
        {
            File.Copy(file, Path.Combine(target, Path.GetFileName(file)));
        }

        return target;
    }

    /// <summary>Six hundred lines of eighty characters: well past the 32,767 characters Application.Run can return, so the response travels through out/response.json.</summary>
    private string WriteBigOutputProject()
    {
        var projectDir = Path.Combine(workDir, "Big" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(
            Path.Combine(projectDir, "Main.bas"),
            "Attribute VB_Name = \"Main\"\r\nOption Explicit\r\n\r\nPublic Sub Main()\r\n    Dim i As Long\r\n    For i = 1 To 600\r\n        Debug.Print String(80, \"x\")\r\n    Next i\r\nEnd Sub\r\n");
        return projectDir;
    }

    private string WriteFailingProject()
    {
        var projectDir = Path.Combine(workDir, "Failing" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"vbang\" } ] }");
        File.WriteAllText(
            Path.Combine(projectDir, "Tests.bas"),
            "Attribute VB_Name = \"Tests\"\r\nOption Explicit\r\n\r\n'@Test\r\nPublic Sub Passes()\r\n    Assert.IsTrue True\r\nEnd Sub\r\n\r\n'@Test\r\nPublic Sub Fails()\r\n    Assert.AreEqual 3, 1 + 1, \"sum\"\r\nEnd Sub\r\n");
        return projectDir;
    }

    private static CliResult RunCli(string cli, params string[] arguments) => RunCliIn(cli, null, arguments);

    /// <summary>The CLI in a folder of its own, for a command that writes where it is run (vbang report).</summary>
    private static CliResult RunCliIn(string cli, string? workingDirectory, params string[] arguments)
    {
        var info = new ProcessStartInfo(cli)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workingDirectory ?? Directory.GetCurrentDirectory(),
        };
        foreach (var argument in arguments)
        {
            info.ArgumentList.Add(argument);
        }

        using var process = Process.Start(info) ?? throw new InvalidOperationException("vbang did not start.");
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(CommandTimeout))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("vbang " + string.Join(' ', arguments) + " did not finish within " + CommandTimeout);
        }

        return new CliResult(process.ExitCode, output.Result, error.Result);
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

    private sealed record CliResult(int ExitCode, string Output, string Error)
    {
        public override string ToString() => $"exit {ExitCode}{Environment.NewLine}stdout:{Environment.NewLine}{Output}{Environment.NewLine}stderr:{Environment.NewLine}{Error}";
    }
}
