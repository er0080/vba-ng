using System.Reflection.Metadata;
using System.Xml.Linq;

using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// End-to-end tests for the build pipeline with no Excel involved (CLAUDE.md R9): build the Hello
/// sample, load it through the same hosting code the add-in uses, run it, and inspect the PDB.
/// Tests in this class share <see cref="Host.Current"/>, so they must not run in parallel with
/// each other; xunit runs the tests of one class sequentially.
/// </summary>
public sealed class PipelineTests : IDisposable
{
    private readonly string workDir = Path.Combine(Path.GetTempPath(), "vbang-tests", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try
        {
            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
            // Shadow-copied assemblies stay mapped until their load context is collected.
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Fact]
    public void Build_HelloSample_ProducesAssemblySymbolsAndGeneratedSource()
    {
        var projectDir = CopySample("Hello");

        var result = ProjectCompiler.Build(projectDir);

        Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics));
        Assert.Equal("Hello", result.ProjectName);
        Assert.True(File.Exists(ProjectPaths.AssemblyPath(projectDir)));
        Assert.True(File.Exists(ProjectPaths.SymbolsPath(projectDir)));
        Assert.True(File.Exists(ProjectPaths.BuildInfoPath(projectDir)));
        var generated = File.ReadAllText(Path.Combine(ProjectPaths.GeneratedDir(projectDir), "Hello.cs"));
        Assert.Contains("#line 10 \"" + Path.Combine(projectDir, "Hello.bas") + "\"", generated, StringComparison.Ordinal);
    }

    /// <summary>
    /// out/build.json records every source as it is on disk, a form included, so the add-in can tell from the files alone
    /// that nothing changed. The compiler hashed a form after blanking its designer block, which no reader of the file
    /// could reproduce, so a workbook with a UserForm was rebuilt on every open.
    /// </summary>
    [Fact]
    public void Build_RecordsEachSourceAsItIsOnDisk()
    {
        var projectDir = Path.Combine(workDir, "Recorded" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        var form = Path.Combine(projectDir, "Dialog.frm");
        File.WriteAllText(form, string.Join("\r\n", [
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog ",
            "   Caption         =   \"Dialog\"",
            "End",
            "Attribute VB_Name = \"Dialog\"",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = False",
            "Public Sub Shout()",
            "End Sub",
            string.Empty,
        ]));
        var module = Path.Combine(projectDir, "Main.bas");
        File.WriteAllText(module, "Attribute VB_Name = \"Main\"\r\nPublic Sub Go()\r\nEnd Sub\r\n");

        Assert.True(ProjectCompiler.Build(projectDir).Success);

        using var info = System.Text.Json.JsonDocument.Parse(File.ReadAllText(ProjectPaths.BuildInfoPath(projectDir)));
        var inputs = info.RootElement.GetProperty("Inputs");
        foreach (var path in new[] { form, module })
        {
            var onDisk = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(File.ReadAllText(path))));
            Assert.Equal(onDisk, inputs.GetProperty(Path.GetFileName(path)).GetString());
        }
    }

    /// <summary>
    /// A project with a UserForm builds, so the rest of such a workbook can be tested (ROADMAP.md
    /// D-D): the form is a warning rather than an error, its code compiles as a class so callers
    /// bind, and entering that code raises a run-time error naming the gap.
    /// </summary>
    [Fact]
    public void Build_ProjectWithAUserForm_WarnsAndKeepsTheRestRunnable()
    {
        var projectDir = Path.Combine(workDir, "Forms" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);

        // No class header and no designer block, which is what the importer writes for a form.
        File.WriteAllText(Path.Combine(projectDir, "Dialog.frm"), string.Join("\r\n", [
            "Attribute VB_Name = \"Dialog\"",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = False",
            "Public Sub Shout()",
            "    Label1.Caption = \"x\"",
            "End Sub",
            string.Empty,
        ]));
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Option Explicit",
            "Public Sub Go()",
            "    Debug.Print \"the rest still runs\"",
            "End Sub",
            "Public Sub Touch()",
            "    On Error Resume Next",
            "    Dialog.Shout",
            "    Debug.Print Err.Number & \" \" & Err.Description",
            "End Sub",
            string.Empty,
        ]));

        var result = ProjectCompiler.Build(projectDir);

        Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics));
        var form = Assert.Single(result.Diagnostics, d => d.Id == DiagnosticIds.NotSupported);
        Assert.False(form.IsError, form.ToString());
        Assert.EndsWith("Dialog.frm", form.FilePath, StringComparison.OrdinalIgnoreCase);

        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));
        Assert.Equal(["the rest still runs"], RunCapturing(host, projectDir, "Main.Go"));

        // A control the form never declares binds, because nothing in the body is bound at all.
        var touched = RunCapturing(host, projectDir, "Main.Touch");
        Assert.StartsWith("438 UserForms are not supported yet", Assert.Single(touched), StringComparison.Ordinal);
    }

    /// <summary>
    /// <c>MsgBox</c> used as a value. The intrinsic is declared to return Long, so the generated C#
    /// has to be an int expression; it was a <c>Variant.FromInt32</c> wrapped around a
    /// Variant-returning method, which Roslyn rejected as VBA0003, an internal compiler error. No
    /// test built MsgBox as an expression until a corpus project did (ROADMAP.md WP6).
    /// </summary>
    [Fact]
    public void Build_MsgBoxAsAValue_Compiles()
    {
        var projectDir = Path.Combine(workDir, "Dialogs" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Option Explicit",
            "Public Sub Go()",
            "    Dim answer As Integer",
            "    answer = MsgBox(\"ok?\", 0, \"Title\")",
            "    Debug.Print answer",
            "End Sub",
            string.Empty,
        ]));

        var result = ProjectCompiler.Build(projectDir);

        Assert.True(result.Success, string.Join(Environment.NewLine, result.Diagnostics));
    }

    [Fact]
    public void Run_HelloSample_PrintsDebugOutput()
    {
        var projectDir = CopySample("Hello");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var output = RunCapturing(host, projectDir, "Hello.Main");

        Assert.Equal(["Hello from vba-ng", "Total = 15"], output);
    }

    /// <summary>The sample the locals window check runs (ARCHITECTURE.md section 9) fills every kind of native storage with the values its Stop shows; outside a debugger the Stop does nothing and the values print.</summary>
    [Fact]
    public void Run_StorageSample_FillsEveryKindOfStorage()
    {
        var projectDir = CopySample("Storage");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal([" 1  9 9/11/2026 2:30:00 PM 19.99 Truehelloab   42 forty-twotwo 20 d 125  125  2 widgetX-1   "], RunCapturing(host, projectDir, "Locals.Inspect"));
    }

    [Fact]
    public void Run_IsCaseInsensitive_LikeVba()
    {
        var projectDir = CopySample("Hello");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var output = RunCapturing(host, projectDir, "hello.MAIN");

        Assert.Equal("Hello from vba-ng", output[0]);
    }

    [Fact]
    public void Build_MapsEverySequencePointToTheBasFile()
    {
        var projectDir = CopySample("Hello");
        Assert.True(ProjectCompiler.Build(projectDir).Success);

        using var stream = File.OpenRead(ProjectPaths.SymbolsPath(projectDir));
        using var provider = MetadataReaderProvider.FromPortablePdbStream(stream);
        var reader = provider.GetMetadataReader();
        var mappedLines = new SortedSet<int>();
        foreach (var handle in reader.MethodDebugInformation)
        {
            var information = reader.GetMethodDebugInformation(handle);
            if (information.SequencePointsBlob.IsNil)
            {
                continue;
            }

            foreach (var point in information.GetSequencePoints())
            {
                if (point.IsHidden)
                {
                    continue;
                }

                var documentName = reader.GetString(reader.GetDocument(point.Document).Name);
                Assert.EndsWith("Hello.bas", documentName, StringComparison.OrdinalIgnoreCase);
                mappedLines.Add(point.StartLine);
            }
        }

        // For, body, Next, Print, Print, End Sub. Dim lines are not executable in VBA and get no sequence point.
        Assert.Equal([7, 8, 9, 10, 11, 12], mappedLines);
    }

    [Fact]
    public void Build_IsDeterministic()
    {
        var projectDir = CopySample("Hello");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        var firstAssembly = File.ReadAllBytes(ProjectPaths.AssemblyPath(projectDir));
        var firstSymbols = File.ReadAllBytes(ProjectPaths.SymbolsPath(projectDir));

        Assert.True(ProjectCompiler.Build(projectDir).Success);

        Assert.Equal(firstAssembly, File.ReadAllBytes(ProjectPaths.AssemblyPath(projectDir)));
        Assert.Equal(firstSymbols, File.ReadAllBytes(ProjectPaths.SymbolsPath(projectDir)));
    }

    [Fact]
    public void Load_ReloadsWhenTheAssemblyChanges()
    {
        var projectDir = CopySample("Hello");
        var source = Path.Combine(projectDir, "Hello.bas");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));
        Assert.Equal("Hello from vba-ng", RunCapturing(host, projectDir, "Hello.Main")[0]);

        File.WriteAllText(source, File.ReadAllText(source).Replace("Hello from vba-ng", "Hello again", StringComparison.Ordinal));
        Assert.True(ProjectCompiler.Build(projectDir).Success);

        Assert.Equal("Hello again", RunCapturing(host, projectDir, "Hello.Main")[0]);
        Assert.Single(host.Projects);
    }

    [Fact]
    public void Test_ProceduralSample_PassesEveryTest()
    {
        var projectDir = CopySample("Procedural");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var result = RunTests(host, projectDir);

        Assert.True(result.AllPassed, result.ToText());
        Assert.Equal(12, result.Results.Count);
        Assert.All(result.Results, r => Assert.Equal("Tests", r.Module));
        Assert.StartsWith("12 tests: 12 passed, 0 failed, 0 errors.", result.Summary, StringComparison.Ordinal);
        var report = XDocument.Parse(result.ToJUnitXml());
        Assert.Equal("Procedural", report.Root!.Attribute("name")!.Value);
        Assert.Equal("12", report.Root.Attribute("tests")!.Value);
        var suite = Assert.Single(report.Root.Elements("testsuite"));
        Assert.Equal("Tests", suite.Attribute("name")!.Value);
        Assert.Equal(12, suite.Elements("testcase").Count());
        Assert.Empty(suite.Descendants("failure"));
    }

    /// <summary>The M5 exit criterion (ROADMAP.md): the class-heavy sample passes without Excel.</summary>
    [Fact]
    public void Test_ClassesSample_PassesEveryTest()
    {
        var projectDir = CopySample("Classes");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var result = RunTests(host, projectDir);

        Assert.True(result.AllPassed, result.ToText());
        Assert.Equal(10, result.Results.Count);
        Assert.StartsWith("10 tests: 10 passed, 0 failed, 0 errors.", result.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Test_FailuresAndErrors_AreReportedPerProcedure()
    {
        var projectDir = Path.Combine(workDir, "Failing.vbang");
        Directory.CreateDirectory(projectDir);
        WriteVbangManifest(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Tests.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Tests\"",
            "Option Explicit",
            "'@Test",
            "Public Sub Passes()",
            "    Debug.Print \"hello\"",
            "    Assert.AreEqual 1, 1",
            "End Sub",
            "'@Test",
            "Public Sub Fails()",
            "    Debug.Print \"before\"",
            "    Assert.AreEqual 3, 1 + 1, \"sum\"",
            "    Debug.Print \"after\"",
            "End Sub",
            "'@Test",
            "Public Sub Errors()",
            "    Dim x As Long, zero As Long",
            "    x = 1 \\ zero",
            "End Sub",
            "'@Test",
            "Public Sub HandlerCannotHideAFailure()",
            "    On Error Resume Next",
            "    Assert.Fail \"must surface\"",
            "    Debug.Print \"not reached\"",
            "End Sub",
            "'@Test",
            "Public Sub SeesTheErrNumber()",
            "    On Error Resume Next",
            "    Err.Raise 5",
            "    Assert.AreEqual 5, Err.Number",
            "End Sub",
            "Public Sub NotATest()",
            "    Assert.Fail",
            "End Sub",
            string.Empty,
        ]));
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var result = RunTests(host, projectDir);

        Assert.Equal(
            ["Passes:Passed", "Fails:Failed", "Errors:Error", "HandlerCannotHideAFailure:Failed", "SeesTheErrNumber:Passed"],
            result.Results.Select(r => r.Name + ":" + r.Outcome));
        Assert.Equal(["hello"], result.Results[0].Output);
        Assert.Equal("Assert.AreEqual failed. Expected:<3>. Actual:<2>. sum", result.Results[1].Message);
        Assert.Equal(["before"], result.Results[1].Output);
        Assert.Equal("Run-time error 11: Division by zero", result.Results[2].Message);
        Assert.Equal("Assert.Fail failed. must surface", result.Results[3].Message);
        Assert.Empty(result.Results[3].Output);
        Assert.StartsWith("5 tests: 2 passed, 2 failed, 1 error.", result.Summary, StringComparison.Ordinal);

        var lines = result.ToText().Split(Environment.NewLine);
        Assert.Equal("passed  Tests.Passes", lines[0]);
        Assert.Equal("FAILED  Tests.Fails", lines[1]);
        Assert.Equal("        Assert.AreEqual failed. Expected:<3>. Actual:<2>. sum", lines[2]);
        Assert.Equal("        | before", lines[3]);
        Assert.Equal("ERROR   Tests.Errors", lines[4]);
        Assert.Equal("        Run-time error 11: Division by zero", lines[5]);

        var report = XDocument.Parse(result.ToJUnitXml());
        Assert.Equal("2", report.Root!.Attribute("failures")!.Value);
        Assert.Equal("1", report.Root.Attribute("errors")!.Value);
        var failing = report.Descendants("testcase").Single(t => t.Attribute("name")!.Value == "Fails");
        Assert.Equal("Assert.AreEqual failed. Expected:<3>. Actual:<2>. sum", failing.Element("failure")!.Attribute("message")!.Value);
        Assert.Equal("before", failing.Element("system-out")!.Value);
        Assert.NotNull(report.Descendants("testcase").Single(t => t.Attribute("name")!.Value == "Errors").Element("error"));
    }

    /// <summary>
    /// A class Property Let with an index parameter reached late-bound: the dispatch table's Put
    /// case must hand a ByRef parameter a temporary, as LateArguments does for Invoke; it passed a
    /// value and Roslyn rejected the generated C# (VBA0003, CS1620; VBA-Dictionary did not build).
    /// </summary>
    [Fact]
    public void Build_ClassPropertyLetWithIndex_ReachesLateBoundPut()
    {
        var projectDir = Path.Combine(workDir, "Store.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Store.cls"), string.Join("\r\n", [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"Store\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
            "Private items As Collection",
            "Private Sub Class_Initialize()",
            "    Set items = New Collection",
            "End Sub",
            "Public Property Get Item(key As Variant) As Variant",
            "    Item = items(key)",
            "End Property",
            "Public Property Let Item(key As Variant, value As Variant)",
            "    items.Add value, key",
            "End Property",
            "",
        ]));
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Public Sub Main()",
            "    Dim o As Object",
            "    Set o = New Store",
            "    o.Item(\"k\") = 42",
            "    Debug.Print o.Item(\"k\")",
            "End Sub",
            "",
        ]));
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal([" 42 "], RunCapturing(host, projectDir, "Main.Main"));
    }

    /// <summary>
    /// A PublicNotCreatable class with a predeclared instance carries the same attributes as a
    /// sheet module (VB_PredeclaredId and VB_Exposed both True), so the workbook beside the folder
    /// decides (ARCHITECTURE.md D19): with a workbook whose CodeNames are Sheet1 and ThisWorkbook,
    /// Helper.cls is a class and its Event compiles; without one, the attributes make it a document
    /// module and the Event is rejected, as before.
    /// </summary>
    [Fact]
    public void Build_WorkbookCodeNames_DecideWhichClassesAreDocumentModules()
    {
        var folder = Path.Combine(workDir, "Book");
        Directory.CreateDirectory(folder);
        var projectDir = Path.Combine(folder, "Book.vbang");
        Directory.CreateDirectory(projectDir);
        var header = string.Join("\r\n", [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"{0}\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Option Explicit",
            "",
        ]);
        File.WriteAllText(Path.Combine(projectDir, "Sheet1.cls"), header.Replace("{0}", "Sheet1", StringComparison.Ordinal) + "Private Sub Worksheet_Change(ByVal Target As Object)\r\nEnd Sub\r\n");
        File.WriteAllText(Path.Combine(projectDir, "Helper.cls"), header.Replace("{0}", "Helper", StringComparison.Ordinal) + "Public Event Changed()\r\nPublic Sub Fire()\r\n    RaiseEvent Changed\r\nEnd Sub\r\n");
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), "Attribute VB_Name = \"Main\"\r\nPublic Sub Main()\r\n    Helper.Fire\r\n    Debug.Print \"fired\"\r\nEnd Sub\r\n");

        var withoutWorkbook = ProjectCompiler.Build(projectDir);
        Assert.False(withoutWorkbook.Success);
        Assert.Contains(withoutWorkbook.Diagnostics, d => d.Message.Contains("Event declarations are only valid in a class module", StringComparison.Ordinal));

        File.Copy(Path.Combine(TestPaths.RepositoryRoot, "samples", "Sheets", "Sheets.xlsx"), Path.Combine(folder, "Book.xlsx"));
        var withWorkbook = ProjectCompiler.Build(projectDir);
        Assert.True(withWorkbook.Success, string.Join(Environment.NewLine, withWorkbook.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal(["fired"], RunCapturing(host, projectDir, "Main.Main"));
    }

    [Fact]
    public void Build_MembersNamedLikeTheirModule_Compile()
    {
        var projectDir = Path.Combine(workDir, "Names.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), "Attribute VB_Name = \"Main\"\r\nPublic Sub Main()\r\n    Debug.Print \"main\"\r\nEnd Sub\r\n");
        File.WriteAllText(Path.Combine(projectDir, "Shapes.bas"), "Attribute VB_Name = \"Shapes\"\r\nPublic Type Shapes\r\n    Count As Long\r\nEnd Type\r\nPublic Sub Draw()\r\n    Dim s As Shapes\r\n    s.Count = 2\r\n    Main.Main\r\n    Debug.Print s.Count; Colors.Colors\r\nEnd Sub\r\n");
        File.WriteAllText(Path.Combine(projectDir, "Colors.bas"), "Attribute VB_Name = \"Colors\"\r\nPublic Colors As Long\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var output = RunCapturing(host, projectDir, "Shapes.Draw");

        Assert.Equal(["main", " 2  0 "], output);
        Assert.Equal(["main"], RunCapturing(host, projectDir, "Main.Main"));
    }

    [Fact]
    public void Build_FileStatements_ReadAndWriteThroughTheRuntime()
    {
        var projectDir = Path.Combine(workDir, "Files.vbang");
        Directory.CreateDirectory(projectDir);
        var path = Path.Combine(workDir, "files.txt");
        File.WriteAllText(
            Path.Combine(projectDir, "Files.bas"),
            "Attribute VB_Name = \"Files\"\r\n"
            + "Public Type TRec\r\n    Id As Long\r\n    Label As String * 3\r\nEnd Type\r\n"
            + "Public Sub Main()\r\n"
            + "    Dim p As String, f As Integer, s As String, n As Long, r As TRec, r2 As TRec\r\n"
            + "    p = \"" + path + "\"\r\n"
            + "    f = FreeFile\r\n"
            + "    Open p For Output As #f\r\n"
            + "    Width #f, 80\r\n"
            + "    Print #f, \"hello\"; 1\r\n"
            + "    Write #f, \"a\", 2\r\n"
            + "    Close #f\r\n"
            + "    Open p For Input As #f\r\n"
            + "    Line Input #f, s\r\n"
            + "    Debug.Print s\r\n"
            + "    Input #f, s, n\r\n"
            + "    Debug.Print s; n; EOF(f)\r\n"
            + "    Close #f\r\n"
            + "    Kill p\r\n"
            + "    Open p For Binary As #f\r\n"
            + "    r.Id = 7: r.Label = \"ab\"\r\n"
            + "    Put #f, 1, r\r\n"
            + "    Lock #f: Unlock #f\r\n"
            + "    Seek #f, 1\r\n"
            + "    Get #f, , r2\r\n"
            + "    Debug.Print r2.Id; r2.Label; Len(r2); LOF(f)\r\n"
            + "    Close\r\n"
            + "    Name p As p & \".bak\"\r\n"
            + "    Debug.Print Dir(p) = \"\"; Dir(p & \".bak\")\r\n"
            + "    Kill p & \".bak\"\r\n"
            + "    Reset\r\n"
            + "End Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        var output = RunCapturing(host, projectDir, "Files.Main");

        Assert.Equal(["hello 1 ", "a 2 True", " 7 ab  7  7 ", "Truefiles.txt.bak"], output);
    }

    [Fact]
    public void Build_InvalidManifest_ReportsVba0022()
    {
        var projectDir = Path.Combine(workDir, "Broken.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), "Public Sub Main()\r\nEnd Sub\r\n");
        var manifest = Path.Combine(projectDir, "vbang.json");
        File.WriteAllText(manifest, "{ \"references\": [ { \"name\": \"vbang\" }");

        var result = ProjectCompiler.Build(projectDir);

        Assert.False(result.Success);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(DiagnosticIds.InvalidManifest, diagnostic.Id);
        Assert.Equal(manifest, diagnostic.FilePath);

        File.WriteAllText(manifest, "{ \"references\": [ { \"guid\": \"{0}\" } ] }");
        Assert.Equal(DiagnosticIds.InvalidManifest, Assert.Single(ProjectCompiler.Build(projectDir).Diagnostics).Id);
    }

    [Fact]
    public void Build_UnresolvedReference_ReportsVba0023()
    {
        var projectDir = Path.Combine(workDir, "Refs.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), "Public Sub Main()\r\nEnd Sub\r\n");
        var manifest = Path.Combine(projectDir, "vbang.json");
        File.WriteAllText(manifest, "{ \"references\": [ { \"name\": \"Nope\", \"guid\": \"{00000000-0000-0000-0000-000000000001}\", \"version\": \"1.0\" } ] }");

        var result = ProjectCompiler.Build(projectDir);

        Assert.False(result.Success);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(DiagnosticIds.ReferenceNotFound, diagnostic.Id);
        Assert.Equal(manifest, diagnostic.FilePath);
        Assert.Contains("Nope", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_WithoutTheVbangReference_AssertMeansWhatVbaMeans()
    {
        var projectDir = Path.Combine(workDir, "Plain.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Tests.bas"), "Attribute VB_Name = \"Tests\"\r\nOption Explicit\r\n'@Test\r\nPublic Sub T()\r\n    Assert.AreEqual 1, 1\r\nEnd Sub\r\n");

        var result = ProjectCompiler.Build(projectDir);

        Assert.False(result.Success);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(DiagnosticIds.VariableNotDefined, diagnostic.Id);
        Assert.Equal(5, diagnostic.Line);

        WriteVbangManifest(projectDir);
        Assert.True(ProjectCompiler.Build(projectDir).Success);
    }

    [Fact]
    public void Build_UnsupportedStatement_ReportsVba0002InCanonicalFormat()
    {
        var projectDir = Path.Combine(workDir, "Broken.vbang");
        Directory.CreateDirectory(projectDir);
        var source = Path.Combine(projectDir, "Module1.bas");
        File.WriteAllText(source, "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n\r\nPublic Declare PtrSafe Sub F Lib \"k\" (ParamArray a() As Variant)\r\n");

        var result = ProjectCompiler.Build(projectDir);

        Assert.False(result.Success);
        var diagnostic = result.Diagnostics.First(d => d.Id == DiagnosticIds.NotSupported);
        Assert.Equal(source + "(4,39): error VBA0002: A ParamArray parameter of a Declare needs its own marshaling.", diagnostic.ToString());
    }

    [Fact]
    public void Build_SyntaxError_ReportsVba0001()
    {
        var projectDir = Path.Combine(workDir, "Broken.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Module1.bas"), "Public Sub Main()\r\n    Debug.Print (1\r\nEnd Sub\r\n");

        var result = ProjectCompiler.Build(projectDir);

        Assert.False(result.Success);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(DiagnosticIds.SyntaxError, diagnostic.Id);
        Assert.Equal(2, diagnostic.Line);
        Assert.Equal(19, diagnostic.Column);
    }

    /// <summary>
    /// End and an unload reset the project as VBA does at End and when its workbook closes
    /// (docs/vba-quirks.md): module-level variables and Static locals return to their initial
    /// values, and no Class_Terminate runs, neither for the module's objects nor for the locals End
    /// unwinds; nothing after End runs; a normal release afterwards runs Class_Terminate again.
    /// </summary>
    [Fact]
    public void End_AndUnload_ResetTheProjectWithoutClassTerminate()
    {
        var projectDir = Path.Combine(workDir, "Reset.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(
            Path.Combine(projectDir, "Probe.cls"),
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
            + "Attribute VB_Name = \"Probe\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            + "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
            + "Option Explicit\r\n"
            + "Public Tag As String\r\n"
            + "Private Sub Class_Terminate()\r\n    Debug.Print \"T\" & Tag\r\nEnd Sub\r\n");
        File.WriteAllText(
            Path.Combine(projectDir, "Main.bas"),
            "Attribute VB_Name = \"Main\"\r\n"
            + "Option Explicit\r\n"
            + "Public first As Probe\r\n"
            + "Public text As String\r\n"
            + "Public Sub SetUp()\r\n    Set first = New Probe\r\n    first.Tag = \"first\"\r\n    text = \"kept\"\r\n    Counter\r\nEnd Sub\r\n"
            + "Public Function Counter() As Long\r\n    Static n As Long\r\n    n = n + 1\r\n    Counter = n\r\nEnd Function\r\n"
            + "Public Sub Halt()\r\n    Dim held As Probe\r\n    Set held = New Probe\r\n    held.Tag = \"held\"\r\n    Inner held\r\n    Debug.Print \"after\"\r\nEnd Sub\r\n"
            + "Private Sub Inner(ByVal p As Probe)\r\n    Dim inside As Probe\r\n    Set inside = New Probe\r\n    inside.Tag = \"inside\"\r\n    End\r\nEnd Sub\r\n"
            + "Public Sub Show()\r\n    Debug.Print CStr(first Is Nothing) & \",\" & Len(text) & \",\" & Counter()\r\nEnd Sub\r\n"
            + "Public Sub Control()\r\n    Dim c As Probe\r\n    Set c = New Probe\r\n    c.Tag = \"control\"\r\nEnd Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Empty(RunCapturing(host, projectDir, "Main.SetUp"));
        Assert.Empty(RunCapturing(host, projectDir, "Main.Halt"));
        Assert.Equal(["True,0,1"], RunCapturing(host, projectDir, "Main.Show"));
        Assert.Equal(["Tcontrol"], RunCapturing(host, projectDir, "Main.Control"));
        Assert.Empty(RunCapturing(host, projectDir, "Main.SetUp"));

        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = capture;
        try
        {
            host.Unload(projectDir);
        }
        finally
        {
            Host.Current = previous;
        }

        Assert.Empty(capture.Lines);
    }

    /// <summary>
    /// A Class_Terminate that stores Me brings the object back, as VBA-TDD's SpecDefinition hands itself to its suite
    /// (Lifetime cases): its variables stay, and when that reference goes they go too, with no second Class_Terminate.
    /// vba-ng released them right after Class_Terminate, so a spec read back later said 91.
    /// </summary>
    [Fact]
    public void ClassTerminate_StoringMe_KeepsTheObjectUntilThatReferenceGoes()
    {
        var projectDir = Path.Combine(workDir, "Revive.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(
            Path.Combine(projectDir, "Phoenix.cls"),
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
            + "Attribute VB_Name = \"Phoenix\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            + "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
            + "Option Explicit\r\n"
            + "Public Tag As String\r\n"
            + "Public Child As Collection\r\n"
            + "Private Sub Class_Initialize()\r\n    Set Child = New Collection\r\nEnd Sub\r\n"
            + "Private Sub Class_Terminate()\r\n    Debug.Print \"T\" & Tag\r\n    If Revive Then Keep.Add Me\r\nEnd Sub\r\n");
        File.WriteAllText(
            Path.Combine(projectDir, "Main.bas"),
            "Attribute VB_Name = \"Main\"\r\n"
            + "Option Explicit\r\n"
            + "Public Keep As New Collection\r\n"
            + "Public Revive As Boolean\r\n"
            + "Public Sub Run()\r\n"
            + "    Revive = True\r\n"
            + "    With New Phoenix\r\n        .Tag = \"b\"\r\n        .Child.Add New Phoenix\r\n        .Child(1).Tag = \"c\"\r\n    End With\r\n"
            + "    Revive = False\r\n"
            + "    Debug.Print Keep(1).Tag & \",\" & Keep(1).Child.Count\r\n"
            + "    Keep.Remove 1\r\n"
            + "    Debug.Print \"removed\"\r\n"
            + "End Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal(["Tb", "b,1", "Tc", "removed"], RunCapturing(host, projectDir, "Main.Run"));
    }

    /// <summary>
    /// A class instance keeps its variables in a block of its own (ROADMAP.md M7 E4): VarPtr of a
    /// class variable is the address of its storage, which CopyMemory reads and writes, and the
    /// address stays the same from call to call.
    /// </summary>
    [Fact]
    public void Build_ClassInstance_ShowsItsVariablesToTheDebugger()
    {
        var projectDir = Path.Combine(workDir, "View.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(
            Path.Combine(projectDir, "Tally.cls"),
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
            + "Attribute VB_Name = \"Tally\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            + "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
            + "Option Explicit\r\n"
            + "Private count As Long\r\n"
            + "Private label As String\r\n"
            + "Public Other As Tally\r\n"
            + "Public Sub Bump()\r\n    count = count + 1\r\n    label = \"bumped\"\r\n    Set Other = Me\r\nEnd Sub\r\n");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));
        var type = host.Load(projectDir).Assembly.GetType("Tally")!;
        var tally = Activator.CreateInstance(type)!;
        type.GetMethod("Bump")!.Invoke(tally, null);

        // The class names its view the way the debugger finds it (ARCHITECTURE.md section 9): the variables by value, and none of the runtime's own members.
        var proxyName = type.GetCustomAttributes(typeof(System.Diagnostics.DebuggerTypeProxyAttribute), inherit: false).Cast<System.Diagnostics.DebuggerTypeProxyAttribute>().Single().ProxyTypeName;
        var proxyType = type.GetNestedTypes(System.Reflection.BindingFlags.NonPublic).Single(t => proxyName.StartsWith(t.FullName!, StringComparison.Ordinal));
        var proxy = Activator.CreateInstance(proxyType, tally)!;
        var shown = proxyType.GetProperties().ToDictionary(p => p.Name, p => p.GetValue(proxy));
        Assert.Equal(["count", "label", "Other"], shown.Keys);
        Assert.Equal(1, shown["count"]);
        Assert.Equal("bumped", shown["label"]!.ToString());
        Assert.Same(tally, shown["Other"]!.GetType().GetProperty("Target")!.GetValue(shown["Other"]));
    }

    [Fact]
    public void Build_ClassVariable_HasAStableAddressInTheInstanceBlock()
    {
        var projectDir = Path.Combine(workDir, "Block.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(
            Path.Combine(projectDir, "Cell.cls"),
            "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
            + "Attribute VB_Name = \"Cell\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            + "Attribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n"
            + "Option Explicit\r\n"
            + "Private n As Long\r\n"
            + "Public Sub SetN(ByVal value As Long)\r\n    n = value\r\nEnd Sub\r\n"
            + "Public Function Address() As LongPtr\r\n    Address = VarPtr(n)\r\nEnd Function\r\n"
            + "Public Function ReadN() As Long\r\n    Dim r As Long\r\n    MoveMemory r, ByVal VarPtr(n), 4\r\n    ReadN = r\r\nEnd Function\r\n"
            + "Public Sub WriteN(ByVal value As Long)\r\n    MoveMemory ByVal VarPtr(n), value, 4\r\nEnd Sub\r\n"
            + "Public Function Value() As Long\r\n    Value = n\r\nEnd Function\r\n");
        File.WriteAllText(
            Path.Combine(projectDir, "Main.bas"),
            "Attribute VB_Name = \"Main\"\r\n"
            + "Option Explicit\r\n"
            + "Public Declare PtrSafe Sub MoveMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)\r\n"
            + "Public Sub Main()\r\n"
            + "    Dim c As Cell, first As LongPtr\r\n"
            + "    Set c = New Cell\r\n"
            + "    c.SetN 42\r\n"
            + "    first = c.Address()\r\n"
            + "    Debug.Print CStr(first <> 0) & \",\" & c.ReadN()\r\n"
            + "    c.WriteN 7\r\n"
            + "    Debug.Print c.Value() & \",\" & CStr(c.Address() = first)\r\n"
            + "End Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal(["True,42", "7,True"], RunCapturing(host, projectDir, "Main.Main"));
    }

    /// <summary>
    /// A Print that leaves its line open (MS-VBAL 5.4.5.6: the list ends in ";") still reaches the
    /// output when the run ends, as the Immediate window shows it; the output of a host command
    /// (vbang run, vbang test) lost it.
    /// </summary>
    [Fact]
    public void Run_PrintLeavingItsLineOpen_ReachesTheOutputWhenTheRunEnds()
    {
        var projectDir = Path.Combine(workDir, "OpenLine.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), "Attribute VB_Name = \"Main\"\r\nPublic Sub Go()\r\n    Debug.Print \"a\";\r\n    Debug.Print \"b\";\r\nEnd Sub\r\n");
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal(["ab"], RunCapturing(host, projectDir, "Main.Go"));
    }

    /// <summary>
    /// A ByRef argument passed through a temporary (a Long or an array to a ByRef Variant) goes
    /// back when its call returns, before the statement stores the call's result (MS-VBAL
    /// 5.6.13.1; Procedures golden), so <c>n = Bumped(n)</c> keeps the result. VBA-Better-Array's
    /// <c>Current = MultiToJagged(Current)</c> depends on it.
    /// </summary>
    [Fact]
    public void Run_ByRefArgumentOfTheCallAssigned_GoesBackBeforeTheResultIsStored()
    {
        var projectDir = Path.Combine(workDir, "CopyBackOrder.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Public Function Bumped(v As Variant) As Long",
            "    v = v + 10",
            "    Bumped = v * 2",
            "End Function",
            "Public Function Halved(v As Variant) As Variant()",
            "    Dim r() As Variant",
            "    ReDim r(0 To 0)",
            "    r(0) = UBound(v)",
            "    Halved = r",
            "End Function",
            "Public Sub Go()",
            "    Dim n As Long",
            "    Dim a() As Variant",
            "    n = 1",
            "    n = Bumped(n)",
            "    ReDim a(0 To 2)",
            "    a = Halved(a)",
            "    Debug.Print n; UBound(a); a(0)",
            "End Sub",
            string.Empty,
        ]));
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal([" 22  0  2 "], RunCapturing(host, projectDir, "Main.Go"));
    }

    /// <summary>
    /// A procedure may walk its ParamArray with For Each more than once: the array is the caller's
    /// temporary, which a loop in the callee must not take over and destroy when it ends
    /// (Procedures golden). VBA-Better-Array's ArrayGenerator walks its arguments twice.
    /// </summary>
    [Fact]
    public void Run_ParamArrayWalkedTwice_KeepsItsElements()
    {
        var projectDir = Path.Combine(workDir, "Walked.vbang");
        Directory.CreateDirectory(projectDir);
        File.WriteAllText(Path.Combine(projectDir, "Main.bas"), string.Join("\r\n", [
            "Attribute VB_Name = \"Main\"",
            "Public Function TwiceOver(ParamArray items() As Variant) As String",
            "    Dim item As Variant",
            "    For Each item In items",
            "        TwiceOver = TwiceOver & item",
            "    Next",
            "    For Each item In items",
            "        TwiceOver = TwiceOver & item",
            "    Next",
            "End Function",
            "Public Function Widths(ParamArray items() As Variant) As String",
            "    Dim item As Variant",
            "    For Each item In items",
            "        Widths = Widths & UBound(item)",
            "    Next",
            "    For Each item In items",
            "        Widths = Widths & UBound(item)",
            "    Next",
            "End Function",
            "Public Sub Go()",
            "    Debug.Print TwiceOver(1, \"a\", 2)",
            "    Debug.Print Widths(Array(1, 2), Array(3))",
            "End Sub",
            string.Empty,
        ]));
        Assert.True(ProjectCompiler.Build(projectDir).Success);
        using var host = new ProjectHost(Path.Combine(workDir, "shadow"));

        Assert.Equal(["1a21a2", "1010"], RunCapturing(host, projectDir, "Main.Go"));
    }

    private static TestRunResult RunTests(ProjectHost host, string projectDir)
    {
        var previous = Host.Current;
        Host.Current = new CapturingHostServices();
        try
        {
            return TestRunner.Run(host.Load(projectDir));
        }
        finally
        {
            Host.Current = previous;
        }
    }

    private static IReadOnlyList<string> RunCapturing(ProjectHost host, string projectDir, string procedure)
    {
        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = capture;
        try
        {
            host.Load(projectDir).Run(procedure);
        }
        finally
        {
            Host.Current = previous;
        }

        return capture.Lines;
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

    /// <summary>A manifest referencing the vbang library, which the Assert module needs (ARCHITECTURE.md section 8).</summary>
    private static void WriteVbangManifest(string projectDir) =>
        File.WriteAllText(Path.Combine(projectDir, "vbang.json"), "{ \"references\": [ { \"name\": \"vbang\" } ] }");

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
