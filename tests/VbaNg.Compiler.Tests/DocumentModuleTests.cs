using VbaNg.Compiler.Binding;
using VbaNg.Runtime.TypeLibraries;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Document modules (ARCHITECTURE.md section 3): class modules with a predeclared instance that
/// the host binds to a workbook object. Bound against a committed subset of the Excel type
/// library model (Global, Workbook, Worksheet, Range, and their event interfaces), so no Office
/// is needed here.
/// </summary>
public sealed class DocumentModuleTests
{
    private const string Header = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

    private static readonly ComLibrary Excel = ComLibrary.Load(Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures", "TypeLibs", "ExcelSubset.json"));

    private static string Sheet(string name, string body, bool predeclared = true) =>
        Header + $"Attribute VB_Name = \"{name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = {(predeclared ? "True" : "False")}\r\nAttribute VB_Exposed = True\r\nOption Explicit\r\n" + body;

    [Fact]
    public void Worksheet_Module_BindsMeAndItsOwnMembers()
    {
        var text = Generate(
            @"C:\test\Sheet1.cls",
            Sheet("Sheet1", "Private Sub Worksheet_Change(ByVal Target As Range)\r\n    Range(\"A1\").Value = Target.Address\r\n    Cells(1, 2).Value = Name\r\n    Me.Range(\"B2\") = Me.Name\r\nEnd Sub\r\n"));

        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaDocumentModuleAttribute(\"Worksheet\")]", text, StringComparison.Ordinal);
        Assert.Contains("public static global::VbaNg.Runtime.IDispatchObject? Me;", text, StringComparison.Ordinal);
        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"Worksheet\", \"Change\")]", text, StringComparison.Ordinal);
        Assert.Contains("internal static void Worksheet_Change(global::VbaNg.Runtime.IDispatchObject? Target)", text, StringComparison.Ordinal);
        // Range, Cells, and Name resolve on the sheet itself (dispids of _Worksheet), not on the application object.
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(Me, 197, global::VbaNg.Runtime.InvokeKind.PropertyGet, [", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(Me, 238, global::VbaNg.Runtime.InvokeKind.PropertyGet, [])", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(Me, 110, global::VbaNg.Runtime.InvokeKind.PropertyGet, [])", text, StringComparison.Ordinal);
        Assert.DoesNotContain("__app_Excel", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Workbook_Module_IsAWorkbookWithItsEvents()
    {
        var text = Generate(
            @"C:\test\ThisWorkbook.cls",
            Sheet("ThisWorkbook", "Private Sub Workbook_Open()\r\n    Worksheets(1).Range(\"A1\").Value = 1\r\nEnd Sub\r\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\r\n    Cancel = True\r\nEnd Sub\r\nPrivate Sub Helper_Format()\r\nEnd Sub\r\n"));

        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaDocumentModuleAttribute(\"Workbook\")]", text, StringComparison.Ordinal);
        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"Workbook\", \"Open\")]", text, StringComparison.Ordinal);
        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"Workbook\", \"BeforeClose\")]", text, StringComparison.Ordinal);
        Assert.Contains("internal static void Workbook_BeforeClose(ref bool Cancel)", text, StringComparison.Ordinal);
        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"Helper\", \"Format\")]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Control_Handlers_NameTheControl()
    {
        var text = Generate(
            @"C:\test\Sheet1.cls",
            Sheet("Sheet1", "Private Sub CommandButton1_Click()\r\n    Range(\"A1\").Value = \"clicked\"\r\nEnd Sub\r\nPrivate Sub Worksheet_NoSuchEvent()\r\nEnd Sub\r\n"));

        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"CommandButton1\", \"Click\")]", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"NoSuchEvent\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Manifest_DocumentsOverrideTheKind()
    {
        var manifest = new ProjectManifest([new ManifestReference("Excel")]) { Documents = new Dictionary<string, string> { ["Chart1"] = "Chart" } };
        var diagnostics = new List<Diagnostic>();

        var generated = ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Chart1.cls", Sheet("Chart1", "Private Sub Chart_Activate()\r\nEnd Sub\r\n"))], diagnostics, manifest: manifest, libraries: [Excel]);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
        var text = Assert.Single(generated!).Text;
        Assert.Contains("VbaDocumentModuleAttribute(\"Chart\")", text, StringComparison.Ordinal);
        Assert.Contains("VbaEventHandlerAttribute(\"Chart\", \"Activate\")", text, StringComparison.Ordinal);
        Assert.Contains("public static object? Me;", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// The manifest's documents map is the whole list rather than a set of exceptions: a
    /// predeclared and exposed class module that is not in it is a class, which the attributes
    /// alone cannot say (ARCHITECTURE.md D19). stdVBA's Test carries exactly those attributes, and
    /// taking it for a sheet module put a host object behind Me, hiding its own members.
    /// </summary>
    [Fact]
    public void Manifest_DocumentsAreTheWholeList_SoAnUnlistedPredeclaredClassIsAClass()
    {
        var manifest = new ProjectManifest([new ManifestReference("Excel")]) { Documents = new Dictionary<string, string> { ["ThisWorkbook"] = "Workbook" } };
        var diagnostics = new List<Diagnostic>();

        var generated = ProjectCompiler.Generate(
            "Test",
            [new SourceFile(@"C:\test\Helper.cls", Sheet("Helper", "Friend Sub Hidden()\r\nEnd Sub\r\nPublic Sub Use()\r\n    Me.Hidden\r\nEnd Sub\r\n"))],
            diagnostics,
            manifest: manifest,
            libraries: [Excel]);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
        var text = Assert.Single(generated!).Text;
        Assert.Contains("public sealed class Helper : global::VbaNg.Runtime.VbaClassObject", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaDocumentModuleAttribute", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// <c>ws.[Name] = value</c> assigns to the default member of what the sheet's Evaluate returns,
    /// late-bound, exactly as <c>[Name] = value</c> does through the Application (MS-VBAL 3.3.5.3).
    /// It was VBA0019 before, which stopped four of VBA-Web's sheet modules (ROADMAP.md WP6).
    /// </summary>
    [Fact]
    public void Worksheet_BracketedMember_IsAnAssignmentTargetThroughEvaluate()
    {
        var text = Generate(@"C:\test\Sheet1.cls", Sheet("Sheet1", "Public Sub Go()\r\n    Me.[Distance] = \"x\"\r\nEnd Sub\r\n"));

        Assert.Contains("LateBound.Let(", text, StringComparison.Ordinal);
        Assert.Contains("\"Evaluate\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void StandardModule_ResolvesRangeOnTheApplicationObject()
    {
        var text = Generate(@"C:\test\Module1.bas", "Attribute VB_Name = \"Module1\"\r\nSub M()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub\r\nSub Worksheet_Change()\r\nEnd Sub\r\n");

        Assert.Contains("__app_Excel", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaEventHandlerAttribute", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaDocumentModuleAttribute", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A .cls without VB_Exposed = True is a class of the project, not a module bound to a
    /// workbook object: it becomes a class with instance members, no Me field, and no document
    /// attribute (ARCHITECTURE.md section 3).
    /// </summary>
    [Fact]
    public void ClassModule_IsAClassNotADocumentModule()
    {
        var diagnostics = new List<Diagnostic>();

        var generated = ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Customer.cls", Sheet("Customer", "Public Name As String\r\n", predeclared: false))], diagnostics, libraries: [Excel]);

        Assert.Empty(diagnostics);
        var text = Assert.Single(generated!).Text;
        Assert.Contains("public sealed class Customer : global::VbaNg.Runtime.VbaClassObject", text, StringComparison.Ordinal);
        Assert.Contains("public global::VbaNg.Runtime.VbaString Name = global::VbaNg.Runtime.VbaString.Null;", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaDocumentModuleAttribute", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Me_InAStandardModule_IsAnError()
    {
        var diagnostics = new List<Diagnostic>();

        ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Module1.bas", "Sub M()\r\nDebug.Print Me.Name\r\nEnd Sub\r\n")], diagnostics, libraries: [Excel]);

        Assert.Contains(diagnostics, d => d.Id == DiagnosticIds.NotSupported && d.Line == 2);
    }

    /// <summary>
    /// WithEvents on a library type (MS-VBAL 5.2.3.1.4): the coclass's default source interface
    /// names the events, the handlers become the class's sink, and Set advises through the runtime
    /// with the handled events by dispid (ARCHITECTURE.md section 6, "Events"). A procedure whose
    /// prefix is the variable but whose suffix is no event of the source is an ordinary one.
    /// </summary>
    [Fact]
    public void ClassModule_WithEventsOnALibraryType_AdvisesTheSourceInterface()
    {
        var text = Generate(
            @"C:\test\Watcher.cls",
            Sheet("Watcher", "Private WithEvents ws As Worksheet\r\nPublic Sub Watch(target As Worksheet)\r\n    Set ws = target\r\nEnd Sub\r\nPrivate Sub ws_Change(ByVal Target As Range)\r\n    Debug.Print Target.Address\r\nEnd Sub\r\nPrivate Sub ws_Helper()\r\nEnd Sub\r\n", predeclared: false));

        Assert.Contains("public sealed class Watcher : global::VbaNg.Runtime.VbaClassObject, global::VbaNg.Runtime.IVbaEventSink", text, StringComparison.Ordinal);
        Assert.Contains("private static readonly global::VbaNg.Runtime.ComEventInterface __ws_Events = new(new global::System.Guid(\"00024411-0000-0000-c000-000000000046\"), \"DocEvents\", new (int, string)[] { (1545, \"Change\") });", text, StringComparison.Ordinal);
        Assert.Contains("private global::System.IDisposable? __ws_Advisory;", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.ObjectRefs.Subscribe(ref ws, target.Target, this, \"ws\", __ws_Events, ref __ws_Advisory)", text, StringComparison.Ordinal);
        Assert.Contains("case \"WS.CHANGE\":", text, StringComparison.Ordinal);
        Assert.DoesNotContain("WS.HELPER", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaEventHandlerAttribute", text, StringComparison.Ordinal);
    }

    /// <summary>A document module handles a WithEvents variable's events too (Private WithEvents App As Application); its static handlers need a nested object as the sink, and the host is not asked to connect them.</summary>
    [Fact]
    public void DocumentModule_WithEventsOnALibraryType_HasANestedSink()
    {
        var text = Generate(
            @"C:\test\ThisWorkbook.cls",
            Sheet("ThisWorkbook", "Private WithEvents Book As Workbook\r\nPrivate Sub Workbook_Open()\r\n    Set Book = Me\r\nEnd Sub\r\nPrivate Sub Book_BeforeClose(Cancel As Boolean)\r\n    Cancel = True\r\nEnd Sub\r\n"));

        Assert.Contains("[global::VbaNg.Runtime.Hosting.VbaEventHandlerAttribute(\"Workbook\", \"Open\")]", text, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaEventHandlerAttribute(\"Book\"", text, StringComparison.Ordinal);
        Assert.Contains("private static global::System.IDisposable? __Book_Advisory;", text, StringComparison.Ordinal);
        Assert.Contains("private static readonly __EventSink __events = new();", text, StringComparison.Ordinal);
        Assert.Contains("private sealed class __EventSink : global::VbaNg.Runtime.IVbaEventSink", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.ObjectRefs.Subscribe(ref Book, Me, __events, \"Book\", __Book_Events, ref __Book_Advisory)", text, StringComparison.Ordinal);
        Assert.Contains("case \"BOOK.BEFORECLOSE\":", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A workbook beside the folder decides (ARCHITECTURE.md D19): a sheet module it names is a
    /// document module without the attributes saying so; one it does not name compiles as a class
    /// with a warning (VBA0024), since the host could never bind it; and without a workbook the
    /// attributes decide as before.
    /// </summary>
    [Fact]
    public void WorkbookBesideTheFolder_DecidesTheDocumentModules()
    {
        var sheet = Sheet("Sheet1", "Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n");
        var withSheet = ProjectManifest.Empty with { WorkbookDocuments = new Dictionary<string, string> { ["ThisWorkbook"] = "Workbook", ["Sheet1"] = "Worksheet" } };
        var withoutSheet = ProjectManifest.Empty with { WorkbookDocuments = new Dictionary<string, string> { ["ThisWorkbook"] = "Workbook" } };

        var named = new List<Diagnostic>();
        var namedText = Assert.Single(ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Sheet1.cls", sheet)], named, manifest: withSheet, libraries: [Excel])!).Text;
        var unnamed = new List<Diagnostic>();
        var unnamedText = Assert.Single(ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Sheet1.cls", sheet)], unnamed, manifest: withoutSheet, libraries: [Excel])!).Text;

        Assert.Empty(named);
        Assert.Contains("VbaDocumentModuleAttribute(\"Worksheet\")", namedText, StringComparison.Ordinal);
        var warning = Assert.Single(unnamed);
        Assert.Equal(DiagnosticIds.UnboundDocumentModule, warning.Id);
        Assert.Equal(DiagnosticSeverity.Warning, warning.Severity);
        Assert.Contains("'Sheet1' has a document module's attributes", warning.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("VbaDocumentModuleAttribute", unnamedText, StringComparison.Ordinal);
        Assert.Contains("public sealed class Sheet1 : global::VbaNg.Runtime.VbaClassObject", unnamedText, StringComparison.Ordinal);
    }

    [Fact]
    public void WithEvents_OnATypeWithoutEvents_IsAnError()
    {
        var diagnostics = new List<Diagnostic>();

        ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Watcher.cls", Sheet("Watcher", "Private WithEvents r As Range\r\n", predeclared: false))], diagnostics, libraries: [Excel]);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(DiagnosticIds.TypeMismatch, diagnostic.Id);
        Assert.Equal("Object does not source automation events.", diagnostic.Message);
    }

    [Fact]
    public void WithEvents_InAStandardModule_IsAnError()
    {
        var diagnostics = new List<Diagnostic>();

        ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Module1.bas", "Private WithEvents ws As Worksheet\r\n")], diagnostics, libraries: [Excel]);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(DiagnosticIds.InvalidStatementPlacement, diagnostic.Id);
        Assert.Equal("Only valid in object module.", diagnostic.Message);
    }

    private static string Generate(string path, string source)
    {
        var diagnostics = new List<Diagnostic>();
        var generated = ProjectCompiler.Generate("Test", [new SourceFile(path, source)], diagnostics, libraries: [Excel]);
        Assert.Empty(diagnostics.Select(d => d.ToString()));
        return Assert.Single(generated!).Text;
    }
}
