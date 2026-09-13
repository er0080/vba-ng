using VbaNg.Compiler.Binding;
using VbaNg.Runtime.TypeLibraries;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Shapes the corpus spec suites use that the binder rejected when they were first imported
/// (ROADMAP.md WP3): each is valid VBA that compiles in Excel. Bound against the Excel subset
/// fixture so the sheet shapes need no Office.
/// </summary>
public sealed class CorpusShapesTests
{
    private const string ClassHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

    private static readonly ComLibrary Excel = ComLibrary.Load(Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures", "TypeLibs", "ExcelSubset.json"));

    private static string Class(string name, string body, bool document = false) =>
        ClassHeader + $"Attribute VB_Name = \"{name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = {(document ? "True" : "False")}\r\nAttribute VB_Exposed = {(document ? "True" : "False")}\r\nOption Explicit\r\n" + body;

    /// <summary>DisplayRunner.IdCol = 1 in VBA-Dictionary's specs: a Property Let of a standard module through the module's name.</summary>
    [Fact]
    public void PropertyLet_OnAStandardModuleThroughItsName_IsATarget()
    {
        var text = Generate(
            ("Runner.bas", "Attribute VB_Name = \"Runner\"\r\nPrivate pCol As Integer\r\nPublic Property Get Col() As Integer\r\n    Col = pCol\r\nEnd Property\r\nPublic Property Let Col(Value As Integer)\r\n    pCol = Value\r\nEnd Property\r\n"),
            ("Main.bas", "Attribute VB_Name = \"Main\"\r\nSub M()\r\n    Runner.Col = 1\r\n    Debug.Print Runner.Col\r\nEnd Sub\r\n"));

        Assert.Contains("global::Runner.Let_Col(ref __t", text["Main"], StringComparison.Ordinal);
        Assert.Contains("global::Runner.Get_Col()", text["Main"], StringComparison.Ordinal);
    }

    /// <summary>Specs as a statement inside module Specs, whose function Specs it calls (VBA-Dictionary's specs).</summary>
    [Fact]
    public void ProcedureNamedLikeItsModule_CalledAsAStatementInside_Binds()
    {
        var text = Generate(("Specs.bas", "Attribute VB_Name = \"Specs\"\r\nPublic Function Specs() As Long\r\n    Specs = 1\r\nEnd Function\r\nPublic Sub Run()\r\n    Specs\r\n    Call Specs\r\n    Debug.Print Specs()\r\nEnd Sub\r\n"));

        Assert.Contains("Specs", text["Specs"], StringComparison.Ordinal);
    }

    /// <summary>Credentials.Values("Google")("id") in VBA-Web's specs: arguments after a parameterless function or property index its result.</summary>
    [Fact]
    public void ArgumentsAfterAParameterlessProcedure_IndexItsResult()
    {
        var text = Generate(
            ("Credentials.bas", "Attribute VB_Name = \"Credentials\"\r\nPublic Property Get Values() As Object\r\n    Set Values = Nothing\r\nEnd Property\r\nPublic Function Items() As Object\r\n    Set Items = Nothing\r\nEnd Function\r\nPublic Function Names() As Variant\r\n    Names = Array(\"a\", \"b\")\r\nEnd Function\r\n"),
            ("Main.bas", "Attribute VB_Name = \"Main\"\r\nSub M()\r\n    Debug.Print Credentials.Values(\"Google\")(\"id\")\r\n    Debug.Print Values(\"Google\")\r\n    Debug.Print Items(1)\r\n    Debug.Print Credentials.Items(1)\r\n    Debug.Print Names(1)\r\nEnd Sub\r\n"));

        Assert.Contains("LateBound.Index(global::VbaNg.Runtime.LateBound.Index(global::VbaNg.Runtime.Variant.FromObject(global::Credentials.Get_Values()), [global::VbaNg.Runtime.Variant.ViewString(__str_0)]), [global::VbaNg.Runtime.Variant.ViewString(__str_1)])", text["Main"], StringComparison.Ordinal);
        Assert.Contains("LateBound.Index(global::VbaNg.Runtime.Variant.FromObject(global::Credentials.Items()), [", text["Main"], StringComparison.Ordinal);
    }

    /// <summary>Reporter.ConnectTo SpecRunner in VBA-Web's specs: a sheet module's name is its Worksheet, and reaches a ByRef Worksheet parameter as one.</summary>
    [Fact]
    public void DocumentModule_ReachesAByRefParameterOfItsType()
    {
        var manifest = ProjectManifest.Empty with { WorkbookDocuments = new Dictionary<string, string> { ["ThisWorkbook"] = "Workbook", ["SpecRunner"] = "Worksheet" } };
        var text = Generate(manifest,
            ("SpecRunner.cls", Class("SpecRunner", string.Empty, document: true)),
            ("Reporter.cls", Class("Reporter", "Private pSheet As Worksheet\r\nPublic Sub ConnectTo(Sheet As Worksheet)\r\n    Set pSheet = Sheet\r\nEnd Sub\r\n")),
            ("Main.bas", "Attribute VB_Name = \"Main\"\r\nSub M()\r\n    Dim r As New Reporter\r\n    r.ConnectTo SpecRunner\r\n    r.ConnectTo ActiveSheet\r\nEnd Sub\r\n"));

        Assert.Contains("ConnectTo(", text["Main"], StringComparison.Ordinal);
    }

    /// <summary>
    /// An object of one type reaches a ByRef parameter of another object type: a class instance to
    /// As Object, an Object to a class, a Worksheet to As Object; a Variant variable does not
    /// (Classes golden; verified in Excel on 2026-09-09).
    /// </summary>
    [Fact]
    public void ObjectArguments_ReachByRefParametersOfOtherObjectTypes()
    {
        var text = Generate(
            ("Store.cls", Class("Store", "Public Sub Bump()\r\nEnd Sub\r\n")),
            ("Main.bas", "Attribute VB_Name = \"Main\"\r\nSub TakesObject(x As Object)\r\nEnd Sub\r\nSub TakesStore(x As Store)\r\nEnd Sub\r\nSub M()\r\n    Dim s As New Store\r\n    Dim o As Object\r\n    Set o = s\r\n    TakesObject s\r\n    TakesStore o\r\n    TakesObject Nothing\r\n    TakesStore New Store\r\nEnd Sub\r\n"));

        Assert.Contains("TakesStore(", text["Main"], StringComparison.Ordinal);
    }

    [Fact]
    public void VariantVariable_ToAByRefObjectParameter_IsRejected()
    {
        var diagnostics = new List<Diagnostic>();

        ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Main.bas", "Attribute VB_Name = \"Main\"\r\nSub TakesObject(x As Object)\r\nEnd Sub\r\nSub M()\r\n    Dim v As Variant\r\n    TakesObject v\r\nEnd Sub\r\n")], diagnostics, libraries: [Excel]);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(DiagnosticIds.ByRefTypeMismatch, diagnostic.Id);
    }

    private static Dictionary<string, string> Generate(params (string File, string Source)[] modules) => Generate(ProjectManifest.Empty, modules);

    private static Dictionary<string, string> Generate(ProjectManifest manifest, params (string File, string Source)[] modules)
    {
        var diagnostics = new List<Diagnostic>();
        var generated = ProjectCompiler.Generate("Test", modules.Select(m => new SourceFile(@"C:\test\" + m.File, m.Source)).ToList(), diagnostics, manifest: manifest, libraries: [Excel]);
        Assert.Empty(diagnostics.Select(d => d.ToString()));
        return generated!.ToDictionary(g => g.Module, g => g.Text, StringComparer.OrdinalIgnoreCase);
    }
}
