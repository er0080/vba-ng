using VbaNg.Compiler.Binding;
using VbaNg.Runtime.TypeLibraries;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Early binding against a type library (ARCHITECTURE.md section 6): the Scripting runtime's
/// model, read once from the registry and committed as a fixture so these tests need no COM at
/// all. Checks what the binder resolves and the C# the emitter produces for it.
/// </summary>
public sealed class ComBindingTests
{
    private const string FilePath = @"C:\test\Module1.bas";

    private static readonly ComLibrary Scripting = ComLibrary.Load(Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures", "TypeLibs", "Scripting.json"));

    [Fact]
    public void Fixture_IsTheScriptingLibrary()
    {
        Assert.Equal("Scripting", Scripting.Name);
        Assert.NotNull(Scripting.FindType("Dictionary"));
    }

    [Fact]
    public void Bind_DeclaredComType_CallsMembersByDispid()
    {
        var text = Generate(
            "Sub M()\r\nDim d As Scripting.Dictionary\r\nSet d = New Scripting.Dictionary\r\nd.Add \"a\", 1\r\nDim n As Long\r\nn = d.Count\r\nDebug.Print d.Item(\"a\"); d(\"a\")\r\nd.Item(\"a\") = 2\r\nd(\"b\") = 3\r\nEnd Sub\r\n");

        // A typed object variable is an object slot, the interface pointer itself (ROADMAP.md M7 E3), read through Target.
        Assert.Contains("global::VbaNg.Runtime.ObjectSlot<global::VbaNg.Runtime.IDispatchObject> d = default;", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.Com.CreateInstance(\"ee09b103-97e0-11cf-978f-00a02463e06f\")", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(d.Target, 1, global::VbaNg.Runtime.InvokeKind.Method, [", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.Coerce.ToInt32(global::VbaNg.Runtime.EarlyBound.Invoke(d.Target, 2, global::VbaNg.Runtime.InvokeKind.PropertyGet, []))", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(d.Target, 0, global::VbaNg.Runtime.InvokeKind.PropertyGet, [", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Put(d.Target, 0, [", text, StringComparison.Ordinal);
        Assert.DoesNotContain("LateBound", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Bind_UnqualifiedTypeNameAndAsNew_ResolveThroughTheLibrary()
    {
        var text = Generate("Sub M()\r\nDim d As New Dictionary\r\nd.Add \"a\", 1\r\nEnd Sub\r\n");

        Assert.Contains("ObjectRefs.AutoNew(ref d, () => global::VbaNg.Runtime.Com.CreateInstance(\"ee09b103-97e0-11cf-978f-00a02463e06f\"))", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Bind_ArgumentsAreLetCoercedToTheDeclaredParameterTypes()
    {
        // OpenTextFile(FileName As String, [IOMode As IOMode = ForReading], [Create As Boolean], [Format As Tristate]).
        var text = Generate("Sub M()\r\nDim fso As Scripting.FileSystemObject, ts As Scripting.TextStream\r\nSet fso = New FileSystemObject\r\nSet ts = fso.OpenTextFile(5, Create:=1)\r\nts.Close\r\nEnd Sub\r\n");

        Assert.Contains("global::VbaNg.Runtime.Variant.ViewString(global::VbaNg.Runtime.Coerce.ToText(", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.Variant.Missing, global::VbaNg.Runtime.Variant.FromBoolean(global::VbaNg.Runtime.Coerce.ToBoolean(", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.Coerce.ToObject<global::VbaNg.Runtime.IDispatchObject>(global::VbaNg.Runtime.EarlyBound.Invoke(fso.Target, ", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Bind_LibraryConstantsFoldToTheirValues()
    {
        var text = Generate("Sub M()\r\nDim m As Long\r\nm = ForAppending\r\nm = IOMode.ForReading\r\nm = Scripting.ForWriting\r\nDim t As IOMode\r\nt = 1\r\nEnd Sub\r\n");

        Assert.Contains("m = 8;", text, StringComparison.Ordinal);
        Assert.Contains("m = 1;", text, StringComparison.Ordinal);
        Assert.Contains("m = 2;", text, StringComparison.Ordinal);
        Assert.Contains("int t = 0;", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(DiagnosticIds.VariableNotDefined, 3, "Sub M()\r\nDim d As Dictionary\r\nd.Nope\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ArgumentNotOptional, 3, "Sub M()\r\nDim d As Dictionary\r\nd.Add\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.WrongNumberOfArguments, 3, "Sub M()\r\nDim d As Dictionary\r\nd.Add 1, 2, 3\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.NamedArgumentNotFound, 3, "Sub M()\r\nDim d As Dictionary\r\nd.Add Nope:=1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeNotDefined, 2, "Sub M()\r\nDim r As Excel.Range\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 4, "Sub M()\r\nDim d As Dictionary\r\nDim x\r\nx = d.RemoveAll()\r\nEnd Sub\r\n")]
    public void Bind_ReportsTheDiagnostic(string id, int line, string source)
    {
        var diagnostics = new List<Diagnostic>();
        ProjectCompiler.Generate("Test", [new SourceFile(FilePath, source)], diagnostics, libraries: [Scripting]);

        Assert.Contains(diagnostics, d => d.Id == id && d.Line == line);
    }

    [Fact]
    public void Bind_WithoutTheReference_LibraryNamesAreUnknown()
    {
        var diagnostics = new List<Diagnostic>();
        ProjectCompiler.Generate("Test", [new SourceFile(FilePath, "Option Explicit\r\nSub M()\r\nDim d As Dictionary\r\nDim m As Long\r\nm = ForReading\r\nEnd Sub\r\n")], diagnostics);

        Assert.Contains(diagnostics, d => d.Id == DiagnosticIds.TypeNotDefined && d.Line == 3);
        Assert.Contains(diagnostics, d => d.Id == DiagnosticIds.VariableNotDefined && d.Line == 5);
    }

    private static string Generate(string source)
    {
        var diagnostics = new List<Diagnostic>();
        var generated = ProjectCompiler.Generate("Test", [new SourceFile(FilePath, source)], diagnostics, libraries: [Scripting]);
        Assert.Empty(diagnostics.Select(d => d.ToString()));
        return Assert.Single(generated!).Text;
    }
}
