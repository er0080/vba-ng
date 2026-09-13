using VbaNg.Compiler.Binding;
using VbaNg.Runtime.TypeLibraries;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// A member the type library does not list binds late on an interface without
/// TYPEFLAG_FNONEXTENSIBLE and is a compile error on one with it (docs/vba-quirks.md, verified
/// in Excel on 2026-09-09: Range, Application, and MSForms.Control take unknown members at run
/// time, Worksheet does not). A synthetic library with one type of each kind stands in for them.
/// </summary>
public sealed class ExtensibleInterfaceTests
{
    private static readonly ComLibrary Library = new("Widgets", new Guid("11111111-2222-3333-4444-555555555555"), 1, 0,
    [
        new ComType("Loose", ComTypeKind.CoClass, new Guid("11111111-2222-3333-4444-000000000001"), [], [new ComImplemented("Widgets.ILoose", IsDefault: true, IsSource: false, IsRestricted: false)], null, false, false, false, false),
        new ComType("ILoose", ComTypeKind.Dispatch, new Guid("11111111-2222-3333-4444-000000000002"), [Name(1)], [new ComImplemented("stdole.IDispatch", IsDefault: false, IsSource: false, IsRestricted: false)], null, false, false, false, IsDual: true, IsExtensible: true),
        new ComType("Strict", ComTypeKind.CoClass, new Guid("11111111-2222-3333-4444-000000000003"), [], [new ComImplemented("Widgets.IStrict", IsDefault: true, IsSource: false, IsRestricted: false)], null, false, false, false, false),
        new ComType("IStrict", ComTypeKind.Dispatch, new Guid("11111111-2222-3333-4444-000000000004"), [Name(1)], [new ComImplemented("stdole.IDispatch", IsDefault: false, IsSource: false, IsRestricted: false)], null, false, false, false, IsDual: true, IsExtensible: false),
    ]);

    private static ComMember Name(int dispId) =>
        new("Name", dispId, ComMemberKind.PropertyGet, new ComTypeRef(8, null, IsArray: false, IsPointer: false), [], IsHidden: false, IsRestricted: false, Value: null);

    [Fact]
    public void UnknownMember_OnAnExtensibleInterface_BindsLate()
    {
        var diagnostics = new List<Diagnostic>();

        var generated = ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Module1.bas", "Sub M(w As Widgets.Loose)\r\n    Debug.Print w.Name\r\n    w.Caption = \"x\"\r\n    Debug.Print w.Caption\r\nEnd Sub\r\n")], diagnostics, libraries: [Library]);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
        var text = Assert.Single(generated!).Text;
        Assert.Contains("global::VbaNg.Runtime.EarlyBound.Invoke(w.Target, 1, global::VbaNg.Runtime.InvokeKind.PropertyGet, [])", text, StringComparison.Ordinal);
        Assert.Contains("global::VbaNg.Runtime.LateBound.Let(", text, StringComparison.Ordinal);
        Assert.Contains("\"Caption\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void UnknownMember_OnANonExtensibleInterface_IsReported()
    {
        var diagnostics = new List<Diagnostic>();

        ProjectCompiler.Generate("Test", [new SourceFile(@"C:\test\Module1.bas", "Sub M(w As Widgets.Strict)\r\n    Debug.Print w.Name\r\n    Debug.Print w.Caption\r\nEnd Sub\r\n")], diagnostics, libraries: [Library]);

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(DiagnosticIds.VariableNotDefined, diagnostic.Id);
        Assert.Equal("Method or data member not found: 'Caption'.", diagnostic.Message);
        Assert.Equal(3, diagnostic.Line);
    }
}
