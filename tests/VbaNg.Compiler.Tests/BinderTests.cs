using VbaNg.Compiler.Binding;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Binder diagnostics (docs/diagnostics.md): every id has a test that triggers it (CLAUDE.md R14),
/// plus checks of what the binder accepts. Snippets are whole modules; the module name comes
/// from the file name.
/// </summary>
public sealed class BinderTests
{
    private const string FilePath = @"C:\test\Module1.bas";

    [Theory]
    [InlineData(DiagnosticIds.VariableNotDefined, 3, "Option Explicit\r\nSub M()\r\nx = 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.AmbiguousName, 3, "Sub M()\r\nDim x As Long\r\nDim x As Long\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.AmbiguousName, 2, "Public x As Long\r\nSub x()\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ProcedureNotDefined, 2, "Sub M()\r\nFoo 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.LabelNotDefined, 2, "Sub M()\r\nGoTo Done\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.WrongNumberOfArguments, 3, "Sub M()\r\nDim s\r\ns = Left(\"abc\", 1, 2)\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.WrongNumberOfArguments, 4, "Sub Foo(a)\r\nEnd Sub\r\nSub M()\r\nFoo 1, 2\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ArgumentNotOptional, 4, "Sub Foo(a)\r\nEnd Sub\r\nSub M()\r\nFoo\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ExpectedArray, 3, "Sub M()\r\nDim x As Long\r\nx(1) = 2\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ExpectedArray, 3, "Sub M()\r\nDim x\r\nx = UBound(\"abc\")\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeNotDefined, 2, "Sub M()\r\nDim r As Range\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.NamedArgumentNotFound, 4, "Sub Foo(a)\r\nEnd Sub\r\nSub M()\r\nFoo b:=1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ByRefTypeMismatch, 5, "Sub Foo(s As String)\r\nEnd Sub\r\nSub M()\r\nDim x As Long\r\nFoo x\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ByRefTypeMismatch, 5, "Sub Foo(n As Long)\r\nEnd Sub\r\nSub M()\r\nDim v As Variant\r\nFoo v\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ConstantExpressionRequired, 2, "Sub M()\r\nConst X = Len(\"a\")\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.InvalidStatementPlacement, 2, "Sub M()\r\nExit For\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.InvalidStatementPlacement, 2, "Sub M()\r\nExit Do\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ObjectRequired, 3, "Sub M()\r\nDim x As Long\r\nSet x = Nothing\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.DuplicateLabel, 3, "Sub M()\r\nDone:\r\nDone:\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ExpectedVariable, 3, "Sub M()\r\nDim s As String\r\nLen(s) = 3\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 4, "Sub M()\r\nDim a() As Long\r\nDim x As Long\r\nx = a\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 3, "Sub M()\r\nDim x\r\nx = Len(12345)\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.NotSupported, 3, "Sub M()\r\nDim x\r\nx = IMEStatus\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.NotSupported, 1, "Declare PtrSafe Sub F Lib \"k\" (ParamArray a() As Variant)\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 3, "Sub M()\r\nDim n As Long\r\nLSet n = \"x\"\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.InvalidTestProcedure, 2, "'@Test\r\nPrivate Sub T()\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.InvalidTestProcedure, 2, "'@Test\r\nPublic Sub T(n As Long)\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.InvalidTestProcedure, 2, "'@Test\r\nPublic Function T() As Long\r\nEnd Function\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 3, "Sub M()\r\nDim x\r\nx = Debug.Assert(True)\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.VariableNotDefined, 3, "Option Explicit\r\nSub M()\r\nAssert.AreEqual 1, 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.VariableNotDefined, 3, "Option Explicit\r\nSub M()\r\nvbang.Assert.AreEqual 1, 1\r\nEnd Sub\r\n")]
    public void Bind_ReportsTheDiagnostic(string id, int line, string source)
    {
        var diagnostics = Bind(source);

        Assert.Contains(diagnostics, d => d.Id == id && d.Line == line);
        Assert.All(diagnostics, d => Assert.Equal(FilePath, d.FilePath));
    }

    /// <summary>The Assert module exists only with the vbang reference in the manifest (ARCHITECTURE.md section 8).</summary>
    [Theory]
    [InlineData(DiagnosticIds.ProcedureNotDefined, 2, "Sub M()\r\nAssert.Equals 1, 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ProcedureNotDefined, 2, "Sub M()\r\nvbang.Assert.Equals 1, 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.TypeMismatch, 3, "Sub M()\r\nDim x\r\nx = Assert.IsTrue(True)\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.ArgumentNotOptional, 2, "Sub M()\r\nAssert.AreEqual 1\r\nEnd Sub\r\n")]
    [InlineData(DiagnosticIds.WrongNumberOfArguments, 2, "Sub M()\r\nAssert.Fail \"a\", \"b\"\r\nEnd Sub\r\n")]
    public void Bind_WithTheVbangReference_ReportsTheDiagnostic(string id, int line, string source)
    {
        var diagnostics = Bind(source, Vbang);

        Assert.Contains(diagnostics, d => d.Id == id && d.Line == line);
    }

    [Theory]
    [InlineData("Sub M()\r\nDim x As Long\r\nx = 1\r\nEnd Sub\r\n")]
    [InlineData("Option Explicit\r\nPublic Type T\r\n    X As Long\r\nEnd Type\r\nSub M()\r\nDim t As T\r\nt.X = 1\r\nEnd Sub\r\n")]
    [InlineData("Enum E\r\n    A\r\n    B = 5\r\nEnd Enum\r\nFunction F() As Long\r\nF = B\r\nEnd Function\r\n")]
    [InlineData("Function Fact(ByVal n As Long) As Long\r\nIf n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)\r\nEnd Function\r\n")]
    [InlineData("Sub M()\r\nOn Error GoTo Fail\r\nExit Sub\r\nFail:\r\nResume Next\r\nEnd Sub\r\n")]
    [InlineData("Property Get P() As Long\r\nP = 1\r\nEnd Property\r\nProperty Let P(v As Long)\r\nEnd Property\r\nSub M()\r\nP = P + 1\r\nEnd Sub\r\n")]
    // Constant expressions (MS-VBAL 5.2.3.2): the VBA library's enums as qualifiers, and enum members as array bounds and Optional defaults.
    [InlineData("Private Const LongLongType = VbVarType.vbLongLong\r\nPrivate Const Yes = VBA.vbYes\r\nSub M()\r\nDebug.Print LongLongType + Yes\r\nEnd Sub\r\n")]
    [InlineData("Private Enum Bounds\r\nFirst = 2\r\nLast = 5\r\nEnd Enum\r\nPrivate Type Table\r\nItems(First To Last) As Long\r\nEnd Type\r\nPrivate Const Width = Last - First\r\nSub M(Optional ByVal edge As Bounds = Last)\r\nDim t As Table\r\nt.Items(Width) = edge\r\nEnd Sub\r\n")]
    // The VBA library's modules qualify its functions without the VBA prefix (MS-VBAL 6.1): Strings.Join, Math.Abs.
    [InlineData("Sub M()\r\nDebug.Print Strings.Left(\"abc\", 1); Math.Abs(-1); Conversion.CStr(1); Strings.Len(\"x\")\r\nEnd Sub\r\n")]
    // Named arguments to library functions by their MS-VBAL 6.1 parameter names, and to late-bound members by name at run time.
    [InlineData("Sub M()\r\nDim o As Object\r\nDebug.Print Replace(\"aXb\", \"X\", \"-\", Count:=1); Left(String:=\"ab\", Length:=1)\r\no.Add Key:=\"a\", Item:=1\r\no.Sort \"x\", Order:=2\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\nDim c As New Collection\r\nc.Add 1, \"k\"\r\nDebug.Print c.Count; c(\"k\")\r\nEnd Sub\r\n")]
    [InlineData("Private Assert As Object\r\nSub M()\r\nAssert.Whatever 1\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\nAssert.AreEqual 1, 1\r\nEnd Sub\r\n")]
    // An Optional default of library.Enum.Member, which the VBE folds (MS-VBAL 6.1.1) and stdVBA's
    // stdError writes; it was VBA0015, Constant expression required (ROADMAP.md WP6).
    [InlineData("Sub M(Optional ByVal Crit As VBA.VbMsgBoxStyle = VBA.VbMsgBoxStyle.vbExclamation)\r\nDebug.Print Crit\r\nEnd Sub\r\n")]
    [InlineData("Sub M(Optional ByVal Kind As VbVarType = VBA.VbVarType.vbLongLong)\r\nDebug.Print Kind\r\nEnd Sub\r\n")]
    // VBA checks neither which End keyword closes a procedure nor which Exit leaves one; the VBE
    // accepts every pairing of the three, and stdVBA's sources rely on it (docs/vba-quirks.md, R3).
    [InlineData("Property Get P() As Long\r\nP = 1\r\nEnd Function\r\n")]
    [InlineData("Property Get P() As Long\r\nExit Function\r\nP = 1\r\nEnd Property\r\n")]
    [InlineData("Function F() As Long\r\nExit Property\r\nF = 1\r\nEnd Sub\r\n")]
    [InlineData("Sub S()\r\nExit Function\r\nEnd Sub\r\n")]
    public void Bind_AcceptsValidModules(string source)
    {
        var diagnostics = Bind(source);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
    }

    /// <summary>A named argument a library function does not have is a compile error in the VBE, as it is here (VBA0013).</summary>
    [Fact]
    public void Bind_NamedArgumentALibraryFunctionLacks_IsReported()
    {
        var diagnostics = Bind("Sub M()\r\nDebug.Print Left(Text:=\"hello\", Length:=2)\r\nEnd Sub\r\n");

        var reported = Assert.Single(diagnostics);
        Assert.Equal(DiagnosticIds.NamedArgumentNotFound, reported.Id);
        Assert.Contains("Named argument not found: 'Text'", reported.ToString(), StringComparison.Ordinal);
    }

    /// <summary>InStr and StrComp take no named arguments at all in Excel, whatever the Object Browser shows (docs/vba-quirks.md; the Strings golden's probes of 2026-09-09).</summary>
    [Theory]
    [InlineData("InStr(Start:=2, String1:=\"abcabc\", String2:=\"c\")", "Start")]
    [InlineData("InStr(1, \"abcabc\", \"c\", Compare:=vbBinaryCompare)", "Compare")]
    [InlineData("StrComp(String1:=\"a\", String2:=\"A\")", "String1")]
    [InlineData("StrComp(\"a\", \"A\", Compare:=vbTextCompare)", "Compare")]
    public void Bind_NamedArgumentToInStrOrStrComp_IsReported(string call, string name)
    {
        var diagnostics = Bind("Sub M()\r\nDebug.Print " + call + "\r\nEnd Sub\r\n");

        // The named arguments fall away, so a required parameter left unfilled is reported too, as it would be.
        Assert.Contains(diagnostics, d => d.Id == DiagnosticIds.NamedArgumentNotFound && d.ToString().Contains($"Named argument not found: '{name}'", StringComparison.Ordinal));
        Assert.All(diagnostics, d => Assert.Contains(d.Id, new[] { DiagnosticIds.NamedArgumentNotFound, DiagnosticIds.ArgumentNotOptional }));
    }

    /// <summary>
    /// An array element passed ByRef to a Declare hands the callee the element's address in VBA,
    /// and a Fortran-convention DLL reads and writes the elements after it (ROADMAP.md WP0). Since M7 C4
    /// the elements lie in the array's native storage, so the element itself travels by reference;
    /// a String element, which VBA converts on the way, and an element of another type than the
    /// parameter's are still refused rather than handed a temporary (ARCHITECTURE.md D20).
    /// </summary>
    [Fact]
    public void Bind_ArrayElementByRefToDeclare_TravelsByAddress_ButNotAStringOrAnotherType()
    {
        var diagnostics = Bind(string.Join("\r\n", [
            "Private Declare PtrSafe Sub MoveDoubles Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef destination As Double, ByRef source As Double, ByVal length As LongPtr)",
            "Private Declare PtrSafe Sub MoveAny Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)",
            "Sub M()",
            "Dim a(1 To 3) As Double, b(1 To 3) As Double, x As Double, s(1 To 2) As String, l(1 To 2) As Long",
            "MoveDoubles b(1), a(1), 24",
            "MoveAny b(1), a(1), 24",
            "MoveDoubles x, a(1), 8",
            "MoveDoubles x, x, 8",
            "MoveAny s(1), s(2), 8",
            "MoveDoubles l(1), x, 8",
            "End Sub",
            "",
        ]));

        var refused = diagnostics.Where(d => d.Id == DiagnosticIds.NotSupported).Select(d => d.ToString()).ToList();
        Assert.Equal(3, refused.Count);
        Assert.All(refused, message => Assert.Contains("needs an address of its own", message, StringComparison.Ordinal));
        Assert.DoesNotContain(diagnostics, d => d.Id != DiagnosticIds.NotSupported);
    }

    [Theory]
    [InlineData("'@Test\r\nPublic Sub T()\r\nAssert.AreEqual 1, 1\r\nAssert.AreNotEqual 1, 2, \"why\"\r\nAssert.IsTrue True\r\nAssert.IsFalse False, \"why\"\r\nAssert.IsNothing Nothing\r\nAssert.IsNotNothing Nothing\r\nAssert.Fail\r\nAssert.Fail \"why\"\r\nEnd Sub\r\n")]
    [InlineData("Sub T()\r\nvbang.Assert.AreEqual 1, 1\r\nVBANG.assert.istrue True\r\nEnd Sub\r\n")]
    [InlineData("Private Assert As Object\r\nSub M()\r\nAssert.Whatever 1\r\nEnd Sub\r\n")]
    public void Bind_WithTheVbangReference_AcceptsAssertCalls(string source)
    {
        var diagnostics = Bind(source, Vbang);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
    }

    [Theory]
    [InlineData("'@Test\r\nPublic Sub T()\r\nEnd Sub\r\n", true)]
    [InlineData("' @test  checks something\r\nSub T()\r\nEnd Sub\r\n", true)]
    [InlineData("Rem @Test\r\nSub T()\r\nEnd Sub\r\n", true)]
    [InlineData("' A header comment\r\n\r\n'@Test\r\n' More words\r\nSub T()\r\nEnd Sub\r\n", true)]
    [InlineData("Sub Other()\r\nEnd Sub\r\n'@Test\r\nSub T()\r\nEnd Sub\r\n", true)]
    [InlineData("Sub T()\r\nEnd Sub\r\n", false)]
    [InlineData("'@Tests\r\nSub T()\r\nEnd Sub\r\n", false)]
    [InlineData("' See @Test below\r\nSub T()\r\nEnd Sub\r\n", false)]
    [InlineData("'@Test\r\nSub Other()\r\nEnd Sub\r\nSub T()\r\nEnd Sub\r\n", false)]
    public void Bind_MarksTestProceduresFromTheAnnotation(string source, bool isTest)
    {
        var diagnostics = new List<Diagnostic>();

        var generated = ProjectCompiler.Generate("Test", [new SourceFile(FilePath, source)], diagnostics);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
        var text = Assert.Single(generated!).Text;
        var attributed = System.Text.RegularExpressions.Regex.IsMatch(text, @"\[global::VbaNg\.Runtime\.Hosting\.VbaTestAttribute\]\n#line [^\n]*\n    public static void T\(\)");
        Assert.Equal(isTest, attributed);
    }

    [Fact]
    public void Bind_ResolvesPublicMembersAcrossModules()
    {
        var diagnostics = new List<Diagnostic>();
        var sources = new[]
        {
            new SourceFile(@"C:\test\Lib.bas", "Attribute VB_Name = \"Lib\"\r\nPublic Const Max As Long = 3\r\nPublic Function Twice(n As Long) As Long\r\nTwice = n * 2\r\nEnd Function\r\n"),
            new SourceFile(@"C:\test\Main.bas", "Attribute VB_Name = \"Main\"\r\nSub Run()\r\nDim x As Long\r\nx = Twice(Max) + Lib.Twice(1)\r\nEnd Sub\r\n"),
        };

        var generated = ProjectCompiler.Generate("Test", sources, diagnostics);

        Assert.Empty(diagnostics.Select(d => d.ToString()));
        Assert.NotNull(generated);
        Assert.Contains(generated!, g => g.Module == "Main" && g.Text.Contains("global::Lib.Twice(", StringComparison.Ordinal));
    }

    private static readonly ProjectManifest Vbang = new([new ManifestReference(ProjectManifest.VbangLibrary)]);

    private static List<Diagnostic> Bind(string source, ProjectManifest? manifest = null)
    {
        var diagnostics = new List<Diagnostic>();
        ProjectCompiler.Generate("Test", [new SourceFile(FilePath, source)], diagnostics, manifest: manifest);
        return diagnostics;
    }
}
