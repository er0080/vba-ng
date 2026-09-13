using VbaNg.Compiler;
using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// <c>Application.Run</c> of the project's own procedures (ARCHITECTURE.md D23), through the Run
/// tables the compiler gives standard modules, with the outcomes Excel showed for each case
/// (docs/vba-quirks.md, "Host, Application.Run"). The calls go straight to the runtime helper the
/// generated code calls; a name the project does not define goes on to the Application, here
/// Nothing, so it raises 91 where Excel would raise 1004.
/// </summary>
public sealed class ApplicationRunTests : IDisposable
{
    private const int RunDispId = 259;
    private readonly string workDir = Path.Combine(Path.GetTempPath(), "vbang-tests", Guid.NewGuid().ToString("N"));
    private readonly ProjectHost host;
    private readonly Type targets;

    public ApplicationRunTests()
    {
        var projectDir = Path.Combine(workDir, "Runner.vbang");
        Directory.CreateDirectory(projectDir);
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
            "Public Function Twice(ByVal x As Long) As Long",
            "    Twice = x * 2",
            "End Function",
            "Public Sub Append(ByRef s As String)",
            "    s = s & \"!\"",
            "End Sub",
            "Public Function Kind(ByVal v As Variant) As String",
            "    Kind = TypeName(v)",
            "End Function",
            "Public Function Total(ParamArray p() As Variant) As Long",
            "    Total = UBound(p) - LBound(p) + 1",
            "End Function",
            "Public Sub Fails()",
            "    Err.Raise 5",
            "End Sub",
            "Public Sub Dup()",
            "End Sub",
            "Public Property Get Prop() As Long",
            "    Prop = 3",
            "End Property",
            ""]));
        File.WriteAllText(Path.Combine(projectDir, "Other.bas"), "Attribute VB_Name = \"Other\"\r\nPublic Sub Dup()\r\n    Targets.Counter = Targets.Counter + 20\r\nEnd Sub\r\n");
        var build = ProjectCompiler.Build(projectDir);
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));
        host = new ProjectHost(Path.Combine(workDir, "shadow"));
        targets = host.Load(projectDir).Assembly.GetType("Targets")!;
    }

    public void Dispose()
    {
        host.Dispose();
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
    public void Run_BareQualifiedAndBookQualifiedNames_RunTheProceduresPrivateOnesToo()
    {
        Run("Bump");
        Run("Targets.Bump");
        Run("'Runner.xlsm'!targets.BUMP");
        Run("Runner.xlsm!Bump");
        Run("Hidden");
        Run("Other.Dup");

        Assert.Equal(124, Counter());
    }

    [Fact]
    public void Run_Function_ReturnsItsValueAndASubReturnsEmpty()
    {
        Assert.Equal(42, Coerce.ToInt32(Run("Twice", Variant.FromInt32(21))));
        Assert.Equal(VarType.Empty, Run("Bump").Type);
    }

    [Fact]
    public void Run_Arguments_PassByValueAndKeepTheirSubtype()
    {
        var text = Variant.FromString("a");
        Run("Append", text);

        Assert.Equal("a", Coerce.ToString(text));
        Assert.Equal("Integer", Coerce.ToString(Run("Kind", Variant.FromInt16(5))));
        Assert.Equal(3, Coerce.ToInt32(Run("Total", Variant.FromInt32(1), Variant.FromInt32(2), Variant.FromInt32(3))));
        Assert.Equal(0, Coerce.ToInt32(Run("Total")));
    }

    [Fact]
    public void Run_TrailingMissingArguments_AreDroppedBeforeCounting()
    {
        Assert.Equal(6, Coerce.ToInt32(Run("Twice", Variant.FromInt32(3), Variant.Missing)));
        Assert.Equal(449, Raises(() => Run("Twice", Variant.Missing)).Number);
        Assert.Equal(450, Raises(() => Run("Twice", Variant.Missing, Variant.FromInt32(3))).Number);
    }

    [Fact]
    public void Run_ArgumentErrors_ReachTheCallerFromTheProject()
    {
        var tooFew = Raises(() => Run("Twice"));
        var tooMany = Raises(() => Run("Twice", Variant.FromInt32(1), Variant.FromInt32(2)));
        var mismatch = Raises(() => Run("Twice", Variant.FromString("x")));

        Assert.Equal((449, "VBAProject"), (tooFew.Number, tooFew.Source));
        Assert.Equal((450, "VBAProject"), (tooMany.Number, tooMany.Source));
        Assert.Equal((13, "VBAProject"), (mismatch.Number, mismatch.Source));
    }

    [Fact]
    public void Run_BareNameTwoModulesDefine_Raises1004FromExcel()
    {
        var error = Raises(() => Run("Dup"));

        Assert.Equal(1004, error.Number);
        Assert.Equal("Microsoft Excel", error.Source);
        Assert.Equal("Cannot run the macro 'Dup'. The macro may not be available in this workbook or all macros may be disabled.", error.Description);
    }

    [Fact]
    public void Run_ErrorTheProcedureLeavesUnhandled_EscapesAsOneNoHandlerTraps()
    {
        var escaped = Assert.Throws<UnhandledErrorException>(() => Run("Fails"));

        Assert.Equal(5, escaped.Error!.Number);
        Assert.Null(VbaException.From(escaped));
    }

    [Fact]
    public void Run_NameTheProjectLacks_GoesToTheApplication()
    {
        // A property, a missing procedure, and a missing module are Excel's to report; the Application here is Nothing.
        Assert.Equal(91, Raises(() => Run("Prop")).Number);
        Assert.Equal(91, Raises(() => Run("Targets.NoSuch")).Number);
        Assert.Equal(91, Raises(() => Run("NoSuch.Bump")).Number);
    }

    private Variant Run(string macro, params Variant[] arguments) =>
        ApplicationRun.Invoke(targets, null, RunDispId, [Variant.FromString(macro), .. arguments]);

    private int Counter() => (int)targets.GetField("Counter")!.GetValue(null)!;

    private static VbaException Raises(Func<Variant> call) => Assert.Throws<VbaException>(() => call());
}
