using System.Globalization;

namespace VbaNg.Compiler;

public enum DiagnosticSeverity
{
    Warning,
    Error,
}

/// <summary>
/// A compiler diagnostic. <see cref="ToString"/> renders the MSBuild canonical format
/// <c>path(line,col): error VBA0001: message</c> (ARCHITECTURE.md D5, CLAUDE.md R11).
/// </summary>
public sealed record Diagnostic(
    string Id,
    DiagnosticSeverity Severity,
    string Message,
    string FilePath,
    int Line,
    int Column)
{
    public bool IsError => Severity == DiagnosticSeverity.Error;

    public override string ToString()
    {
        var severity = IsError ? "error" : "warning";
        return string.Create(CultureInfo.InvariantCulture, $"{FilePath}({Line},{Column}): {severity} {Id}: {Message}");
    }
}

/// <summary>
/// Diagnostic ids are stable, never renumbered or reused, and documented with a triggering
/// snippet in docs/diagnostics.md (CLAUDE.md R11).
/// </summary>
public static class DiagnosticIds
{
    /// <summary>Syntax error: the source could not be tokenized or parsed.</summary>
    public const string SyntaxError = "VBA0001";

    /// <summary>The construct is valid VBA but not supported by this version of the compiler.</summary>
    public const string NotSupported = "VBA0002";

    /// <summary>Internal compiler error: generated C# failed to compile. Always a compiler bug.</summary>
    public const string InternalError = "VBA0003";

    /// <summary>A conditional compilation expression (#If, #ElseIf, #Const) could not be evaluated.</summary>
    public const string ConditionalCompilationError = "VBA0004";

    /// <summary>Variable not defined: an undeclared name under Option Explicit (MS-VBAL 5.2.1.2).</summary>
    public const string VariableNotDefined = "VBA0005";

    /// <summary>Ambiguous name detected: two module-level members, or two locals, share a name.</summary>
    public const string AmbiguousName = "VBA0006";

    /// <summary>Sub or Function not defined: a call to an unknown procedure.</summary>
    public const string ProcedureNotDefined = "VBA0007";

    /// <summary>Label not defined: GoTo, GoSub, On Error GoTo, or Resume names an unknown label.</summary>
    public const string LabelNotDefined = "VBA0008";

    /// <summary>Wrong number of arguments or invalid property assignment.</summary>
    public const string WrongNumberOfArguments = "VBA0009";

    /// <summary>Argument not optional: a required parameter has no argument.</summary>
    public const string ArgumentNotOptional = "VBA0010";

    /// <summary>Expected array: ReDim, Erase, or an index applied to something that is not an array.</summary>
    public const string ExpectedArray = "VBA0011";

    /// <summary>User-defined type not defined: an As clause names an unknown type.</summary>
    public const string TypeNotDefined = "VBA0012";

    /// <summary>Named argument not found.</summary>
    public const string NamedArgumentNotFound = "VBA0013";

    /// <summary>ByRef argument type mismatch: a typed variable passed to a ByRef parameter of another type.</summary>
    public const string ByRefTypeMismatch = "VBA0014";

    /// <summary>Constant expression required: a Const value, array bound, or Optional default is not a constant.</summary>
    public const string ConstantExpressionRequired = "VBA0015";

    /// <summary>A statement is not allowed where it appears: Exit For outside For, Exit Do outside Do, Event or Implements outside a class module, WithEvents in a procedure or a standard module, a leading . or ! with no With block.</summary>
    public const string InvalidStatementPlacement = "VBA0016";

    /// <summary>Object required: Set with a non-object variable, or a member access on a value that cannot have members.</summary>
    public const string ObjectRequired = "VBA0017";

    /// <summary>Duplicate label within the procedure.</summary>
    public const string DuplicateLabel = "VBA0018";

    /// <summary>The left side of an assignment is not a variable, array element, or property.</summary>
    public const string ExpectedVariable = "VBA0019";

    /// <summary>Type mismatch detectable at compile time: an array assigned to a scalar, a record to a Variant, an Enum or type name used as a value.</summary>
    public const string TypeMismatch = "VBA0020";

    /// <summary>A '@Test annotation on something other than a Public Sub without parameters.</summary>
    public const string InvalidTestProcedure = "VBA0021";

    /// <summary>The project manifest (vbang.json) is not valid JSON or a reference has no name.</summary>
    public const string InvalidManifest = "VBA0022";

    /// <summary>Can't find project or library: a manifest reference names a type library that is not registered.</summary>
    public const string ReferenceNotFound = "VBA0023";

    /// <summary>A class with a document module's attributes that the workbook beside the folder does not name as a CodeName; it compiles as a class, and its event procedures never run.</summary>
    public const string UnboundDocumentModule = "VBA0024";
}
