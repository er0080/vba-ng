using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// VarPtr, StrPtr, and ObjPtr (ARCHITECTURE.md section 5, "Pointers"; ROADMAP.md M7 E1). StrPtr
/// is the BSTR a String holds, ObjPtr the interface pointer an object reference is, and VarPtr
/// the address of a variable's storage, let through for storage whose address is stable for the
/// variable's lifetime: a local, a parameter, or a With alias on the stack, a module-level or
/// Static variable of a standard module whose type holds no managed reference, and an element of
/// an array in its native storage. An object variable is its interface pointer since E3, so its
/// address is the pointer's; a record's members lie inline in it since C5, a fixed-length string
/// reaches VarPtr as a temporary copy, and a class instance's variables lie in its block since E4.
/// </summary>
public sealed partial class Binder
{
    private static PointerKind? PointerKindOf(string name) => name.ToUpperInvariant() switch
    {
        "VARPTR" => PointerKind.Variable,
        "STRPTR" => PointerKind.String,
        "OBJPTR" => PointerKind.Object,
        _ => null,
    };

    private BoundExpression BindPointer(SyntaxNode syntax, SyntaxToken keyword, PointerKind kind, ExpressionSyntax operandSyntax)
    {
        var operand = BindExpression(operandSyntax);
        switch (kind)
        {
            case PointerKind.String:
                if (operand is BoundConstant { Constant.Name: var constant } && constant.Equals("vbNullString", StringComparison.OrdinalIgnoreCase))
                {
                    // vbNullString is the null BSTR, which StrPtr reports as 0 (Memory golden).
                    return new BoundLiteral(syntax, Variant.FromInt64(0), VbaType.LongLong);
                }

                return new BoundPointer(syntax, kind, Convert(operand, VbaType.String));
            case PointerKind.Object:
                if (!operand.Type.IsObject && !operand.Type.IsVariant)
                {
                    Report(DiagnosticIds.TypeMismatch, keyword, "Type mismatch: ObjPtr needs an object.");
                }

                return new BoundPointer(syntax, kind, Convert(operand, VbaType.Variant));
            default:
                if (AddressProblem(operand) is { } problem)
                {
                    Report(DiagnosticIds.NotSupported, keyword, problem);
                    return new BoundLiteral(syntax, Variant.FromInt64(0), VbaType.LongLong);
                }

                return new BoundPointer(syntax, kind, operand);
        }
    }

    /// <summary>Why VarPtr cannot report a stable address for the operand yet; null when it can.</summary>
    private static string? AddressProblem(BoundExpression operand)
    {
        if (operand.Type.Kind == TypeKind.FixedString && operand is BoundVariable or BoundField or BoundElement)
        {
            // A fixed-length string is no BSTR, so VBA hands VarPtr a temporary copy of it: its address is the copy's, and what lies there is not StrPtr's (Memory golden).
            return null;
        }

        switch (operand)
        {
            case BoundVariable { Variable: var variable }:
                if (variable.Type.IsArray)
                {
                    return $"VarPtr of the array '{variable.Name}' is the address of its descriptor pointer, which a Declare of VBE7's VarPtr taking the array (VarPtrArray) reports.";
                }

                if (variable.Type.IsObject && (variable.IsWithEvents || ReferenceEquals(variable, variable.Module?.Me) || ReferenceEquals(variable, variable.Module?.Instance)))
                {
                    return $"VarPtr of '{variable.Name}' needs an object variable; a WithEvents variable, Me, and a predeclared instance hold the object where the runtime keeps it.";
                }

                // Every other variable's storage holds no managed reference and stays where it is: a local or a parameter on the stack, a standard module's variable in a static outside the GC heap, a class instance's in its block (E4).
                return null;
            case BoundField { Record: var owner }:
                return AddressProblem(owner);
            case BoundElement { Array.Type.IsArray: true }:
                // Every element lies in the array's native storage, a record's too since C5.
                return null;
            case BoundElement { Array.Type.IsVariant: true }:
                // An element of the array a Variant holds lies in that array's storage too (Memory golden).
                return null;
            case { IsLValue: false, Type.IsObject: false, Type.IsArray: false }:
                // A value that is not a variable (54& + n, a literal) reaches VarPtr as a temporary holding it, as ByRef As Any hands a DLL one: the temporary's address, for the statement (Memory golden).
                return null;
            default:
                return "VarPtr needs a variable, an array element, or a member of a user-defined type.";
        }
    }

    /// <summary>An element type whose elements a DLL can reach at their address in the array's storage: anything but strings, which VBA converts on the way, objects, and records.</summary>
    private static bool IsNativeElement(VbaType type) =>
        type.Kind == TypeKind.Builtin && !type.IsString && !type.IsObject && !ReferenceEquals(type, VbaType.Any);
}
