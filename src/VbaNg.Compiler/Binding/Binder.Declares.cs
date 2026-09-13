using VbaNg.Compiler.Syntax;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// Declare statements (MS-VBAL 5.2.3.5): a procedure that lives in a DLL. The name, the library,
/// and the entry point come from the declaration; calls to it bind like any other procedure call,
/// and the emitter writes the P/Invoke (Declares golden).
/// </summary>
public sealed partial class Binder
{
    private void DeclareExternal(DeclareStatementSyntax syntax)
    {
        var name = syntax.Name.NameValue;
        var isFunction = syntax.ProcedureKeyword.Kind == SyntaxKind.FunctionKeyword;
        var symbol = new ProcedureSymbol(name, isFunction ? ProcedureKind.Function : ProcedureKind.Sub, module, ExternalBody(syntax))
        {
            IsPublic = !syntax.Modifiers.Any(SyntaxKind.PrivateKeyword),
            EmitName = MemberEmitName(name),
            ExternalLibrary = LibraryName(syntax.LibraryName),

            // Without an Alias the declared name is the entry point, as the error VBA raises for a missing one shows (Declares golden).
            ExternalEntryPoint = syntax.AliasName is { } alias ? alias.NameValue : name,
        };

        if (syntax.CDeclKeyword is not null)
        {
            Report(DiagnosticIds.NotSupported, syntax.CDeclKeyword, "CDecl is only honoured on the Macintosh; a CDecl Declare needs its own calling convention.");
        }

        if (AddMember(symbol, syntax.Name))
        {
            module.Externals.Add(symbol);
        }
    }

    /// <summary>The library name as written, without the quotes the lexer keeps.</summary>
    private static string LibraryName(SyntaxToken token) => token.Value as string ?? token.Text.Trim('"');

    /// <summary>
    /// A Declare has no body, but a procedure symbol needs a declaration to carry its parameters
    /// and its line for diagnostics; the statement itself stands in for one.
    /// </summary>
    private static ProcedureDeclarationSyntax ExternalBody(DeclareStatementSyntax syntax) =>
        new(
            syntax.Modifiers,
            syntax.ProcedureKeyword,
            null,
            syntax.Name,
            syntax.Parameters,
            syntax.AsClause,
            new SyntaxList<StatementSyntax>([]),
            new EndBlockStatementSyntax(syntax.ProcedureKeyword, syntax.ProcedureKeyword));

    /// <summary>
    /// How an external parameter travels (Declares golden): a ByVal String reaches the callee as a
    /// buffer it may write into, so the caller sees the change and the parameter is passed by
    /// reference in the generated code; As Any passes the variable's own storage.
    /// </summary>
    private void ResolveExternalSignature(ProcedureSymbol symbol)
    {
        ResolveSignature(symbol);
        foreach (var parameter in symbol.Parameters)
        {
            parameter.IsExternal = true;
            if (parameter.IsParamArray)
            {
                // A ParamArray reaches a DLL as a SafeArray of Variants, which needs its own marshaling.
                Report(DiagnosticIds.NotSupported, parameter.Syntax!.FirstToken()!, "A ParamArray parameter of a Declare needs its own marshaling.");
                continue;
            }

            if (ReferenceEquals(parameter.Type, VbaType.Any))
            {
                parameter.EmitByRef = !parameter.IsByVal;
                continue;
            }

            if (parameter.Type.IsString && parameter.IsByVal)
            {
                parameter.EmitByRef = true;
                continue;
            }

            if (symbol.IsRuntimeVarPtr && !parameter.IsByVal)
            {
                // VBE7's VarPtr takes whatever it is handed by reference, an array variable included (VarPtrArray).
                continue;
            }

            if (ExternalMarshalingGap(parameter) is { } gap)
            {
                Report(DiagnosticIds.NotSupported, parameter.Syntax!.FirstToken()!, gap);
            }
        }

        if (symbol.ReturnType is { } returnType && !symbol.IsRuntimeCallByName && (returnType.IsString || returnType.IsObject || returnType.ElementType is { IsRecord: true } || returnType.IsRecord || ReferenceEquals(returnType, VbaType.Any)))
        {
            Report(DiagnosticIds.NotSupported, symbol.Syntax.Name, $"A Declare returning {returnType.Name} needs its own marshaling.");
        }
    }

    /// <summary>
    /// Why a parameter cannot cross to a DLL as VBA passes it yet, or null when it can (MS-VBAL
    /// 5.2.3.5; Declares golden; ROADMAP.md M7 E5): a Variant ByRef travels as its VARIANT;
    /// or as a copy; an object as its interface pointer, or ByRef as the address of the pointer;
    /// a record ByRef as its block, its Strings in ANSI.
    /// </summary>
    private static string? ExternalMarshalingGap(ParameterSymbol parameter)
    {
        var type = parameter.Type;
        var how = parameter.IsByVal ? "ByVal" : "ByRef";
        if (type.IsVariant)
        {
            // ByVal is not the x64 convention's pointer to a copy: VariantChangeType given a ByVal Variant faulted inside Excel, so what VBA passes is not recorded (2026-09-11).
            return parameter.IsByVal ? "A Variant passed ByVal to a Declare needs VBA's own way of passing it, which is not recorded yet." : null;
        }

        if (type.IsObject)
        {
            // A type library's interface travels as that interface's pointer, queried for ByVal; ByRef, as the address of the variable's pointer, which the DLL fills (Declares golden: IPicture).
            return null;
        }

        if (type.IsRecord && !parameter.IsByVal)
        {
            return LaidOutAsVba(type.Record!) || DeclareImaged(type.Record!) ? null : $"A parameter of type {type.Name} passed ByRef to a Declare needs the record laid out as VBA lays it out: a Boolean, dynamic array, array of records, or fixed-length string member is not yet.";
        }

        if (type.IsArray)
        {
            // An array travels as the address of its descriptor pointer (SAFEARRAY**); VBA converts a String array's elements to ANSI, and a record array carries its record information, neither of which a Declare does yet.
            return type.ElementType is { IsString: true } or { IsRecord: true }
                ? $"An array of {type.ElementType.Name} passed to a Declare needs its elements converted as VBA converts them."
                : null;
        }

        // A Boolean ByRef travels as its two bytes, a String ByRef as the address of a copy in ANSI (Declares golden).
        return type.Kind == TypeKind.FixedString || type.IsRecord
            ? $"A parameter of type {type.Name} passed {how} to a Declare needs its own marshaling."
            : null;
    }

    /// <summary>
    /// A record with a fixed-length string member, which VBA hands a DLL as a copy with each such
    /// string as its ANSI bytes and reads back after the call (RecordImage.DeclareImage; Declares
    /// golden): every member a number, a Date, a Boolean, a fixed-length string, a fixed-size array
    /// of numbers, or a record of the same, since the copy carries no String, Variant, or object.
    /// </summary>
    internal static bool DeclareImaged(RecordSymbol record) => HasFixedString(record) && Imageable(record);

    private static bool HasFixedString(RecordSymbol record) =>
        record.Fields.Any(f => f.Type.Kind == TypeKind.FixedString || (f.Type.IsRecord && HasFixedString(f.Type.Record!)));

    private static bool Imageable(RecordSymbol record) =>
        record.Fields.All(f => f.Type.Kind == TypeKind.FixedString || (f.Type.IsRecord ? Imageable(f.Type.Record!) : f.IsEmbeddedArray ? IsPlainScalar(f.Type.ElementType!) : IsPlainScalar(f.Type)));

    private static bool IsPlainScalar(VbaType type) =>
        type.IsEnum || (type.Kind == TypeKind.Builtin && type.VarType is Runtime.VarType.Byte or Runtime.VarType.Integer or Runtime.VarType.Long or Runtime.VarType.LongLong
            or Runtime.VarType.Single or Runtime.VarType.Double or Runtime.VarType.Currency or Runtime.VarType.Date or Runtime.VarType.Boolean);

    /// <summary>A record whose block is byte for byte VBA's, so a DLL can read it: no fixed-size array member but one laid inline, no fixed-length string member, nested records and arrays of them included.</summary>
    private static bool LaidOutAsVba(RecordSymbol record) =>
        record.Fields.All(f => (f.Bounds is null || f.IsEmbeddedArray) && f.Type.Kind != TypeKind.FixedString && (!f.Type.IsRecord || LaidOutAsVba(f.Type.Record!))
            && (!f.IsEmbeddedArray || f.Type.ElementType is not { IsRecord: true } element || LaidOutAsVba(element.Record!)));
}
