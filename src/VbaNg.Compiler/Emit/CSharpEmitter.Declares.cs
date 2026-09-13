using System.Globalization;

using VbaNg.Compiler.Binding;
using VbaNg.Compiler.Syntax;

namespace VbaNg.Compiler.Emit;

/// <summary>
/// Declare statements (MS-VBAL 5.2.3.5) as P/Invoke: one extern per declaration with the entry
/// point the Alias names, and a wrapper that marshals the parameters VBA passes differently from
/// .NET and turns a missing library or entry point into the errors VBA raises (Declares golden).
/// </summary>
public sealed partial class CSharpEmitter
{
    private const string Interop = "global::System.Runtime.InteropServices.";

    private void EmitExternals()
    {
        foreach (var external in Symbol.Externals)
        {
            EmitExternal(external);
        }
    }

    private void EmitExternal(ProcedureSymbol symbol)
    {
        if (symbol.IsRuntimeVarPtr)
        {
            EmitRuntimeVarPtr(symbol);
            return;
        }

        if (symbol.IsRuntimeCallByName)
        {
            EmitRuntimeCallByName(symbol);
            return;
        }

        var library = symbol.ExternalLibrary!;
        var entryPoint = symbol.ExternalEntryPoint!;
        var externName = "__extern_" + symbol.EmitName.TrimStart('@');
        var externReturn = ExternType(symbol.ReturnType);
        writer.HiddenLine($"[{Interop}DllImport({Quote(library)}, EntryPoint = {Quote(entryPoint)}, ExactSpelling = true)]");
        writer.HiddenLine($"private static extern {externReturn} {externName}({string.Join(", ", symbol.Parameters.Select(ExternParameter))});");
        writer.Line();

        var access = symbol.IsPublic ? "public" : "internal";
        var shared = Symbol.Kind == ModuleKind.Class ? string.Empty : "static ";
        var returnType = symbol.ReturnType?.CSharpName ?? "void";
        writer.MappedLine(Line(symbol.Syntax), $"{access} {shared}{returnType} {symbol.EmitName}({string.Join(", ", symbol.Parameters.Select(ParameterDeclaration))})");
        writer.Open();

        // A ByVal String travels as an ANSI buffer the callee may write into (Declares golden).
        var buffers = symbol.Parameters.Where(IsStringBuffer).ToList();
        foreach (var parameter in buffers)
        {
            writer.HiddenLine($"var {BufferName(parameter)} = {R}Native.Ansi({parameter.EmitName});");
        }

        // A ByRef Variant or array is the DLL's to change: what it leaves there becomes the runtime's after the call (Native.TakeOver).
        var variants = symbol.Parameters.Where(p => (p.Type.IsVariant || p.Type.IsArray) && !p.IsByVal).ToList();
        foreach (var parameter in variants)
        {
            writer.HiddenLine($"var {BeforeName(parameter)} = {parameter.EmitName}{(parameter.Type.IsArray ? ".Descriptor" : string.Empty)};");
        }

        // A String ByRef travels as a BSTR of its own holding the text in ANSI, read back after the call (Native.AnsiCopy, FromAnsiCopy).
        var strings = symbol.Parameters.Where(p => p.Type.IsVariableString && !p.IsByVal).ToList();
        foreach (var parameter in strings)
        {
            writer.HiddenLine($"var {AnsiName(parameter)} = {R}Native.AnsiCopy({parameter.EmitName});");
            writer.HiddenLine($"var {OriginalName(parameter)} = {AnsiName(parameter)};");
        }

        writer.HiddenLine("try");
        writer.Open();
        var call = $"{externName}({string.Join(", ", symbol.Parameters.Select(Forward))})";
        writer.HiddenLine(symbol.ReturnType is null ? call + ";" : $"return {FromExtern(call, symbol.ReturnType)};");
        writer.Close();
        writer.HiddenLine("catch (global::System.DllNotFoundException)");
        writer.Open();
        writer.HiddenLine($"throw {R}Native.LibraryNotFound({Quote(library)});");
        writer.Close();
        writer.HiddenLine("catch (global::System.EntryPointNotFoundException)");
        writer.Open();
        writer.HiddenLine($"throw {R}Native.EntryPointNotFound({Quote(entryPoint)}, {Quote(library)});");
        writer.Close();
        if (buffers.Count > 0 || variants.Count > 0 || strings.Count > 0)
        {
            writer.HiddenLine("finally");
            writer.Open();
            foreach (var parameter in buffers)
            {
                writer.HiddenLine($"{R}ObjectRefs.Assign(ref {parameter.EmitName}, {BufferName(parameter)}.ReadText());");
                writer.HiddenLine($"{BufferName(parameter)}.Dispose();");
            }

            foreach (var parameter in variants)
            {
                writer.HiddenLine($"{R}Native.TakeOver({BeforeName(parameter)}, ref {parameter.EmitName});");
            }

            foreach (var parameter in strings)
            {
                writer.HiddenLine($"{R}Native.FromAnsiCopy({OriginalName(parameter)}, {AnsiName(parameter)}, ref {parameter.EmitName});");
            }

            writer.Close();
        }

        writer.Close();
        writer.Hidden();
        writer.Line();
    }

    private static bool IsStringBuffer(ParameterSymbol parameter) => parameter.Type.IsVariableString && parameter.IsByVal;

    private static string BufferName(ParameterSymbol parameter) => "__buffer_" + parameter.EmitName.TrimStart('@');

    private static string BeforeName(ParameterSymbol parameter) => "__before_" + parameter.EmitName.TrimStart('@');

    private static string AnsiName(ParameterSymbol parameter) => "__ansi_" + parameter.EmitName.TrimStart('@');

    private static string OriginalName(ParameterSymbol parameter) => "__original_" + parameter.EmitName.TrimStart('@');

    /// <summary>The C# type the DLL sees for a parameter: a pointer for a string buffer, the variable's own storage for As Any, the type itself otherwise.</summary>
    private static string ExternParameter(ParameterSymbol parameter)
    {
        if (ReferenceEquals(parameter.Type, VbaType.Any))
        {
            return parameter.IsByVal ? $"nint {parameter.EmitName}" : $"ref byte {parameter.EmitName}";
        }

        if (IsStringBuffer(parameter))
        {
            return $"nint {parameter.EmitName}";
        }

        if (parameter.Type.IsRecord)
        {
            // A record travels as its block, which the call site hands over as bytes (DeclareRecord).
            return $"ref byte {parameter.EmitName}";
        }

        if (parameter.Type.IsVariableString)
        {
            // A String ByRef travels as the address of a BSTR holding its text in ANSI (Native.AnsiCopy; Declares golden).
            return $"ref nint {parameter.EmitName}";
        }

        if (parameter.Type.IsArray)
        {
            // An array travels as the address of its descriptor pointer (SAFEARRAY**), which the DLL may replace (Declares golden).
            return $"ref nint {parameter.EmitName}";
        }

        if (parameter.Type.IsObject)
        {
            // An object travels as its interface pointer; ByRef as the address of the pointer, which the DLL may overwrite without AddRef, as in VBA (Declares golden).
            return parameter.IsByVal ? $"nint {parameter.EmitName}" : $"ref nint {parameter.EmitName}";
        }

        var type = ExternType(parameter.Type);
        return parameter.EmitByRef ? $"ref {type} {parameter.EmitName}" : $"{type} {parameter.EmitName}";
    }

    /// <summary>A Boolean crosses as the two-byte value VBA passes; every other type crosses as itself.</summary>
    private static string ExternType(VbaType? type) => type switch
    {
        null => "void",
        _ when ReferenceEquals(type, VbaType.Boolean) => "short",
        _ when ReferenceEquals(type, VbaType.Any) || type.IsArray => "nint",
        _ => type.CSharpName,
    };

    private static string Forward(ParameterSymbol parameter)
    {
        if (IsStringBuffer(parameter))
        {
            return BufferName(parameter) + ".Pointer";
        }

        if (parameter.Type.IsVariableString)
        {
            return "ref " + AnsiName(parameter);
        }

        if (parameter.Type.IsArray)
        {
            return $"ref global::System.Runtime.CompilerServices.Unsafe.As<{R}VbaArray, nint>(ref {parameter.EmitName})";
        }

        if (parameter.Type.IsObject)
        {
            return parameter.IsByVal && parameter.Type.IsCom
                ? $"{R}Native.InterfacePointer({parameter.EmitName}, new global::System.Guid({Quote(parameter.Type.ComInterface!.Guid.ToString("D"))}))"
                : parameter.IsByVal
                ? $"(nint){R}Pointers.ObjPtr({R}Variant.FromObject({parameter.EmitName}))"
                : $"ref global::System.Runtime.CompilerServices.Unsafe.As<{SlotType(parameter.Type)}, nint>(ref {parameter.EmitName})";
        }

        if (parameter.EmitByRef)
        {
            return "ref " + parameter.EmitName;
        }

        return ReferenceEquals(parameter.Type, VbaType.Boolean)
            ? $"({parameter.EmitName} ? (short)-1 : (short)0)"
            : parameter.EmitName;
    }

    private static string FromExtern(string call, VbaType returnType) =>
        ReferenceEquals(returnType, VbaType.Boolean) ? $"{call} != 0"
        : returnType.IsArray ? $"{R}Native.ArrayResult({call}, {VarTypeName(returnType.ElementType!)})"
        : returnType.IsVariant ? $"{R}Native.VariantResult({call})"
        : call;

    /// <summary>
    /// An As Any argument (MS-VBAL 5.2.3.5): the variable's own storage travels to the DLL, so the
    /// call site hands the callee a reference to those bytes.
    /// </summary>
    private string AnyArgument(BoundArgument argument)
    {
        var parameter = argument.Parameter;
        if (argument.Value is null)
        {
            Report(parameter.Syntax!, $"'{parameter.Name}' is As Any and needs an argument.");
            return "ref " + NewTemporary();
        }

        if (parameter.IsByVal)
        {
            return $"(nint){Convert(EmitExpression(argument.Value), VbaType.LongLong)}";
        }

        if (argument.Mode == ArgumentMode.ByVal)
        {
            // ByVal at the call: the value itself is the address the DLL receives (MoveMemory x, ByVal pointer, n).
            return $"ref {R}Pointers.At({Convert(EmitExpression(argument.Value), VbaType.LongLong)})";
        }

        if (argument.Value.IsLValue && argument.Value.Type.IsRecord)
        {
            return DeclareRecord(argument.Value);
        }

        if (argument.Value is BoundElement { Array.Type.IsArray: true } element)
        {
            // An array element travels as its place in the array's native storage (M7 C4).
            return "ref " + ElementReference(element, "byte");
        }

        if (argument.Value.Type.IsObject && SlotReference(argument.Value, argument.Value.Type) is { } slot)
        {
            // A typed object variable's storage is its interface pointer (ROADMAP.md M7 E3): the DLL reads or writes the pointer itself, without AddRef, as in VBA (Memory golden).
            return $"ref global::System.Runtime.CompilerServices.Unsafe.As<{SlotType(argument.Value.Type)}, byte>(ref {slot})";
        }

        var value = EmitExpression(argument.Value);
        var fixedSize = !(value.Type.IsVariant || value.Type.IsArray || value.Type.IsRecord || value.Type.IsObject || value.Type.IsString);
        if (!argument.Value.IsLValue && fixedSize)
        {
            // A value that is not a variable travels as the address of a temporary holding it, a Boolean as its two bytes (MoveMemory ByVal p, 42&, 4; Memory golden).
            var temporary = NewTemporary();
            var boolean = ReferenceEquals(value.Type, VbaType.Boolean);
            before.Add(boolean
                ? $"short {temporary} = {value.Code} ? (short)-1 : (short)0;"
                : $"{value.Type.CSharpName} {temporary} = {value.Code};");
            return $"ref global::System.Runtime.CompilerServices.Unsafe.As<{(boolean ? "short" : value.Type.CSharpName)}, byte>(ref {temporary})";
        }

        if (!argument.Value.IsLValue || !fixedSize)
        {
            Report(argument.Value.Syntax, $"An argument for the As Any parameter '{parameter.Name}' must be a variable of a fixed-size type.");
            return "ref " + NewTemporary();
        }

        // Unsafe.As reinterprets the variable's storage without copying it, which is what VBA passes.
        return $"ref global::System.Runtime.CompilerServices.Unsafe.As<{value.Type.CSharpName}, byte>(ref {value.Code})";
    }

    /// <summary>
    /// A record for a ByRef As Any parameter (MS-VBAL 5.2.3.5): its block travels to the DLL. A
    /// record with String members travels as a copy whose strings are ANSI, copied back after
    /// the call with the strings converted again, so the callee never sees the variable's own
    /// BSTRs (Memory golden); a record without them travels as itself.
    /// </summary>
    private string DeclareRecord(BoundExpression record) => DeclareRecord(EmitExpression(record).Code, record.Type, record.Syntax);

    private string DeclareRecord(string storage, VbaType type, SyntaxNode syntax)
    {
        var bytes = $"global::System.Runtime.CompilerServices.Unsafe.As<{type.CSharpName}, byte>";
        var members = new List<string>();
        if (!StringMembers(type.Record!, string.Empty, members))
        {
            if (Binder.DeclareImaged(type.Record!))
            {
                // VBA hands the DLL a copy with each fixed-length string as its ANSI bytes, and reads the copy back after the call (Declares golden).
                var image = NewTemporary();
                before.Add($"var {image} = {L}RecordImage.DeclareImage({storage});");
                after.Add($"{L}RecordImage.FromDeclareImage(ref {storage}, {image});");
                return $"ref {image}[0]";
            }

            Report(syntax, $"A '{type.Name}' with a fixed-length string member and a String, Variant, object, or dynamic array member handed to a Declare needs its ANSI layout.");
        }

        if (members.Count == 0)
        {
            return $"ref {bytes}(ref {storage})";
        }

        var copy = NewTemporary();
        before.Add($"var {copy} = {storage}.Copy();");
        before.AddRange(members.Select(member => $"{R}Native.ToAnsi(ref {copy}{member});"));
        after.AddRange(members.Select(member => $"{R}Native.FromAnsi(ref {copy}{member});"));
        // The variable takes a copy of its own (AssignRecord), and the statement releases the temporary.
        after.Add($"{R}ObjectRefs.AssignRecord(ref {storage}, {copy}.Copy());");
        ReleaseAfter(copy, type);
        return $"ref {bytes}(ref {copy})";
    }

    /// <summary>The paths of a record's variable-length String members, its nested records' included; false when it has a fixed-length string member, which VBA lays out in ANSI for the call.</summary>
    private static bool StringMembers(RecordSymbol record, string prefix, List<string> paths)
    {
        var laidOut = true;
        foreach (var field in record.Fields)
        {
            var path = prefix + "." + field.EmitName;
            if (field.Type.IsVariableString)
            {
                paths.Add(path);
            }
            else if (field.Type.Kind == TypeKind.FixedString)
            {
                laidOut = false;
            }
            else if (field.Type.IsRecord)
            {
                laidOut &= StringMembers(field.Type.Record!, path, paths);
            }
        }

        return laidOut;
    }

    /// <summary>A Declare of VBE7's VarPtr (VarPtrArray and its kin): the runtime answers with the address of what the caller hands it, as VBE7 does, and so works where VBE7 is not loaded too.</summary>
    private void EmitRuntimeVarPtr(ProcedureSymbol symbol)
    {
        var access = symbol.IsPublic ? "public" : "internal";
        var shared = Symbol.Kind == ModuleKind.Class ? string.Empty : "static ";
        var returnType = symbol.ReturnType?.CSharpName ?? "void";
        writer.MappedLine(Line(symbol.Syntax), $"{access} {shared}{returnType} {symbol.EmitName}({string.Join(", ", symbol.Parameters.Select(ParameterDeclaration))})");
        writer.Open();
        if (symbol.Parameters is [{ EmitByRef: true } target] && symbol.ReturnType is { } result)
        {
            writer.HiddenLine($"return ({result.CSharpName}){R}Pointers.Address(ref {target.EmitName});");
        }
        else
        {
            Report(symbol.Syntax, "VBE7's VarPtr takes one ByRef parameter and returns its address.");
        }

        writer.Close();
        writer.Hidden();
        writer.Line();
    }

    /// <summary>
    /// A Declare of VBE7's rtcCallByName: the runtime answers it as CallByName (Native.CallByName).
    /// stdVBA declares it two ways, returning the Variant, or taking the result Variant first and
    /// returning a Long, which at the machine level is the same call: the hidden result pointer
    /// comes first, and the function hands it back.
    /// </summary>
    private void EmitRuntimeCallByName(ProcedureSymbol symbol)
    {
        var access = symbol.IsPublic ? "public" : "internal";
        var shared = Symbol.Kind == ModuleKind.Class ? string.Empty : "static ";
        var returnType = symbol.ReturnType?.CSharpName ?? "void";
        writer.MappedLine(Line(symbol.Syntax), $"{access} {shared}{returnType} {symbol.EmitName}({string.Join(", ", symbol.Parameters.Select(ParameterDeclaration))})");
        writer.Open();
        var p = symbol.Parameters;
        if (p.Count == 5 && symbol.ReturnType is { IsVariant: true })
        {
            writer.HiddenLine($"return {R}Native.CallByName({p[0].EmitName}, (nint){p[1].EmitName}, (int){p[2].EmitName}, {p[3].EmitName});");
        }
        else if (p.Count == 6 && p[0].Type.IsVariant && !p[0].IsByVal && symbol.ReturnType is { IsNumeric: true } result)
        {
            writer.HiddenLine($"{R}ObjectRefs.Assign(ref {p[0].EmitName}, {R}Native.CallByName({p[1].EmitName}, (nint){p[2].EmitName}, (int){p[3].EmitName}, {p[4].EmitName}));");
            writer.HiddenLine($"return ({result.CSharpName}){R}Pointers.Address(ref {p[0].EmitName});");
        }
        else
        {
            Report(symbol.Syntax, "VBE7's rtcCallByName takes an object, a pointer to the member's name, a call type, a Variant array of arguments, and an LCID.");
        }

        writer.Close();
        writer.Hidden();
        writer.Line();
    }

    /// <summary>An element of a typed array as a reference into its native storage, typed for the callee: the element's own type, or byte for As Any.</summary>
    private string ElementReference(BoundElement element, string type)
    {
        var indices = string.Join(", ", element.Indices.Select(i => Convert(EmitExpression(i), VbaType.Long)));
        var array = EmitExpression(element.Array).Code;
        return $"{(element.Array.Type.IsVariant ? array + ".AsArray()" : array)}.ElementRef<{type}>([{indices}])";
    }

    /// <summary>VarPtr, StrPtr, and ObjPtr (ARCHITECTURE.md section 5, "Pointers"): the binder lets VarPtr through only for storage whose address is stable (Binder.Pointers.cs).</summary>
    private string PointerCode(BoundPointer pointer)
    {
        switch (pointer.Kind)
        {
            case PointerKind.String:
                return $"(long)({Convert(EmitExpression(pointer.Operand), VbaType.String)}).Pointer";
            case PointerKind.Object:
                return $"{R}Pointers.ObjPtr({Convert(EmitExpression(pointer.Operand), VbaType.Variant)})";
            default:
                if (pointer.Operand.Type.Kind == TypeKind.FixedString)
                {
                    // VarPtr of a fixed-length string is the address of a temporary BSTR copy of it, which dies with the statement (Memory golden).
                    var copy = NewTemporary();
                    before.Add($"var {copy} = {Convert(EmitExpression(pointer.Operand), VbaType.String)};");
                    return $"{R}Pointers.Address(ref {copy})";
                }

                if (pointer.Operand is BoundElement element && (element.Array.Type.IsArray || element.Array.Type.IsVariant))
                {
                    // An element's place in its array's storage, the array a Variant holds included (Memory golden).
                    var indices = string.Join(", ", element.Indices.Select(i => Convert(EmitExpression(i), VbaType.Long)));
                    var array = EmitExpression(element.Array).Code;
                    return $"{(element.Array.Type.IsVariant ? array + ".AsArray()" : array)}.ElementAddress([{indices}])";
                }

                if (!pointer.Operand.IsLValue)
                {
                    // A value that is not a variable: the address of a temporary holding it, a Boolean as its two bytes (Memory golden).
                    var value = EmitExpression(pointer.Operand);
                    var temporary = NewTemporary();
                    before.Add(ReferenceEquals(value.Type, VbaType.Boolean)
                        ? $"short {temporary} = {value.Code} ? (short)-1 : (short)0;"
                        : $"var {temporary} = {value.Code};");
                    return $"{R}Pointers.Address(ref {temporary})";
                }

                return $"{R}Pointers.Address(ref {SlotReference(pointer.Operand, pointer.Operand.Type) ?? EmitExpression(pointer.Operand).Code})";
        }
    }

    /// <summary>
    /// AddressOf (MS-VBAL 5.6.16.8): a delegate of the procedure's own shape, kept alive by a static
    /// field so the pointer stays valid, and its address as the machine word VBA yields.
    /// </summary>
    private string AddressOf(ProcedureSymbol target)
    {
        callbacks.Add(target);
        return $"((long){Interop}Marshal.GetFunctionPointerForDelegate({ThunkName(target)}))";
    }

    private static string ThunkName(ProcedureSymbol target) => "__thunk_" + target.Module.EmitName + "_" + target.EmitName.TrimStart('@');

    private void EmitCallbacks()
    {
        foreach (var target in callbacks.OrderBy(t => ThunkName(t), StringComparer.Ordinal))
        {
            var name = ThunkName(target);
            var parameters = string.Join(", ", target.Parameters.Select(ParameterDeclaration));
            var callee = ReferenceEquals(target.Module, Symbol) ? target.EmitName : $"global::{target.Module.EmitName}.{target.EmitName}";
            if (target.ReturnType is { IsVariableString: true })
            {
                // A String result is a BSTR the native caller owns: the thunk takes it from the statement it would have died with (ObjectRefs.HandOut; Declares golden).
                var arguments = string.Join(", ", target.Parameters.Select(p => (p.EmitByRef ? "ref " : string.Empty) + p.EmitName));
                writer.HiddenLine($"[{Interop}UnmanagedFunctionPointer({Interop}CallingConvention.StdCall)]");
                writer.HiddenLine($"private delegate nint {name}__signature({parameters});");
                writer.HiddenLine($"private static nint {name}__body({parameters}) => {R}ObjectRefs.HandOut({callee}({arguments}));");
                writer.HiddenLine($"private static readonly {name}__signature {name} = {name}__body;");
                writer.Line();
                continue;
            }

            if (target.ReturnType is { IsVariant: true })
            {
                // A Variant result travels as x64 returns a VARIANT: through a hidden pointer the native caller passes first, which the thunk fills and hands back (ObjectRefs.HandOut; Declares golden).
                var arguments = string.Join(", ", target.Parameters.Select(p => (p.EmitByRef ? "ref " : string.Empty) + p.EmitName));
                var hidden = parameters.Length == 0 ? "nint __result" : "nint __result, " + parameters;
                writer.HiddenLine($"[{Interop}UnmanagedFunctionPointer({Interop}CallingConvention.StdCall)]");
                writer.HiddenLine($"private delegate nint {name}__signature({hidden});");
                writer.HiddenLine($"private static nint {name}__body({hidden}) => {R}ObjectRefs.HandOut({callee}({arguments}), __result);");
                writer.HiddenLine($"private static readonly {name}__signature {name} = {name}__body;");
                writer.Line();
                continue;
            }

            writer.HiddenLine($"[{Interop}UnmanagedFunctionPointer({Interop}CallingConvention.StdCall)]");
            writer.HiddenLine($"private delegate {target.ReturnType?.CSharpName ?? "void"} {name}__signature({parameters});");
            writer.HiddenLine($"private static readonly {name}__signature {name} = {callee};");
            writer.Line();
        }
    }
}
