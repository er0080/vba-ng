using System.Globalization;

using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>Expression binding: names, calls and arguments, operators and their declared types, and constant folding (MS-VBAL 5.6).</summary>
public sealed partial class Binder
{
    private static readonly CultureInfo DateLiteralCulture = CultureInfo.GetCultureInfo("en-US");

    /// <summary>
    /// Resolves a simple name (MS-VBAL 5.6.10): the procedure's locals and parameters, then the
    /// module's members, then the public members of the other modules, then the VBA library.
    /// </summary>
    private Symbol? Lookup(string name)
    {
        if (procedure is not null && procedure.Locals.TryGetValue(name, out var local))
        {
            return local;
        }

        if (module.Members.TryGetValue(name, out var member))
        {
            return member;
        }

        // In a document module the members of its own object come next (MS-VBAL 5.6.10, the enclosing class step).
        if (module.Me is { Type.IsCom: true } me && libraries.Members(me.Type.ComInterface!, name).Any())
        {
            return new DocumentMemberSymbol(name, me);
        }

        foreach (var other in project.Modules)
        {
            if (ReferenceEquals(other, module) || !other.Members.TryGetValue(name, out var found) || !IsPublic(found))
            {
                continue;
            }

            // The variables and procedures of a class or document module belong to its instances, not to the
            // project (MS-VBAL 5.6.10); the members of a Public Enum it declares are project-wide constants.
            if (other.Kind == ModuleKind.Standard || found is EnumSymbol || (found is ConstantSymbol constantMember && other.Enums.Any(e => e.Members.Contains(constantMember))))
            {
                return found;
            }
        }

        if (StandardLibrary.TryGetConstant(name, out var constant))
        {
            return constant;
        }

        if (StandardLibrary.TryGetFunction(name, out var intrinsic))
        {
            return intrinsic;
        }

        // Referenced type libraries: enum and module constants, application object members (ARCHITECTURE.md section 6).
        return libraries.FindGlobal(name);
    }

    private static bool IsPublic(Symbol symbol) => symbol switch
    {
        VariableSymbol variable => variable.IsPublic,
        ConstantSymbol constant => constant.IsPublic,
        ProcedureSymbol procedure => procedure.IsPublic,
        PropertySymbol property => property.IsPublic,
        RecordSymbol record => record.IsPublic,
        EnumSymbol enumeration => enumeration.IsPublic,
        _ => false,
    };

    private BoundExpression BindExpression(ExpressionSyntax syntax)
    {
        switch (syntax)
        {
            case LiteralExpressionSyntax literal:
                return BindLiteral(literal);
            case ParenthesizedExpressionSyntax parenthesized:
                return new BoundParenthesized(parenthesized, BindExpression(parenthesized.Expression));
            case IdentifierNameSyntax name:
                return BindName(name);
            case UnaryExpressionSyntax unary:
                return BindUnary(unary);
            case BinaryExpressionSyntax binary:
                return BindBinary(binary);
            case IndexExpressionSyntax index:
                return BindIndex(index);
            case MemberAccessExpressionSyntax member:
                return BindMember(member, null);
            case NewExpressionSyntax created:
                {
                    var type = BindNamedType(created.Type);
                    if (!type.IsObject)
                    {
                        Report(DiagnosticIds.ObjectRequired, created.NewKeyword, "Expected: class name after New.");
                    }

                    return new BoundNewObject(created, type);
                }

            case DictionaryAccessExpressionSyntax dictionary:
                return BindDictionaryAccess(dictionary);
            case TypeOfExpressionSyntax typeOf:
                return BindTypeOf(typeOf);
            case AddressOfExpressionSyntax addressOf:
                return BindAddressOf(addressOf);
            default:
                Report(DiagnosticIds.NotSupported, syntax.FirstToken()!, $"{syntax.Kind} is not supported yet.");
                return Empty(syntax);
        }
    }

    private static BoundLiteral Empty(SyntaxNode syntax) => new(syntax, Variant.Empty, VbaType.Variant);

    /// <summary>Literal typing (MS-VBAL 3.3.2, 3.3.3, 5.6.5).</summary>
    private BoundLiteral BindLiteral(LiteralExpressionSyntax literal)
    {
        var token = literal.Token;
        switch (token.Kind)
        {
            case SyntaxKind.IntegerLiteralToken:
                {
                    var value = IntegerLiteral(token.Text);
                    return new BoundLiteral(literal, value, VbaType.FromVarType(value.Type));
                }

            case SyntaxKind.FloatLiteralToken:
                {
                    var value = LiteralValues.ParseFloat(token.Text);
                    Variant typed = token.Text[^1] switch
                    {
                        '!' => Variant.FromSingle((float)value),
                        '@' => Variant.FromCurrency(Currency.FromDouble(value)),
                        _ => Variant.FromDouble(value),
                    };
                    return new BoundLiteral(literal, typed, VbaType.FromVarType(typed.Type));
                }

            case SyntaxKind.StringLiteralToken:
                return new BoundLiteral(literal, Variant.FromString((string)token.Value!), VbaType.String);
            case SyntaxKind.DateLiteralToken:
                {
                    var inner = token.Text.Trim('#').Trim();
                    if (!DateText.TryParse(inner, DateLiteralCulture, out var date))
                    {
                        Report(DiagnosticIds.SyntaxError, token, "Invalid date literal.");
                    }

                    return new BoundLiteral(literal, Variant.FromDate(date), VbaType.Date);
                }

            case SyntaxKind.TrueKeyword:
                return new BoundLiteral(literal, Variant.True, VbaType.Boolean);
            case SyntaxKind.FalseKeyword:
                return new BoundLiteral(literal, Variant.False, VbaType.Boolean);
            case SyntaxKind.NullKeyword:
                return new BoundLiteral(literal, Variant.Null, VbaType.Variant);
            case SyntaxKind.EmptyKeyword:
                return new BoundLiteral(literal, Variant.Empty, VbaType.Variant);
            case SyntaxKind.NothingKeyword:
                return new BoundLiteral(literal, Variant.Nothing, VbaType.Object);
            default:
                Report(DiagnosticIds.SyntaxError, token, "Expected: expression.");
                return Empty(literal);
        }
    }

    /// <summary>Integer literal typing (MS-VBAL 3.3.2): the suffix, else the narrowest of Integer, Long, Double; hex and octal read the sign bit.</summary>
    private static Variant IntegerLiteral(string text)
    {
        var suffix = text[^1] is '%' or '&' or '^' ? text[^1] : '\0';
        var magnitude = LiteralValues.ParseInteger(text);
        if (magnitude is double wide)
        {
            return Variant.FromDouble(wide);
        }

        var value = (long)magnitude;
        var radix = text.StartsWith('&');
        switch (suffix)
        {
            case '%':
                return Variant.FromInt16(radix ? (short)(ushort)value : unchecked((short)value));
            case '&':
                return Variant.FromInt32(radix ? (int)(uint)value : unchecked((int)value));
            case '^':
                return Variant.FromInt64(value);
        }

        if (radix)
        {
            return (ulong)value switch
            {
                <= 0xFFFF => Variant.FromInt16((short)(ushort)value),
                <= 0xFFFFFFFF => Variant.FromInt32((int)(uint)value),
                _ => Variant.FromInt64(value),
            };
        }

        return value switch
        {
            <= short.MaxValue => Variant.FromInt16((short)value),
            <= int.MaxValue => Variant.FromInt32((int)value),
            _ => Variant.FromDouble(value),
        };
    }

    private BoundExpression BindName(IdentifierNameSyntax name)
    {
        var text = name.Name;
        if (text.Equals("Err", StringComparison.OrdinalIgnoreCase))
        {
            return new BoundErrObject(name);
        }

        if (text.Equals("Erl", StringComparison.OrdinalIgnoreCase))
        {
            procedure.UsesDispatch = true;
            return new BoundErl(name);
        }

        if (text.Equals("Me", StringComparison.OrdinalIgnoreCase))
        {
            if (module.Me is { } me)
            {
                return new BoundVariable(name, me);
            }

            Report(DiagnosticIds.NotSupported, name.Identifier, module.Kind == ModuleKind.Standard ? "Me is not valid in a standard module." : "Me is not valid in this module.");
            return Empty(name);
        }

        switch (Lookup(text))
        {
            case VariableSymbol variable:
                return new BoundVariable(name, variable);
            case ConstantSymbol constant:
                return new BoundConstant(name, constant);
            case ProcedureSymbol callee:
                return BindProcedureCall(callee, [], name, asStatement: false);
            case PropertySymbol property:
                return BindPropertyGet(property, [], name);
            case IntrinsicSymbol intrinsic:
                return BindIntrinsicCall(intrinsic, [], name, name.Identifier.TypeSuffix == '$');
            case ComConstantSymbol libraryConstant:
                return new BoundLiteral(name, libraryConstant.Value, libraryConstant.Type);
            case ComGlobalSymbol global:
                return BindComGlobal(global, [], name, isCallStatement: false);
            case DocumentMemberSymbol documentMember:
                return BindComMember(new BoundVariable(name, documentMember.Me), documentMember.Me.Type.ComInterface!, documentMember.Name, [], name, isCallStatement: false);
            case null when project.FindModule(text) is { Kind: ModuleKind.Class, Instance: { } instance }:
                // A class with VB_PredeclaredId = True: its name reaches the predeclared instance (MS-VBAL 5.2.4.1.1).
                return new BoundVariable(name, instance);
            case null when project.FindModule(text) is { Kind: ModuleKind.Document, Me: { } sheet }:
                // A document module's name is the object it is bound to (Reporter.ConnectTo SpecRunner in VBA-Web's specs).
                return new BoundVariable(name, sheet);
            case null when name.Identifier.Kind == SyntaxKind.ForeignNameToken && BindForeignName(name.Identifier, name, asTarget: false) is { } evaluated:
                return evaluated;
            case null:
                return new BoundVariable(name, ImplicitLocal(name.Identifier));
            default:
                Report(DiagnosticIds.TypeMismatch, name.Identifier, $"'{text}' is a type or module name; expected: expression.");
                return Empty(name);
        }
    }

    /// <summary>A call statement whose target is a simple name: Foo, Foo a, b, or Call Foo(a, b) (MS-VBAL 5.4.2.1).</summary>
    private BoundExpression BindNameCall(IdentifierNameSyntax name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool byValParenthesized)
    {
        switch (Lookup(name.Name))
        {
            case ProcedureSymbol callee:
                return BindProcedureCall(callee, arguments, syntax, asStatement: true, byValParenthesized);
            case VariableSymbol { Kind: VariableKind.Result } when procedure is { Symbol: { } self } && ReferenceEquals(procedure.Symbol.Result, Lookup(name.Name)):
                // The function's own name as a statement inside it calls it again (MS-VBAL 5.4.2.1); as an expression it is the result variable (VBA-Dictionary's Specs).
                return BindProcedureCall(self, arguments, syntax, asStatement: true, byValParenthesized);
            case PropertySymbol property:
                return BindPropertyGet(property, arguments, syntax);
            case IntrinsicSymbol intrinsic:
                return BindIntrinsicCall(intrinsic, arguments, syntax, name.Identifier.TypeSuffix == '$');
            case ComGlobalSymbol global:
                return BindComGlobal(global, arguments, syntax, isCallStatement: true);
            case DocumentMemberSymbol documentMember:
                return BindComMember(new BoundVariable(syntax, documentMember.Me), documentMember.Me.Type.ComInterface!, documentMember.Name, arguments, syntax, isCallStatement: true);
            default:
                Report(DiagnosticIds.ProcedureNotDefined, name.Identifier, $"Sub or Function not defined: '{name.Name}'.");
                return Empty(syntax);
        }
    }

    /// <summary>expression(arguments) (MS-VBAL 5.6.13): an array element, a call, or a default member, decided by what the expression names.</summary>
    private BoundExpression BindIndex(IndexExpressionSyntax index)
    {
        var arguments = index.Arguments.Arguments.ToList();
        switch (index.Expression)
        {
            case IdentifierNameSyntax name:
                {
                    var text = name.Name;
                    if (text.Equals("Array", StringComparison.OrdinalIgnoreCase) && Lookup(text) is not VariableSymbol and not ProcedureSymbol)
                    {
                        var items = arguments.Select(a => a.IsOmitted ? (BoundExpression)new BoundLiteral(index, Variant.Missing, VbaType.Variant) : Convert(BindExpression(a.Expression!), VbaType.Variant)).ToList();
                        return new BoundArrayFunction(index, items, module.Options.Base);
                    }

                    if ((text.Equals("Len", StringComparison.OrdinalIgnoreCase) || text.Equals("LenB", StringComparison.OrdinalIgnoreCase))
                        && arguments is [{ IsNamed: false, Expression: LiteralExpressionSyntax { Token.Kind: SyntaxKind.IntegerLiteralToken or SyntaxKind.FloatLiteralToken or SyntaxKind.TrueKeyword or SyntaxKind.FalseKeyword } }])
                    {
                        // The VBE rejects Len of a numeric or Boolean literal at compile time.
                        Report(DiagnosticIds.TypeMismatch, name.Identifier, "Type mismatch: Len needs a string or a variable.");
                        return Empty(index);
                    }

                    if ((text.Equals("Len", StringComparison.OrdinalIgnoreCase) || text.Equals("LenB", StringComparison.OrdinalIgnoreCase))
                        && arguments is [{ IsNamed: false, Expression: IdentifierNameSyntax recordName }]
                        && Lookup(recordName.Name) is VariableSymbol { IsArray: false } recordVariable && recordVariable.Type.IsRecord)
                    {
                        // Len of a user-defined type is the size of its file layout and LenB of its memory layout, as an Integer (FileSystem golden).
                        return new BoundRecordLength(index, new BoundVariable(recordName, recordVariable), text.Length == 4);
                    }

                    if ((text.Equals("Len", StringComparison.OrdinalIgnoreCase) || text.Equals("LenB", StringComparison.OrdinalIgnoreCase))
                        && arguments is [{ IsNamed: false, Expression: IndexExpressionSyntax { Expression: IdentifierNameSyntax elementArray } element }]
                        && Lookup(elementArray.Name) is VariableSymbol { IsArray: true, Type.ElementType.IsRecord: true })
                    {
                        // Len and LenB of an element of an array of records measure the record, as of a record variable (Memory golden).
                        return new BoundRecordLength(index, BindExpression(element), text.Length == 4);
                    }

                    if ((text.Equals("Len", StringComparison.OrdinalIgnoreCase) || text.Equals("LenB", StringComparison.OrdinalIgnoreCase))
                        && arguments is [{ IsNamed: false, Expression: MemberAccessExpressionSyntax recordMember }]
                        && RecordMemberType(recordMember) is { IsRecord: true })
                    {
                        // Len and LenB of a member that is itself a user-defined type measure that type, as of a record variable (Memory golden).
                        return new BoundRecordLength(index, BindExpression(recordMember), text.Length == 4);
                    }

                    if ((text.Equals("Len", StringComparison.OrdinalIgnoreCase) || text.Equals("LenB", StringComparison.OrdinalIgnoreCase))
                        && arguments is [{ IsNamed: false, Expression: IdentifierNameSyntax argumentName }]
                        && Lookup(argumentName.Name) is VariableSymbol typed
                        && !typed.IsArray && typed.Type.Kind == TypeKind.Builtin && !typed.Type.IsVariant && !typed.Type.IsString && !typed.Type.IsObject)
                    {
                        // Len of a declared non-String variable is the size of its type in bytes (Strings golden).
                        return new BoundLenOfDeclared(index, typed.Type, text.Length == 4);
                    }

                    if (text.Equals("LenB", StringComparison.OrdinalIgnoreCase)
                        && arguments is [{ IsNamed: false, Expression: IdentifierNameSyntax fixedName }]
                        && Lookup(fixedName.Name) is VariableSymbol { IsArray: false, Type.Kind: TypeKind.FixedString } fixedString)
                    {
                        // LenB of a fixed-length String variable is its size, an Integer, where Len of it is a Long (Memory and Strings goldens; docs/vba-quirks.md).
                        return new BoundLenOfDeclared(index, fixedString.Type, bytes: true);
                    }

                    if (text.Equals("LenB", StringComparison.OrdinalIgnoreCase)
                        && arguments is [{ IsNamed: false, Expression: IndexExpressionSyntax { Expression: IdentifierNameSyntax fixedArrayName } fixedElement }]
                        && Lookup(fixedArrayName.Name) is VariableSymbol { IsArray: true, Type.ElementType: { Kind: TypeKind.FixedString } fixedElementType })
                    {
                        // LenB of an element of a fixed-length String array is an Integer too (Arrays golden); the element is still read, so a subscript out of range raises 9.
                        return new BoundLenOfDeclared(index, fixedElementType, bytes: true) { Operand = BindExpression(fixedElement) };
                    }

                    if (PointerKindOf(text) is { } pointerKind && arguments is [{ IsNamed: false, Expression: { } pointerOperand }] && Lookup(text) is IntrinsicSymbol)
                    {
                        return BindPointer(index, name.Identifier, pointerKind, pointerOperand);
                    }

                    switch (Lookup(text))
                    {
                        case VariableSymbol { Kind: VariableKind.Result } when procedure.Symbol.Parameters.Count > 0 || !procedure.Symbol.ReturnType!.IsArray:
                            // Inside a Function, its own name with arguments is a recursive call (MS-VBAL 5.3.1.8).
                            return BindProcedureCall(procedure.Symbol, arguments, index, asStatement: false);
                        case VariableSymbol { IsArray: true } array:
                            return BindElement(new BoundVariable(name, array), arguments, index);
                        case VariableSymbol { Type.IsCom: true } com:
                            return BindComDefault(new BoundVariable(name, com), arguments, index);
                        case VariableSymbol { Type.ProjectClass: { } instanceClass } instanceVariable:
                            return BindClassDefault(new BoundVariable(name, instanceVariable), instanceClass, arguments, index);
                        case VariableSymbol variable when variable.Type.IsVariant || variable.Type.IsObject:
                            return BindLateElement(new BoundVariable(name, variable), arguments, index);
                        case VariableSymbol:
                            Report(DiagnosticIds.ExpectedArray, name.Identifier, $"Expected array: '{text}'.");
                            return Empty(index);
                        case ProcedureSymbol callee:
                            return BindCallOrIndex(callee, arguments, index, isCallStatement: false);
                        case PropertySymbol property:
                            return BindPropertyGet(property, arguments, index);
                        case IntrinsicSymbol intrinsic:
                            return BindIntrinsicCall(intrinsic, arguments, index, name.Identifier.TypeSuffix == '$');
                        case ComGlobalSymbol global:
                            return BindComGlobal(global, arguments, index, isCallStatement: false);
                        case DocumentMemberSymbol documentMember:
                            return BindComMember(new BoundVariable(name, documentMember.Me), documentMember.Me.Type.ComInterface!, documentMember.Name, arguments, index, isCallStatement: false);
                        case ConstantSymbol or ComConstantSymbol:
                            Report(DiagnosticIds.ExpectedArray, name.Identifier, $"Expected array: '{text}'.");
                            return Empty(index);
                        case null:
                            Report(DiagnosticIds.ProcedureNotDefined, name.Identifier, $"Sub or Function not defined: '{text}'.");
                            return Empty(index);
                        default:
                            Report(DiagnosticIds.TypeMismatch, name.Identifier, $"'{text}' is a type or module name; expected: expression.");
                            return Empty(index);
                    }
                }

            case MemberAccessExpressionSyntax member:
                return BindMember(member, arguments);
            default:
                {
                    var target = BindExpression(index.Expression);
                    return target.Type.IsCom ? BindComDefault(target, arguments, index) : target.Type.IsArray ? BindElement(target, arguments, index) : BindLateElement(target, arguments, index);
                }
        }
    }

    private BoundExpression BindElement(BoundExpression array, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        var indices = new List<BoundExpression>();
        foreach (var argument in arguments)
        {
            if (argument.IsNamed || argument.IsOmitted)
            {
                Report(DiagnosticIds.SyntaxError, argument.FirstToken() ?? syntax.FirstToken()!, "Expected: array index.");
                continue;
            }

            indices.Add(Convert(BindExpression(argument.Expression!), VbaType.Long));
        }

        if (array is BoundVariable { Variable.Bounds: { } bounds } && bounds.Length != indices.Count)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, syntax.FirstToken()!, "Wrong number of dimensions.");
        }

        if (indices.Count == 0)
        {
            // a() with no index names the array itself (a call-style reference to a dynamic array).
            return array;
        }

        return new BoundElement(syntax, array, indices, array.Type.IndexedType);
    }

    private BoundElement BindLateElement(BoundExpression target, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        var indices = new List<BoundExpression>();
        foreach (var argument in arguments)
        {
            if (argument.IsNamed)
            {
                Report(DiagnosticIds.NamedArgumentNotFound, argument.Name!, "Named arguments need an early-bound target.");
                continue;
            }

            indices.Add(argument.IsOmitted ? new BoundLiteral(syntax, Variant.Missing, VbaType.Variant) : Convert(BindExpression(argument.Expression!), VbaType.Variant));
        }

        return new BoundElement(syntax, target, indices, VbaType.Variant);
    }

    /// <summary>expression.name [arguments] (MS-VBAL 5.6.12): module members, enum members, record fields, object members, or With members.</summary>
    private BoundExpression BindMember(MemberAccessExpressionSyntax member, IReadOnlyList<ArgumentSyntax>? arguments, bool isCallStatement = false, bool asTarget = false)
    {
        arguments ??= [];
        var name = member.Name.Name;
        if (member.Expression is null)
        {
            if (procedure.WithTargets.Count == 0)
            {
                Report(DiagnosticIds.InvalidStatementPlacement, member.DotToken, "Invalid or unqualified reference: a leading '.' needs a With block.");
                return Empty(member);
            }

            return BindMemberOn(new BoundVariable(member, procedure.WithTargets.Peek()), name, arguments, member, isCallStatement);
        }

        if (member.Name.Identifier.Kind == SyntaxKind.ForeignNameToken && member.Expression is { } receiverSyntax)
        {
            // obj.[name]: the member of that name, or the receiver's Evaluate of the bracketed text (MS-VBAL 3.3.5.3).
            return BindForeignMember(BindExpression(receiverSyntax), member.Name.Identifier, arguments, member, isCallStatement, asTarget);
        }

        if (member.Expression is MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax { Name: var vbaLibrary }, Name.Name: var vbaModule }
            && vbaLibrary.Equals("VBA", StringComparison.OrdinalIgnoreCase)
            && StandardLibrary.IsLibraryModule(vbaModule)
            && Lookup("VBA") is null
            && project.FindModule("VBA") is null)
        {
            // VBA.Math.Abs and its kin: the library's own module qualifies the function (MS-VBAL 6.1).
            if (StandardLibrary.TryGetFunction(name, out var qualified))
            {
                return BindIntrinsicCall(qualified, arguments, member, member.Name.Identifier.TypeSuffix == '$');
            }

            if (StandardLibrary.TryGetConstant(name, out var qualifiedConstant))
            {
                return new BoundConstant(member, qualifiedConstant);
            }

            Report(DiagnosticIds.ProcedureNotDefined, member.Name.Identifier, $"Sub or Function not defined: 'VBA.{vbaModule}.{name}'.");
            return Empty(member);
        }

        if (member.Expression is MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax { Name: var library }, Name.Name: var libraryModule }
            && library.Equals(ProjectManifest.VbangLibrary, StringComparison.OrdinalIgnoreCase)
            && libraryModule.Equals("Assert", StringComparison.OrdinalIgnoreCase)
            && vbangReference
            && Lookup(library) is null
            && project.FindModule(library) is null)
        {
            // vbang.Assert.X: the library-qualified form, as VBA.Strings.Left names a VBA function.
            return BindAssertMember(name, arguments, member, isCallStatement);
        }

        if (member.Expression is MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax { Name: var moduleName }, Name.Name: var enumName }
            && project.FindModule(moduleName) is { } enumModule
            && (Lookup(moduleName) is null || DeclaredElsewhere(Lookup(moduleName)!))
            && enumModule.Members.TryGetValue(enumName, out var enumSymbol)
            && enumSymbol is EnumSymbol moduleEnum
            && (ReferenceEquals(enumModule, module) || IsPublic(moduleEnum)))
        {
            // Helpers.Colors.Red: a member reached through its module and its enum (Scope golden).
            var item = moduleEnum.Members.FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (item is null)
            {
                Report(DiagnosticIds.VariableNotDefined, member.Name.Identifier, $"Variable not defined: '{moduleName}.{enumName}.{name}'.");
                return Empty(member);
            }

            return new BoundConstant(member, item);
        }

        if (member.Expression is IdentifierNameSyntax qualifier)
        {
            var text = qualifier.Name;
            if (text.Equals("Err", StringComparison.OrdinalIgnoreCase) && Lookup(text) is not VariableSymbol)
            {
                return BindErrMember(new BoundErrObject(qualifier), name, arguments, member);
            }

            if (text.Equals("Debug", StringComparison.OrdinalIgnoreCase) && name.Equals("Assert", StringComparison.OrdinalIgnoreCase))
            {
                StandardLibrary.TryGetFunction("Debug.Assert", out var assert);
                return BindLibrarySub(assert, arguments, member, isCallStatement);
            }

            if (text.Equals("Assert", StringComparison.OrdinalIgnoreCase) && vbangReference && Lookup(text) is null && project.FindModule(text) is null)
            {
                // The Assert module of test procedures (ARCHITECTURE.md section 8) exists only when the manifest
                // references the vbang library, and a project's own Assert, a variable or a module, still wins.
                return BindAssertMember(name, arguments, member, isCallStatement);
            }

            if (StandardLibrary.IsLibraryModule(text) && Lookup(text) is null && project.FindModule(text) is null)
            {
                // Strings.Join, Math.Abs: the VBA library's module qualifies its function without the VBA prefix (MS-VBAL 6.1).
                if (StandardLibrary.TryGetFunction(name, out var moduleFunction))
                {
                    return BindIntrinsicCall(moduleFunction, arguments, member, member.Name.Identifier.TypeSuffix == '$');
                }

                if (StandardLibrary.TryGetConstant(name, out var moduleConstant))
                {
                    return new BoundConstant(member, moduleConstant);
                }

                Report(DiagnosticIds.ProcedureNotDefined, member.Name.Identifier, $"Sub or Function not defined: '{text}.{name}'.");
                return Empty(member);
            }

            if (text.Equals("VBA", StringComparison.OrdinalIgnoreCase) && Lookup(text) is null)
            {
                if (StandardLibrary.IsLibraryModule(name))
                {
                    // VBA.Math on its own names a module of the library; only VBA.Math.Abs means anything.
                    Report(DiagnosticIds.TypeMismatch, member.Name.Identifier, $"'VBA.{name}' is a module of the VBA library; expected: one of its functions.");
                    return Empty(member);
                }

                if (StandardLibrary.TryGetConstant(name, out var constant))
                {
                    return new BoundConstant(member, constant);
                }

                if (StandardLibrary.TryGetFunction(name, out var function))
                {
                    return BindIntrinsicCall(function, arguments, member, member.Name.Identifier.TypeSuffix == '$');
                }

                Report(DiagnosticIds.ProcedureNotDefined, member.Name.Identifier, $"Sub or Function not defined: 'VBA.{name}'.");
                return Empty(member);
            }

            var symbol = Lookup(text);
            if ((symbol is null || symbol is LibraryGlobal) && project.FindModule(text) is null && BindLibraryConstant(text, name, member) is { } libraryConstant)
            {
                // XlDirection.xlUp, Excel.xlUp: a constant reached through its enum or library name.
                return libraryConstant;
            }

            if (project.FindModule(text) is { } target && (symbol is null || DeclaredElsewhere(symbol)))
            {
                // A module name qualifies its members even when another module exports a member of the
                // same name (Sub Main in module Main); only this module's own names and locals shadow it.
                return BindModuleMember(target, name, arguments, member, isCallStatement);
            }

            if (symbol is EnumSymbol enumeration)
            {
                var item = enumeration.Members.FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (item is null)
                {
                    Report(DiagnosticIds.VariableNotDefined, member.Name.Identifier, $"Variable not defined: '{text}.{name}'.");
                    return Empty(member);
                }

                return new BoundConstant(member, item);
            }
        }

        var bound = BindExpression(member.Expression);
        return BindMemberOn(bound, name, arguments, member, isCallStatement);
    }

    private BoundExpression BindModuleMember(ModuleSymbol target, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        if (target.Kind == ModuleKind.Class && !ReferenceEquals(target, module))
        {
            if (target.Instance is null)
            {
                Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Variable not defined: '{target.Name}' is a class; use New or a variable of the type.");
                return Empty(syntax);
            }

            var instance = new BoundVariable(syntax, target.Instance);
            return ReferenceEquals(syntax, assignmentTarget)
                ? BindClassTarget(instance, target, name, arguments, syntax, assignmentIsSet)
                : BindClassMember(instance, target, name, arguments, syntax, isCallStatement);
        }

        if (!target.Members.TryGetValue(name, out var symbol) || (!ReferenceEquals(target, module) && !IsPublic(symbol)))
        {
            // A document module's name reaches the object it is bound to, as its predeclared instance does in VBA.
            if (target is { Kind: ModuleKind.Document, Me: { } bound })
            {
                return bound.Type.IsCom
                    ? BindComMember(new BoundVariable(syntax, bound), bound.Type.ComInterface!, name, arguments, syntax, isCallStatement)
                    : BindMemberOn(new BoundVariable(syntax, bound), name, arguments, syntax, isCallStatement);
            }

            Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Variable not defined: '{target.Name}.{name}'.");
            return Empty(syntax);
        }

        switch (symbol)
        {
            case VariableSymbol { IsArray: true } array when arguments.Count > 0:
                return BindElement(new BoundVariable(syntax, array), arguments, syntax);
            case VariableSymbol variable:
                return arguments.Count == 0 ? new BoundVariable(syntax, variable) : variable.Type.IsCom ? BindComDefault(new BoundVariable(syntax, variable), arguments, syntax) : BindLateElement(new BoundVariable(syntax, variable), arguments, syntax);
            case ConstantSymbol constant:
                return new BoundConstant(syntax, constant);
            case ProcedureSymbol callee:
                return BindCallOrIndex(callee, arguments, syntax, isCallStatement);
            case PropertySymbol property:
                // Runner.Col = 1: a standard module's Property Let through the module's name (VBA-Dictionary's specs).
                return ReferenceEquals(syntax, assignmentTarget)
                    ? BindPropertyTarget(property, syntax, arguments, assignmentIsSet)
                    : BindPropertyGet(property, arguments, syntax);
            default:
                Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, $"'{target.Name}.{name}' is a type; expected: expression.");
                return Empty(syntax);
        }
    }

    private BoundExpression BindMemberOn(BoundExpression target, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        if (target.Type.ProjectClass is { } projectClass)
        {
            return ReferenceEquals(syntax, assignmentTarget)
                ? BindClassTarget(target, projectClass, name, arguments, syntax, assignmentIsSet)
                : BindClassMember(target, projectClass, name, arguments, syntax, isCallStatement);
        }

        if (target.Type.IsCom)
        {
            return BindComMember(target, target.Type.ComInterface!, name, arguments, syntax, isCallStatement);
        }

        if (target.Type.IsRecord)
        {
            var field = target.Type.Record!.FindField(name);
            if (field is null)
            {
                Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
                return Empty(syntax);
            }

            var access = new BoundField(syntax, target, field);
            if (arguments.Count == 0)
            {
                return access;
            }

            return field.Type.IsArray ? BindElement(access, arguments, syntax) : BindLateElement(access, arguments, syntax);
        }

        if (ReferenceEquals(target.Type, VbaType.Collection))
        {
            return BindCollectionMember(target, name, arguments, syntax);
        }

        if (ReferenceEquals(target.Type, VbaType.ErrObject))
        {
            return BindErrMember(target, name, arguments, syntax);
        }

        if (target.Type.IsVariant || target.Type.IsGenericObject)
        {
            return BindLateMember(target, name, arguments, syntax);
        }

        Report(DiagnosticIds.ObjectRequired, syntax.FirstToken()!, "Object required.");
        return Empty(syntax);
    }

    /// <summary>
    /// Late binding (MS-VBAL 5.6.12): the member is looked up at run time, and so are the names of
    /// named arguments (5.6.13.1), which travel after the positional ones with their names.
    /// </summary>
    private BoundMember BindLateMember(BoundExpression target, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        var values = new List<BoundExpression>();
        var named = new List<(string Name, BoundExpression Value)>();
        foreach (var argument in arguments)
        {
            if (argument.IsNamed)
            {
                named.Add((argument.Name!.NameValue, Convert(BindExpression(argument.Expression!), VbaType.Variant)));
                continue;
            }

            if (named.Count > 0)
            {
                Report(DiagnosticIds.NamedArgumentNotFound, argument.Expression?.FirstToken() ?? syntax.FirstToken()!, "Expected: named parameter.");
            }

            values.Add(argument.IsOmitted ? new BoundLiteral(syntax, Variant.Missing, VbaType.Variant) : Convert(BindExpression(argument.Expression!), VbaType.Variant));
        }

        values.AddRange(named.Select(n => n.Value));
        return new BoundMember(syntax, target, name, values, VbaType.Variant) { NamedArguments = named.Select(n => n.Name).ToList() };
    }

    /// <summary>
    /// target!name (MS-VBAL 5.6.13.2 dictionary-access): the target's default member with the name
    /// as a string argument. The member is reached at run time, as VBA reaches it, so the same
    /// form works for a project class, a COM object, and a Variant (Classes golden).
    /// </summary>
    private BoundExpression BindDictionaryAccess(DictionaryAccessExpressionSyntax dictionary)
    {
        BoundExpression target;
        if (dictionary.Expression is null)
        {
            if (procedure.WithTargets.Count == 0)
            {
                Report(DiagnosticIds.InvalidStatementPlacement, dictionary.BangToken, "Invalid or unqualified reference: a leading '!' needs a With block.");
                return Empty(dictionary);
            }

            target = new BoundVariable(dictionary, procedure.WithTargets.Peek());
        }
        else
        {
            target = BindExpression(dictionary.Expression);
        }

        var name = new BoundLiteral(dictionary.Name, Variant.FromString(dictionary.Name.Name), VbaType.String);
        if (ReferenceEquals(target.Type, VbaType.Collection))
        {
            return new BoundMember(dictionary, target, "Item", [Convert(name, VbaType.Variant)], VbaType.Variant);
        }

        if (!target.Type.IsObject && !target.Type.IsVariant)
        {
            Report(DiagnosticIds.ObjectRequired, dictionary.BangToken, "Object required.");
            return Empty(dictionary);
        }

        return new BoundElement(dictionary, target, [Convert(name, VbaType.Variant)], VbaType.Variant);
    }

    /// <summary>The members of Collection (MS-VBAL 6.1.3.2): Add, Count, Item, Remove.</summary>
    private BoundExpression BindCollectionMember(BoundExpression target, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        switch (name.ToUpperInvariant())
        {
            case "COUNT":
                return new BoundMember(syntax, target, "Count", NamedArguments(arguments, [], syntax), VbaType.Long);
            case "ITEM":
                return new BoundMember(syntax, target, "Item", NamedArguments(arguments, ["Index"], syntax, required: 1), VbaType.Variant);
            case "ADD":
                return new BoundMember(syntax, target, "Add", NamedArguments(arguments, ["Item", "Key", "Before", "After"], syntax, required: 1), VbaType.Variant);
            case "REMOVE":
                return new BoundMember(syntax, target, "Remove", NamedArguments(arguments, ["Index"], syntax, required: 1), VbaType.Variant);
            case "_NEWENUM":
                // [_NewEnum]: the enumerator a class hands to For Each from its own NewEnum (MS-VBAL 5.6.13.2, dispid -4).
                return new BoundMember(syntax, target, "NewEnum", NamedArguments(arguments, [], syntax), VbaType.Object);
            default:
                Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
                return Empty(syntax);
        }
    }

    /// <summary>The Err object's members (MS-VBAL 6.1.3.1).</summary>
    private BoundExpression BindErrMember(BoundExpression target, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        switch (name.ToUpperInvariant())
        {
            case "NUMBER":
            case "HELPCONTEXT":
            case "LASTDLLERROR":
                return new BoundMember(syntax, target, Capitalize(name), NamedArguments(arguments, [], syntax), VbaType.Long);
            case "DESCRIPTION":
            case "SOURCE":
            case "HELPFILE":
                return new BoundMember(syntax, target, Capitalize(name), NamedArguments(arguments, [], syntax), VbaType.String);
            case "RAISE":
                return new BoundMember(syntax, target, "Raise", NamedArguments(arguments, ["Number", "Source", "Description", "HelpFile", "HelpContext"], syntax, required: 1), VbaType.Variant);
            case "CLEAR":
                return new BoundMember(syntax, target, "Clear", NamedArguments(arguments, [], syntax), VbaType.Variant);
            default:
                Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
                return Empty(syntax);
        }
    }

    private static string Capitalize(string name) => name.ToUpperInvariant() switch
    {
        "NUMBER" => "Number",
        "DESCRIPTION" => "Description",
        "SOURCE" => "Source",
        "HELPFILE" => "HelpFile",
        "HELPCONTEXT" => "HelpContext",
        "LASTDLLERROR" => "LastDllError",
        _ => name,
    };

    /// <summary>Matches positional and named arguments to a fixed parameter list, filling gaps with Missing; every value is a Variant.</summary>
    private List<BoundExpression> NamedArguments(IReadOnlyList<ArgumentSyntax> arguments, string[] parameterNames, SyntaxNode syntax, int required = 0)
    {
        var slots = new BoundExpression?[parameterNames.Length];
        var position = 0;
        foreach (var argument in arguments)
        {
            int slot;
            if (argument.IsNamed)
            {
                slot = Array.FindIndex(parameterNames, p => p.Equals(argument.Name!.NameValue, StringComparison.OrdinalIgnoreCase));
                if (slot < 0)
                {
                    Report(DiagnosticIds.NamedArgumentNotFound, argument.Name!, $"Named argument not found: '{argument.Name!.NameValue}'.");
                    continue;
                }
            }
            else
            {
                slot = position++;
                if (slot >= parameterNames.Length)
                {
                    Report(DiagnosticIds.WrongNumberOfArguments, argument.FirstToken() ?? syntax.FirstToken()!, "Wrong number of arguments or invalid property assignment.");
                    continue;
                }
            }

            slots[slot] = argument.IsOmitted ? null : Convert(BindExpression(argument.Expression!), VbaType.Variant);
        }

        for (var i = 0; i < required; i++)
        {
            if (slots[i] is null)
            {
                Report(DiagnosticIds.ArgumentNotOptional, syntax.FirstToken()!, $"Argument not optional: '{parameterNames[i]}'.");
            }
        }

        return slots.Select(s => s ?? new BoundLiteral(syntax, Variant.Missing, VbaType.Variant)).ToList();
    }

    private BoundExpression BindPropertyGet(PropertySymbol property, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, BoundExpression? receiver = null)
    {
        if (property.Get is null)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, syntax.FirstToken()!, $"Property '{property.Name}' has no Get procedure.");
            return Empty(syntax);
        }

        return BindCallOrIndex(property.Get, arguments, syntax, isCallStatement: false, receiver);
    }

    /// <summary>
    /// A call, or arguments after a parameterless function or property that index its result
    /// (MS-VBAL 5.6.13.1): <c>Values("a")("b")</c>, <c>acc.children(i)</c>. The result's default
    /// member takes them, late when the result is an Object or a Variant (Classes golden).
    /// </summary>
    private BoundExpression BindCallOrIndex(ProcedureSymbol callee, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement, BoundExpression? receiver = null)
    {
        if (arguments.Count == 0 || callee.Parameters.Count > 0 || !callee.ReturnsValue || callee.ReturnType is not { } result || !(result.IsObject || result.IsVariant) || arguments.Any(a => a.IsNamed))
        {
            return BindProcedureCall(callee, arguments, syntax, isCallStatement, receiver: receiver);
        }

        var value = BindProcedureCall(callee, [], syntax, asStatement: false, receiver: receiver);
        return IndexValue(value, arguments, syntax);
    }

    /// <summary>Arguments applied to an object value: its default member, by dispid for a library type, by name at run time otherwise.</summary>
    private BoundExpression IndexValue(BoundExpression value, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        if (value.Type.ProjectClass is { DefaultMember: { } defaultMember } projectClass)
        {
            return BindClassMember(value, projectClass, defaultMember.Name, arguments, syntax, isCallStatement: false);
        }

        return value.Type.IsCom ? BindComDefault(value, arguments, syntax) : BindLateElement(value, arguments, syntax);
    }

    /// <summary>
    /// Binds a call to a project procedure (MS-VBAL 5.6.13.1): positional then named arguments,
    /// Optional defaults, ParamArray, and how each argument is passed.
    /// </summary>
    private BoundCall BindProcedureCall(ProcedureSymbol callee, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool asStatement, bool byValParenthesized = false, bool valueParameterPending = false, BoundExpression? receiver = null)
    {
        if (!asStatement && !callee.ReturnsValue)
        {
            Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, $"Expected Function or variable: '{callee.Name}' is a Sub.");
        }

        var parameters = callee.Parameters;
        var paramArray = callee.ParamArray;
        var fixedCount = parameters.Count - (paramArray is null ? 0 : 1) - (valueParameterPending ? 1 : 0);
        var slots = new ArgumentSyntax?[Math.Max(fixedCount, 0)];
        var extra = new List<ArgumentSyntax>();
        var position = 0;
        var namedSeen = false;
        foreach (var argument in arguments)
        {
            if (argument.IsNamed)
            {
                namedSeen = true;
                var index = parameters.FindIndex(p => p.Name.Equals(argument.Name!.NameValue, StringComparison.OrdinalIgnoreCase));
                if (index < 0 || index >= fixedCount)
                {
                    Report(DiagnosticIds.NamedArgumentNotFound, argument.Name!, $"Named argument not found: '{argument.Name!.NameValue}'.");
                    continue;
                }

                if (slots[index] is not null)
                {
                    Report(DiagnosticIds.WrongNumberOfArguments, argument.Name!, $"Named argument already specified: '{argument.Name!.NameValue}'.");
                    continue;
                }

                slots[index] = argument;
            }
            else if (namedSeen)
            {
                Report(DiagnosticIds.WrongNumberOfArguments, argument.FirstToken() ?? syntax.FirstToken()!, "Expected: named parameter.");
            }
            else if (position < fixedCount)
            {
                slots[position++] = argument;
            }
            else if (paramArray is not null)
            {
                extra.Add(argument);
            }
            else
            {
                Report(DiagnosticIds.WrongNumberOfArguments, argument.FirstToken() ?? syntax.FirstToken()!, $"Wrong number of arguments or invalid property assignment: '{callee.Name}'.");
            }
        }

        var bound = new List<BoundArgument>();
        for (var i = 0; i < fixedCount; i++)
        {
            var parameter = parameters[i];
            var argument = slots[i];
            if (argument is null || argument.IsOmitted)
            {
                if (!parameter.IsOptional)
                {
                    Report(DiagnosticIds.ArgumentNotOptional, syntax.FirstToken()!, $"Argument not optional: '{parameter.Name}'.");
                }

                bound.Add(new BoundArgument(parameter, null, ArgumentMode.ByVal));
                continue;
            }

            bound.Add(BindArgument(parameter, argument.Expression!, byValParenthesized || argument.ByValKeyword is not null, argument.ByValKeyword is not null));
        }

        var rest = new List<BoundExpression>();
        foreach (var argument in extra)
        {
            rest.Add(argument.IsOmitted ? new BoundLiteral(syntax, Variant.Missing, VbaType.Variant) : Convert(BindExpression(argument.Expression!), VbaType.Variant));
        }

        return new BoundCall(syntax, callee, bound, rest, callee.ReturnType ?? VbaType.Variant) { Receiver = receiver };
    }

    /// <summary>
    /// How one argument reaches its parameter (MS-VBAL 5.6.13.1): ByVal parameters, parenthesized
    /// arguments, and non-variables pass a coerced value; a variable of the parameter's type is
    /// passed by reference; a Variant variable, or an array element, goes through a temporary that
    /// is copied back; a typed variable of another type is a compile error.
    /// </summary>
    private BoundArgument BindArgument(ParameterSymbol parameter, ExpressionSyntax expression, bool forceByVal, bool explicitByVal)
    {
        var value = BindExpression(expression);
        if (parameter.IsExternal && !parameter.IsByVal && !explicitByVal && value is BoundElement element)
        {
            // VBA hands the DLL the element's address, and a Fortran-convention callee reads and writes the elements after it
            // (ROADMAP.md WP0). Since M7 C4 the elements lie in the array's native storage, so the element itself travels.
            if ((element.Array.Type.IsArray && (IsNativeElement(element.Type) || element.Type.IsRecord) && (ReferenceEquals(parameter.Type, VbaType.Any) || element.Type.SameAs(parameter.Type)))
                || (element.Array.Type.IsVariant && parameter.Type.IsVariant))
            {
                return new BoundArgument(parameter, element, ArgumentMode.ByRef);
            }

            Report(DiagnosticIds.NotSupported, expression.FirstToken()!, $"This array element passed ByRef to a Declare needs an address of its own, which this array does not keep its elements in; '{parameter.Name}' would receive a copy.");
            return new BoundArgument(parameter, Convert(value, ReferenceEquals(parameter.Type, VbaType.Any) ? VbaType.LongLong : parameter.Type), ArgumentMode.ByVal);
        }

        if (ReferenceEquals(parameter.Type, VbaType.Any))
        {
            // As Any takes the argument as it is: by value as a machine word, by reference as its own storage; ByVal at the
            // call hands a ByRef parameter the value itself as the address (MS-VBAL 5.2.3.5; Memory golden: MoveMemory x, ByVal p, n).
            return parameter.IsByVal || explicitByVal
                ? new BoundArgument(parameter, Convert(value, VbaType.LongLong), ArgumentMode.ByVal)
                : new BoundArgument(parameter, value, ArgumentMode.ByRef);
        }

        if (parameter.IsByVal && parameter.EmitByRef && parameter.Type.IsString)
        {
            // A Declare's ByVal String is a buffer the callee may write into, and the caller sees it (Declares golden).
            return new BoundArgument(parameter, Convert(value, parameter.Type), value.IsLValue ? ArgumentMode.ByRefCopyBack : ArgumentMode.ByVal);
        }

        // A class's variable reached through an object is its property, not a variable: a ByRef parameter of any type takes a copy
        // the callee's change does not reach (Classes golden; VBA-Web's TwitterSheet passes a WebResponse's StatusCode to an Integer).
        var classVariable = value is BoundField { Record.Type.IsRecord: false };
        if (parameter.IsByVal || forceByVal || value is BoundParenthesized || !value.IsLValue || classVariable)
        {
            if (!parameter.IsByVal && (parameter.Type.IsArray || parameter.Type.IsRecord))
            {
                if (value.Type.SameAs(parameter.Type))
                {
                    // A call result or expression of the parameter's type reaches a ByRef array or record parameter
                    // as a temporary the callee may change (MS-VBAL 5.6.13.1; Procedures golden).
                    return new BoundArgument(parameter, value, ArgumentMode.ByVal);
                }

                Report(DiagnosticIds.ByRefTypeMismatch, expression.FirstToken()!, "Type mismatch: array or user-defined type expected.");
            }

            return new BoundArgument(parameter, Convert(value, parameter.Type), ArgumentMode.ByVal);
        }

        if (parameter.Type.IsArray && ReferenceEquals(parameter.Type.ElementType, VbaType.Any) && value.Type.IsArray && value is BoundVariable or BoundField)
        {
            // An array parameter of Any type (VBE7's VarPtr declared as VarPtrArray) takes any array variable as itself.
            return new BoundArgument(parameter, value, ArgumentMode.ByRef);
        }

        if (value.Type.SameAs(parameter.Type) || (value.Type.IsEnum && ReferenceEquals(parameter.Type, VbaType.Long)) || (parameter.Type.IsEnum && ReferenceEquals(value.Type, VbaType.Long)))
        {
            return new BoundArgument(parameter, value, value is BoundVariable or BoundField ? ArgumentMode.ByRef : ArgumentMode.ByRefCopyBack);
        }

        if (value.Type.IsArray && parameter.Type.IsVariant)
        {
            // An array reaches a ByRef Variant parameter as itself, and a ReDim inside the callee reaches the caller (Procedures golden).
            return new BoundArgument(parameter, value, ArgumentMode.ByRefCopyBack);
        }

        if (value.Type.IsObject && parameter.Type.IsObject)
        {
            // An object of any type reaches a ByRef parameter of another object type, a class instance to As Object and back (Classes golden; verified in Excel 2026-09-09).
            return new BoundArgument(parameter, value, ArgumentMode.ByRefCopyBack);
        }

        if (value.Type.IsArray || parameter.Type.IsArray || value.Type.IsRecord || parameter.Type.IsRecord)
        {
            Report(DiagnosticIds.ByRefTypeMismatch, expression.FirstToken()!, $"ByRef argument type mismatch: '{parameter.Name}'.");
            return new BoundArgument(parameter, Convert(value, parameter.Type), ArgumentMode.ByVal);
        }

        if (value is BoundVariable && value.Type.IsVariant && !parameter.Type.IsVariant)
        {
            // The VBE rejects a Variant variable passed to a typed ByRef parameter (Procedures bisect, 2026-09-05).
            Report(DiagnosticIds.ByRefTypeMismatch, expression.FirstToken()!, $"ByRef argument type mismatch: '{parameter.Name}'.");
            return new BoundArgument(parameter, Convert(value, parameter.Type), ArgumentMode.ByVal);
        }

        if (value.Type.IsVariant || parameter.Type.IsVariant)
        {
            return new BoundArgument(parameter, value, ArgumentMode.ByRefCopyBack);
        }

        if (value.Type.Kind == TypeKind.FixedString && parameter.Type.IsVariableString)
        {
            // A fixed-length String reaches a ByRef String parameter as a copy, stored back cut to its length (Arrays golden: an element padded "ab " comes back "ab ").
            return new BoundArgument(parameter, value, ArgumentMode.ByRefCopyBack);
        }

        Report(DiagnosticIds.ByRefTypeMismatch, expression.FirstToken()!, $"ByRef argument type mismatch: '{parameter.Name}'.");
        return new BoundArgument(parameter, Convert(value, parameter.Type), ArgumentMode.ByVal);
    }

    /// <summary>
    /// AddressOf procedure (MS-VBAL 5.6.16.8): the address of one of this project's own procedures,
    /// which a DLL calls back. Only shapes a callback can carry are allowed: fixed parameters of
    /// machine types, Strings, and Variants, and a result of one of those too (Declares golden).
    /// </summary>
    private BoundExpression BindAddressOf(AddressOfExpressionSyntax syntax)
    {
        var target = syntax.Procedure switch
        {
            IdentifierNameSyntax name => Lookup(name.Name) as ProcedureSymbol,
            MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax owner, Name.Name: var member } =>
                project.FindModule(owner.Name)?.Members.GetValueOrDefault(member) as ProcedureSymbol,
            _ => null,
        };

        if (target is null || target.IsExternal)
        {
            Report(DiagnosticIds.ProcedureNotDefined, syntax.Procedure.FirstToken()!, $"Sub or Function not defined: '{syntax.Procedure.ToFullString().Trim()}'.");
            return Empty(syntax);
        }

        foreach (var parameter in target.Parameters)
        {
            if (parameter.IsOptional || parameter.IsParamArray || !CallableParameter(parameter.Type))
            {
                Report(DiagnosticIds.NotSupported, syntax.Procedure.FirstToken()!, $"'{target.Name}' cannot be a callback: the parameter '{parameter.Name}' needs its own marshaling.");
                return Empty(syntax);
            }
        }

        if (target.ReturnType is { } returnType && !Callable(returnType) && !returnType.IsVariableString && !returnType.IsVariant)
        {
            Report(DiagnosticIds.NotSupported, syntax.Procedure.FirstToken()!, $"'{target.Name}' cannot be a callback: its result needs its own marshaling.");
            return Empty(syntax);
        }

        return new BoundAddressOf(syntax, target);
    }

    /// <summary>A type a callback can carry across the boundary: the fixed-size machine types.</summary>
    private static bool Callable(VbaType type) =>
        type.Kind is TypeKind.Builtin or TypeKind.Enum && !type.IsVariant && !type.IsObject && !type.IsString;

    /// <summary>
    /// A parameter a native caller can hand a callback as it is: a machine type, a String as its
    /// BSTR (ByRef, its address), a Variant as the address of its VARIANT, as VBA's procedures take
    /// them (Declares golden). The procedure copies a ByVal String or Variant on entry, so the
    /// caller's stays the caller's.
    /// </summary>
    private static bool CallableParameter(VbaType type) => Callable(type) || type.IsVariableString || type.IsVariant;

    /// <summary>True for a name that resolved outside the current module and its procedure: another module's public member or a library name.</summary>
    private bool DeclaredElsewhere(Symbol symbol) => symbol switch
    {
        VariableSymbol variable => variable.Procedure is null && variable.Module != module,
        ConstantSymbol constant => constant.Module != module,
        ProcedureSymbol procedureSymbol => procedureSymbol.Module != module,
        PropertySymbol property => property.Module != module,
        RecordSymbol record => record.Module != module,
        EnumSymbol enumeration => enumeration.Module != module,
        _ => true,
    };

    private BoundExpression BindAssertMember(string name, IReadOnlyList<ArgumentSyntax> arguments, MemberAccessExpressionSyntax member, bool isCallStatement)
    {
        if (StandardLibrary.TryGetFunction("Assert." + name, out var assertion))
        {
            return BindLibrarySub(assertion, arguments, member, isCallStatement);
        }

        Report(DiagnosticIds.ProcedureNotDefined, member.Name.Identifier, $"Sub or Function not defined: 'Assert.{name}'.");
        return Empty(member);
    }

    /// <summary>A library Sub such as Debug.Assert: a call statement, never a value (MS-VBAL 5.6.13: "Expected Function or variable").</summary>
    private BoundExpression BindLibrarySub(IntrinsicSymbol intrinsic, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        if (!isCallStatement)
        {
            Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Expected Function or variable.");
            return Empty(syntax);
        }

        return BindIntrinsicCall(intrinsic, arguments, syntax, dollar: false);
    }

    private BoundExpression BindIntrinsicCall(IntrinsicSymbol intrinsic, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool dollar)
    {
        if (intrinsic.Pending is not null)
        {
            // A reason starting "never" is one vba-ng will not implement; anything else is not implemented yet.
            var reason = intrinsic.Pending.StartsWith("never: ", StringComparison.Ordinal)
                ? $"{intrinsic.Name} is not supported: {intrinsic.Pending["never: ".Length..]}."
                : $"{intrinsic.Name} needs {intrinsic.Pending}.";
            Report(DiagnosticIds.NotSupported, syntax.FirstToken()!, reason);
            return Empty(syntax);
        }

        var values = new List<BoundExpression?>();
        var parameterNames = StandardLibrary.ParameterNames(intrinsic.Name);
        foreach (var argument in arguments)
        {
            if (argument.IsNamed)
            {
                // A library function's parameters are named as MS-VBAL 6.1 documents them: Replace(..., Count:=1).
                var slot = parameterNames.FindIndex(p => p.Equals(argument.Name!.NameValue, StringComparison.OrdinalIgnoreCase));
                if (slot < 0)
                {
                    Report(DiagnosticIds.NamedArgumentNotFound, argument.Name!, $"Named argument not found: '{argument.Name!.NameValue}'.");
                    continue;
                }

                while (values.Count <= slot)
                {
                    values.Add(null);
                }

                values[slot] = Convert(BindExpression(argument.Expression!), VbaType.Variant);
                continue;
            }

            values.Add(argument.IsOmitted ? null : Convert(BindExpression(argument.Expression!), VbaType.Variant));
        }

        while (values.Count > 0 && values[^1] is null)
        {
            values.RemoveAt(values.Count - 1);
        }

        if (intrinsic.Name is "LBound" or "UBound" && values.Count > 0 && values[0] is BoundConversion { Operand.Type: var bounded } && !bounded.IsArray && !bounded.IsVariant)
        {
            // The VBE rejects LBound and UBound of a non-array at compile time.
            Report(DiagnosticIds.ExpectedArray, syntax.FirstToken()!, "Expected array.");
        }

        if (values.Count < intrinsic.MinArguments)
        {
            Report(DiagnosticIds.ArgumentNotOptional, syntax.FirstToken()!, $"Argument not optional: '{intrinsic.Name}'.");
        }
        else if (values.Count > intrinsic.MaxArguments)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, syntax.FirstToken()!, $"Wrong number of arguments or invalid property assignment: '{intrinsic.Name}'.");
        }

        return new BoundIntrinsicCall(syntax, intrinsic, values, dollar ? VbaType.String : intrinsic.ReturnType) { DollarForm = dollar };
    }

    private BoundExpression BindUnary(UnaryExpressionSyntax unary)
    {
        var operand = BindExpression(unary.Operand);
        switch (unary.OperatorToken.Kind)
        {
            case SyntaxKind.MinusToken:
                return new BoundUnary(unary, UnaryKind.Negate, operand, NegateType(operand.Type));
            case SyntaxKind.PlusToken:
                return new BoundUnary(unary, UnaryKind.Plus, operand, operand.Type);
            case SyntaxKind.NotKeyword:
                return new BoundUnary(unary, UnaryKind.Not, operand, NotType(operand.Type));
            case SyntaxKind.HashToken:
                // #n in argument position names a file number, as in Input(5, #1); the number is the value.
                return operand;
            default:
                Report(DiagnosticIds.SyntaxError, unary.OperatorToken, "Expected: expression.");
                return operand;
        }
    }

    private BoundBinary BindBinary(BinaryExpressionSyntax binary)
    {
        var left = BindExpression(binary.Left);
        var right = BindExpression(binary.Right);
        var kind = binary.OperatorToken.Kind switch
        {
            SyntaxKind.PlusToken => BinaryKind.Add,
            SyntaxKind.MinusToken => BinaryKind.Subtract,
            SyntaxKind.AsteriskToken => BinaryKind.Multiply,
            SyntaxKind.SlashToken => BinaryKind.Divide,
            SyntaxKind.BackslashToken => BinaryKind.IntegerDivide,
            SyntaxKind.ModKeyword => BinaryKind.Modulo,
            SyntaxKind.CaretToken => BinaryKind.Power,
            SyntaxKind.AmpersandToken => BinaryKind.Concatenate,
            SyntaxKind.EqualsToken => BinaryKind.Equal,
            SyntaxKind.LessThanGreaterThanToken => BinaryKind.NotEqual,
            SyntaxKind.LessThanToken => BinaryKind.LessThan,
            SyntaxKind.GreaterThanToken => BinaryKind.GreaterThan,
            SyntaxKind.LessThanEqualsToken => BinaryKind.LessThanOrEqual,
            SyntaxKind.GreaterThanEqualsToken => BinaryKind.GreaterThanOrEqual,
            SyntaxKind.LikeKeyword => BinaryKind.Like,
            SyntaxKind.IsKeyword => BinaryKind.Is,
            SyntaxKind.AndKeyword => BinaryKind.And,
            SyntaxKind.OrKeyword => BinaryKind.Or,
            SyntaxKind.XorKeyword => BinaryKind.Xor,
            SyntaxKind.EqvKeyword => BinaryKind.Eqv,
            _ => BinaryKind.Imp,
        };

        if (kind == BinaryKind.Is && ((!left.Type.IsObject && !left.Type.IsVariant) || (!right.Type.IsObject && !right.Type.IsVariant)))
        {
            Report(DiagnosticIds.ObjectRequired, binary.OperatorToken, "Object required: Is compares object references.");
        }

        foreach (var operand in new[] { left, right })
        {
            if (operand.Type.IsArray || operand.Type.IsRecord)
            {
                Report(DiagnosticIds.TypeMismatch, operand.Syntax.FirstToken()!, "Type mismatch: arrays and user-defined types have no operators.");
            }
        }

        var declared = DeclaredTypes.None;
        if (left.Type.IsVariant)
        {
            declared |= DeclaredTypes.LeftVariant;
        }

        if (right.Type.IsVariant)
        {
            declared |= DeclaredTypes.RightVariant;
        }

        return new BoundBinary(binary, kind, left, right, declared, BinaryResultType(kind, left.Type, right.Type));
    }

    /// <summary>The declared type of an operator expression (MS-VBAL 5.6.9.3, 5.6.9.5, 5.6.9.7); Variant whenever an operand is.</summary>
    private static VbaType BinaryResultType(BinaryKind kind, VbaType left, VbaType right)
    {
        if (left.IsVariant || right.IsVariant)
        {
            return VbaType.Variant;
        }

        if (left.IsObject || right.IsObject)
        {
            return kind == BinaryKind.Is ? VbaType.Boolean : VbaType.Variant;
        }

        var l = Underlying(left);
        var r = Underlying(right);
        switch (kind)
        {
            case BinaryKind.Add when l == VarType.String && r == VarType.String:
                return VbaType.String;
            case BinaryKind.Add:
            case BinaryKind.Subtract:
            case BinaryKind.Multiply:
                return VbaType.FromVarType(ArithmeticType(l, r));
            case BinaryKind.Divide:
                return l == VarType.Decimal || r == VarType.Decimal ? VbaType.Variant : VbaType.Double;
            case BinaryKind.IntegerDivide:
            case BinaryKind.Modulo:
                return VbaType.FromVarType(IntegralType(l, r));
            case BinaryKind.Power:
                return VbaType.Double;
            case BinaryKind.Concatenate:
                return VbaType.String;
            case BinaryKind.Equal:
            case BinaryKind.NotEqual:
            case BinaryKind.LessThan:
            case BinaryKind.GreaterThan:
            case BinaryKind.LessThanOrEqual:
            case BinaryKind.GreaterThanOrEqual:
            case BinaryKind.Like:
            case BinaryKind.Is:
                return VbaType.Boolean;
            default:
                return VbaType.FromVarType(LogicalType(l, r));
        }
    }

    private static VarType Underlying(VbaType type) => type.Kind switch
    {
        TypeKind.Enum => VarType.Long,
        TypeKind.FixedString => VarType.String,
        _ => type.VarType,
    };

    /// <summary>MS-VBAL 5.6.9.3.1 effective type of + - * on declared types (mirrors Operators.EffectiveArithmeticType).</summary>
    private static VarType ArithmeticType(VarType l, VarType r)
    {
        if (l == VarType.Decimal || r == VarType.Decimal)
        {
            return VarType.Decimal;
        }

        if (l == VarType.Date || r == VarType.Date)
        {
            return VarType.Date;
        }

        if (l == VarType.Currency || r == VarType.Currency)
        {
            return VarType.Currency;
        }

        if (l is VarType.Double or VarType.String || r is VarType.Double or VarType.String)
        {
            return VarType.Double;
        }

        if (l == VarType.Single || r == VarType.Single)
        {
            var other = l == VarType.Single ? r : l;
            return other is VarType.Long or VarType.LongLong ? VarType.Double : VarType.Single;
        }

        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        if (l == VarType.Long || r == VarType.Long)
        {
            return VarType.Long;
        }

        return l == VarType.Byte && r == VarType.Byte ? VarType.Byte : VarType.Integer;
    }

    /// <summary>MS-VBAL 5.6.9.3.5: \ and Mod produce an integral type.</summary>
    private static VarType IntegralType(VarType l, VarType r)
    {
        static VarType Integral(VarType t) => t switch
        {
            VarType.Byte or VarType.Boolean or VarType.Integer => VarType.Integer,
            VarType.LongLong => VarType.LongLong,
            _ => VarType.Long,
        };

        if (l == VarType.Byte && r == VarType.Byte)
        {
            return VarType.Byte;
        }

        var a = Integral(l);
        var b = Integral(r);
        if (a == VarType.LongLong || b == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        return a == VarType.Long || b == VarType.Long ? VarType.Long : VarType.Integer;
    }

    /// <summary>MS-VBAL 5.6.9.7 effective type of the logical operators (mirrors Operators.EffectiveLogicalType).</summary>
    private static VarType LogicalType(VarType l, VarType r)
    {
        if (l == VarType.LongLong || r == VarType.LongLong)
        {
            return VarType.LongLong;
        }

        static bool IsLongClass(VarType t) => t.IsFloatingPoint() || t.IsFixedPoint() || t is VarType.Long or VarType.String or VarType.Date;
        if (IsLongClass(l) || IsLongClass(r))
        {
            return VarType.Long;
        }

        if (l == VarType.Byte && r == VarType.Byte)
        {
            return VarType.Byte;
        }

        return l == VarType.Boolean && r == VarType.Boolean ? VarType.Boolean : VarType.Integer;
    }

    /// <summary>MS-VBAL 5.6.9.3.1 unary minus: Byte and Boolean give Integer, String and Date give Double.</summary>
    private static VbaType NegateType(VbaType type)
    {
        if (type.IsVariant || type.IsObject)
        {
            return VbaType.Variant;
        }

        return Underlying(type) switch
        {
            VarType.Byte or VarType.Boolean => VbaType.Integer,
            VarType.String or VarType.Date => VbaType.Double,
            VarType.Decimal => VbaType.Variant,
            var v => VbaType.FromVarType(v),
        };
    }

    /// <summary>MS-VBAL 5.6.9.7 Not: Boolean stays Boolean; other types become their integral form.</summary>
    private static VbaType NotType(VbaType type)
    {
        if (type.IsVariant || type.IsObject)
        {
            return VbaType.Variant;
        }

        return Underlying(type) switch
        {
            VarType.Boolean => VbaType.Boolean,
            VarType.Byte => VbaType.Byte,
            VarType.Integer => VbaType.Integer,
            VarType.LongLong => VbaType.LongLong,
            _ => VbaType.Long,
        };
    }

    /// <summary>Let-coerces an expression to a declared type (MS-VBAL 5.5.1); reports the mismatches VBA rejects at compile time.</summary>
    private BoundExpression Convert(BoundExpression value, VbaType target)
    {
        var source = value.Type;
        if (source.SameAs(target))
        {
            return value;
        }

        var syntax = value.Syntax;
        if (target.IsVariant)
        {
            if (source.IsRecord)
            {
                Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Only user-defined types defined in public object modules can be coerced to or from a Variant.");
            }

            return new BoundConversion(syntax, value, target);
        }

        if (target.IsArray)
        {
            var fits = source.IsVariant
                || (source.IsArray && source.ElementType!.SameAs(target.ElementType!))
                || (source.IsString && ReferenceEquals(target.ElementType, VbaType.Byte));
            if (!fits)
            {
                Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Can't assign to array.");
            }

            return new BoundConversion(syntax, value, target);
        }

        if (source.IsArray)
        {
            if (!(target.IsString && ReferenceEquals(source.ElementType, VbaType.Byte)))
            {
                Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Type mismatch: an array cannot be assigned to a scalar.");
            }

            return new BoundConversion(syntax, value, target);
        }

        if (target.IsRecord || source.IsRecord)
        {
            Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Type mismatch: user-defined types assign only to the same type.");
            return new BoundConversion(syntax, value, target);
        }

        if (target.IsObject && !source.IsObject && !source.IsVariant)
        {
            Report(DiagnosticIds.ObjectRequired, syntax.FirstToken()!, "Object required.");
        }

        return new BoundConversion(syntax, value, target);
    }

    /// <summary>Evaluates a constant expression at compile time (MS-VBAL 5.2.3.2): literals, constants, and the operators.</summary>
    private Variant? FoldConstant(ExpressionSyntax syntax)
    {
        try
        {
            return Fold(syntax);
        }
        catch (VbaException ex)
        {
            Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, ex.Description + ".");
            return null;
        }
    }

    private Variant? Fold(ExpressionSyntax syntax)
    {
        switch (syntax)
        {
            case LiteralExpressionSyntax literal:
                {
                    var bound = BindLiteral(literal);
                    return bound is BoundLiteral { Value: var value } && !value.IsNothing ? value : ReportNotConstant(syntax);
                }

            case ParenthesizedExpressionSyntax parenthesized:
                return Fold(parenthesized.Expression);
            case IdentifierNameSyntax name:
                return FoldName(name.Name, null, name.Identifier);
            case MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax qualifier } member:
                return FoldName(member.Name.Name, qualifier.Name, member.Name.Identifier);
            case MemberAccessExpressionSyntax { Expression: MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax library, Name.Name: var enumeration } } member
                when Lookup(library.Name) is null && project.FindModule(library.Name) is null:
                // library.Enum.Member, as in an Optional default of VBA.VbMsgBoxStyle.vbExclamation,
                // which the VBE folds (MS-VBAL 6.1.1) and stdVBA writes. The library name drops off
                // and its enum qualifies the member, which is the two-part form above; the guard
                // keeps a variable or module of that name from being read as a library.
                return FoldName(member.Name.Name, enumeration, member.Name.Identifier);
            case UnaryExpressionSyntax unary:
                {
                    var operand = Fold(unary.Operand);
                    if (operand is null)
                    {
                        return null;
                    }

                    return unary.OperatorToken.Kind switch
                    {
                        SyntaxKind.MinusToken => Operators.Negate(operand.Value),
                        SyntaxKind.NotKeyword => Operators.Not(operand.Value),
                        _ => operand,
                    };
                }

            case BinaryExpressionSyntax binary:
                {
                    var left = Fold(binary.Left);
                    var right = Fold(binary.Right);
                    if (left is null || right is null)
                    {
                        return null;
                    }

                    var l = left.Value;
                    var r = right.Value;
                    var compare = module.Options.Compare;
                    return binary.OperatorToken.Kind switch
                    {
                        SyntaxKind.PlusToken => Operators.Add(l, r),
                        SyntaxKind.MinusToken => Operators.Subtract(l, r),
                        SyntaxKind.AsteriskToken => Operators.Multiply(l, r),
                        SyntaxKind.SlashToken => Operators.Divide(l, r),
                        SyntaxKind.BackslashToken => Operators.IntegerDivide(l, r),
                        SyntaxKind.ModKeyword => Operators.Modulo(l, r),
                        SyntaxKind.CaretToken => Operators.Power(l, r),
                        SyntaxKind.AmpersandToken => Operators.Concatenate(l, r),
                        SyntaxKind.EqualsToken => Operators.Equal(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.LessThanGreaterThanToken => Operators.NotEqual(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.LessThanToken => Operators.LessThan(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.GreaterThanToken => Operators.GreaterThan(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.LessThanEqualsToken => Operators.LessThanOrEqual(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.GreaterThanEqualsToken => Operators.GreaterThanOrEqual(l, r, DeclaredTypes.None, compare),
                        SyntaxKind.AndKeyword => Operators.And(l, r),
                        SyntaxKind.OrKeyword => Operators.Or(l, r),
                        SyntaxKind.XorKeyword => Operators.Xor(l, r),
                        SyntaxKind.EqvKeyword => Operators.Eqv(l, r),
                        SyntaxKind.ImpKeyword => Operators.Imp(l, r),
                        _ => ReportNotConstant(syntax),
                    };
                }

            default:
                return ReportNotConstant(syntax);
        }
    }

    private Variant? FoldName(string name, string? qualifier, SyntaxToken at)
    {
        if (qualifier is null && procedure is not null && procedure.Locals.TryGetValue(name, out var local))
        {
            return local is ConstantSymbol localConstant ? localConstant.Value : ReportNotConstant(at);
        }

        if (qualifier is not null && (qualifier.Equals("VBA", StringComparison.OrdinalIgnoreCase) || StandardLibrary.IsEnumName(qualifier)))
        {
            // VBA.vbLongLong, or the library's own enum as the qualifier: VbVarType.vbLongLong (MS-VBAL 6.1.1).
            return StandardLibrary.TryGetConstant(name, out var library) ? library.Value : ReportNotConstant(at);
        }

        if (qualifier is not null && Lookup(qualifier) is EnumSymbol enumeration)
        {
            var member = enumeration.Members.FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            return member is null ? ReportNotConstant(at) : member.Value;
        }

        if (qualifier is null && Lookup(name) is ConstantSymbol visible)
        {
            // A module constant or an enum member in scope, its own module's or a public one of another (MS-VBAL 5.2.3.2).
            return visible.Value;
        }

        if (Lookup(name) is ComConstantSymbol libraryConstant && (qualifier is null || Lookup(qualifier) is null))
        {
            // A referenced library's constant, bare (xlScreen) or under its enum or library name (XlPictureAppearance.xlScreen).
            return libraryConstant.Value;
        }

        foreach (var candidate in project.Modules)
        {
            if (qualifier is not null && !candidate.Name.Equals(qualifier, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (candidate.Members.TryGetValue(name, out var member) && member is ConstantSymbol constant && (ReferenceEquals(candidate, module) || constant.IsPublic))
            {
                return constant.Value;
            }

            if (pendingConstants.TryGetValue(candidate.Name + "." + name, out var pending))
            {
                var current = module;
                module = pending.Module;
                var folded = FoldModuleConstant(pending.Syntax);
                module = current;
                if (folded is not null && (ReferenceEquals(candidate, current) || folded.IsPublic))
                {
                    return folded.Value;
                }
            }
        }

        if (qualifier is null && StandardLibrary.TryGetConstant(name, out var standard))
        {
            return standard.Value;
        }

        return ReportNotConstant(at);
    }

    /// <summary>The type a member access reaches through user-defined types alone (t.Inner, t.Inner.Deeper), or null when the chain starts elsewhere or names no field.</summary>
    private VbaType? RecordMemberType(MemberAccessExpressionSyntax access)
    {
        var receiver = access.Expression switch
        {
            IdentifierNameSyntax root => Lookup(root.Name) is VariableSymbol { IsArray: false } variable && variable.Type.IsRecord ? variable.Type : null,
            MemberAccessExpressionSyntax inner => RecordMemberType(inner),
            _ => null,
        };
        return receiver?.Record?.FindField(access.Name.Name) is { IsArray: false } field ? field.Type : null;
    }

    private Variant? ReportNotConstant(SyntaxNode syntax) => ReportNotConstant(syntax.FirstToken()!);

    private Variant? ReportNotConstant(SyntaxToken at)
    {
        Report(DiagnosticIds.ConstantExpressionRequired, at, "Constant expression required.");
        return null;
    }
}
