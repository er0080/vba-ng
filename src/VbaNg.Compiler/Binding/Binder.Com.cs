using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;
using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// Early binding to referenced type libraries (ARCHITECTURE.md section 6, "Binding modes"): a
/// member of a variable declared with a library type binds by dispid, arguments are let-coerced
/// to the parameter types the library declares, named arguments fall into their positions, and
/// the members of an application object are in scope unqualified.
/// </summary>
public sealed partial class Binder
{
    /// <summary>A member of a COM-typed target: <c>d.Add "a", 1</c>, <c>ws.Name</c>, <c>Cells(1, 2)</c>.</summary>
    private BoundExpression BindComMember(BoundExpression target, ComType owner, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        var candidates = libraries.Members(owner, name).ToList();
        if (candidates.Count == 0)
        {
            if (libraries.IsExtensible(owner))
            {
                // An interface that may grow members at run time (no TYPEFLAG_FNONEXTENSIBLE, as MSForms.Control): VBA binds the name late (docs/vba-quirks.md).
                return BindLateMember(Convert(target, VbaType.Variant), name, arguments, syntax);
            }

            Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
            return Empty(syntax);
        }

        ComMember? get = null;
        ComMember? method = null;
        ComMember? put = null;
        ComMember? putRef = null;
        foreach (var (_, candidate) in candidates)
        {
            switch (candidate.Kind)
            {
                case ComMemberKind.PropertyGet:
                    get ??= candidate;
                    break;
                case ComMemberKind.Method:
                    method ??= candidate;
                    break;
                case ComMemberKind.PropertyPut:
                    put ??= candidate;
                    break;
                case ComMemberKind.PropertyPutRef:
                    putRef ??= candidate;
                    break;
                case ComMemberKind.Constant:
                    return new BoundLiteral(syntax, ReferencedLibraries.ConstantValue(candidate.Value!), libraries.TypeOf(candidate.Type));
                case ComMemberKind.Field:
                    get ??= candidate;
                    break;
            }
        }

        var member = isCallStatement && method is not null ? method : get ?? method ?? put ?? putRef!;
        var kind = member.Kind switch
        {
            ComMemberKind.Method => InvokeKind.Method,
            ComMemberKind.PropertyPut => InvokeKind.PropertyPut,
            ComMemberKind.PropertyPutRef => InvokeKind.PropertyPutRef,
            _ => InvokeKind.PropertyGet,
        };

        // A property with only a put accessor cannot be read; it still binds so an assignment can use it.
        var readable = member.Kind is ComMemberKind.Method or ComMemberKind.PropertyGet or ComMemberKind.Field;
        var parameters = readable ? member.Parameters : member.Parameters.Take(member.Parameters.Count - 1).ToList();
        var type = readable ? libraries.TypeOf(member.Type) : libraries.TypeOf(member.Parameters[^1].Type);

        // Cells(1, 2): a parameterless property returning an object takes the arguments on that object's default member.
        if (arguments.Count > 0 && parameters.Count == 0 && type.IsCom && ReferencedLibraries.DefaultMember(type.ComInterface!) is { } defaultMember)
        {
            var inner = new BoundComCall(syntax, target, member, [], kind, type, put, putRef, null, null);
            return BindComMember(inner, type.ComInterface!, defaultMember.Name, arguments, syntax, isCallStatement: false);
        }

        // The same on a result whose type the referenced libraries do not describe: index it late-bound.
        if (arguments.Count > 0 && parameters.Count == 0 && (type.IsGenericObject || type.IsVariant))
        {
            var inner = new BoundComCall(syntax, target, member, [], kind, type, put, putRef, null, null);
            return BindLateElement(inner, arguments, syntax);
        }

        if (member.Type is null && !isCallStatement && member.Kind == ComMemberKind.Method)
        {
            Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, "Expected Function or variable.");
        }

        var values = BindComArguments(parameters, arguments, syntax, name);
        var putValueType = put is not null ? libraries.TypeOf(put.Parameters[^1].Type) : putRef is not null ? libraries.TypeOf(putRef.Parameters[^1].Type) : null;
        ComMember? defaultPut = null;
        if (put is null && putRef is null && type.IsCom && ReferencedLibraries.DefaultMember(type.ComInterface!) is { } resultDefault)
        {
            defaultPut = libraries.Members(type.ComInterface!, resultDefault.Name).Select(m => m.Member).FirstOrDefault(m => m.Kind == ComMemberKind.PropertyPut);
        }

        return new BoundComCall(syntax, target, member, values, kind, type, put, putRef, defaultPut, putValueType);
    }

    /// <summary>A member of a library's application object used unqualified: ActiveSheet, Range("A1"), Worksheets.Add.</summary>
    private BoundExpression BindComGlobal(ComGlobalSymbol global, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        var appType = libraries.ResolveType(global.Library.Name, global.AppObject.Name) ?? VbaType.Object;
        var target = new BoundAppObject(syntax, global.Library, global.AppObject, appType);
        return BindComMember(target, global.Members, global.Name, arguments, syntax, isCallStatement);
    }

    /// <summary>
    /// [A1] (a foreign name, MS-VBAL 3.3.5.3) that the project does not declare: the host evaluates
    /// its text, which in Excel is Application.Evaluate, so [A1] is a Range and [MyName] a named
    /// range. Null when no referenced library puts an Evaluate in scope. As an assignment target the
    /// value goes to the default member of what Evaluate returns, late-bound, which is what
    /// [A1] = 5 does in VBA.
    /// </summary>
    private BoundExpression? BindForeignName(SyntaxToken foreignName, SyntaxNode syntax, bool asTarget)
    {
        if (Lookup("Evaluate") is not ComGlobalSymbol evaluate)
        {
            return null;
        }

        var argument = ForeignNameArgument(foreignName);
        if (!asTarget)
        {
            return BindComGlobal(evaluate, [argument], syntax, isCallStatement: false);
        }

        var appType = libraries.ResolveType(evaluate.Library.Name, evaluate.AppObject.Name) ?? VbaType.Object;
        var application = Convert(new BoundAppObject(syntax, evaluate.Library, evaluate.AppObject, appType), VbaType.Variant);
        return new BoundMember(syntax, application, evaluate.Name, [Convert(BindExpression(argument.Expression!), VbaType.Variant)], VbaType.Variant);
    }

    /// <summary>
    /// obj.[name]: first a member called name, which is how a member whose identifier is not lexically
    /// valid is reached (items.[_NewEnum]); when an early-bound COM receiver has no such member but
    /// declares Evaluate, the bracketed text is evaluated instead (ws.[A1]). A late-bound receiver
    /// makes the same choice at run time.
    /// </summary>
    private BoundExpression BindForeignMember(BoundExpression target, SyntaxToken foreignName, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement, bool asTarget = false)
    {
        var name = foreignName.Value as string ?? foreignName.Text.Trim('[', ']');
        if (target.Type.IsCom)
        {
            var owner = target.Type.ComInterface!;
            if (!libraries.Members(owner, name).Any() && libraries.Members(owner, "Evaluate").Any())
            {
                // ws.[A1] = 5 assigns to the default member of what Evaluate returns, late-bound,
                // exactly as [A1] = 5 does; VBA-Web's sheet modules write their named ranges that
                // way (ROADMAP.md WP6).
                return asTarget
                    ? new BoundMember(syntax, Convert(target, VbaType.Variant), "Evaluate", [Convert(BindExpression(ForeignNameArgument(foreignName).Expression!), VbaType.Variant)], VbaType.Variant)
                    : BindComMember(target, owner, "Evaluate", [ForeignNameArgument(foreignName)], syntax, isCallStatement);
            }

            return BindComMember(target, owner, name, arguments, syntax, isCallStatement);
        }

        if (target.Type.IsVariant || target.Type.IsGenericObject)
        {
            var member = BindMemberOn(target, name, arguments, syntax, isCallStatement);
            return member is BoundMember late ? new BoundMember(syntax, late.Target, late.Member, late.Arguments, late.Type) { IsForeign = true } : member;
        }

        return BindMemberOn(target, name, arguments, syntax, isCallStatement);
    }

    /// <summary>The bracketed text of a foreign name as a string literal argument, for the host's Evaluate.</summary>
    private static ArgumentSyntax ForeignNameArgument(SyntaxToken foreignName)
    {
        var text = foreignName.Value as string ?? foreignName.Text.Trim('[', ']');
        var literal = new SyntaxToken(SyntaxKind.StringLiteralToken, "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"", foreignName.Start, value: text);
        return new ArgumentSyntax(null, null, null, new LiteralExpressionSyntax(literal));
    }

    /// <summary>
    /// Positional and named arguments against the parameters the library declares: each value is
    /// let-coerced to a scalar parameter's type, objects and Variants pass as they are, omitted
    /// optional parameters between supplied ones become Missing, and trailing ones are dropped.
    /// </summary>
    private List<BoundExpression> BindComArguments(IReadOnlyList<ComParameter> parameters, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, string memberName)
    {
        var paramArray = parameters.Count > 0 && parameters[^1].IsParamArray ? parameters[^1] : null;
        var fixedCount = paramArray is null ? parameters.Count : parameters.Count - 1;
        var slots = new BoundExpression?[fixedCount];
        var rest = new List<BoundExpression>();
        var position = 0;
        foreach (var argument in arguments)
        {
            int slot;
            if (argument.IsNamed)
            {
                var named = argument.Name!.NameValue;
                slot = -1;
                for (var i = 0; i < fixedCount; i++)
                {
                    if (parameters[i].Name.Equals(named, StringComparison.OrdinalIgnoreCase))
                    {
                        slot = i;
                        break;
                    }
                }

                if (slot < 0)
                {
                    Report(DiagnosticIds.NamedArgumentNotFound, argument.Name!, $"Named argument not found: '{named}'.");
                    continue;
                }
            }
            else
            {
                slot = position++;
                if (slot >= fixedCount)
                {
                    if (paramArray is null)
                    {
                        Report(DiagnosticIds.WrongNumberOfArguments, argument.FirstToken() ?? syntax.FirstToken()!, $"Wrong number of arguments or invalid property assignment: '{memberName}'.");
                        continue;
                    }

                    if (!argument.IsOmitted)
                    {
                        rest.Add(Convert(BindExpression(argument.Expression!), VbaType.Variant));
                    }

                    continue;
                }
            }

            if (argument.IsOmitted)
            {
                continue;
            }

            slots[slot] = ComArgument(BindExpression(argument.Expression!), parameters[slot]);
        }

        var last = Array.FindLastIndex(slots, s => s is not null);
        var values = new List<BoundExpression>();
        for (var i = 0; i <= last; i++)
        {
            if (slots[i] is { } value)
            {
                values.Add(value);
            }
            else
            {
                if (!parameters[i].IsOptional)
                {
                    Report(DiagnosticIds.ArgumentNotOptional, syntax.FirstToken()!, $"Argument not optional: '{parameters[i].Name}'.");
                }

                values.Add(new BoundLiteral(syntax, Variant.Missing, VbaType.Variant));
            }
        }

        for (var i = last + 1; i < fixedCount; i++)
        {
            if (!parameters[i].IsOptional)
            {
                Report(DiagnosticIds.ArgumentNotOptional, syntax.FirstToken()!, $"Argument not optional: '{parameters[i].Name}'.");
            }
        }

        values.AddRange(rest);
        return values;
    }

    /// <summary>An argument let-coerced to a scalar parameter type (MS-VBAL 5.6.13.1); objects and Variants pass unchanged.</summary>
    private BoundExpression ComArgument(BoundExpression value, ComParameter parameter)
    {
        var declared = libraries.TypeOf(parameter.Type);
        if (declared.IsScalar && !declared.IsVariant && !value.Type.IsObject && !value.Type.IsVariant)
        {
            return Convert(Convert(value, declared), VbaType.Variant);
        }

        return Convert(value, VbaType.Variant);
    }

    /// <summary>d("k") on a library type: the arguments go to the default member (DISPID_VALUE), bound by dispid.</summary>
    private BoundExpression BindComDefault(BoundExpression target, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        var defaultMember = ReferencedLibraries.DefaultMember(target.Type.ComInterface!);
        if (defaultMember is null)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, syntax.FirstToken()!, "Wrong number of arguments or invalid property assignment.");
            return Empty(syntax);
        }

        return BindComMember(target, target.Type.ComInterface!, defaultMember.Name, arguments, syntax, isCallStatement: false);
    }

    /// <summary>A constant reached through its enum or library name: XlDirection.xlUp, Excel.xlUp.</summary>
    private BoundLiteral? BindLibraryConstant(string qualifier, string name, SyntaxNode syntax)
    {
        foreach (var library in libraries.Libraries)
        {
            var byLibrary = library.Name.Equals(qualifier, StringComparison.OrdinalIgnoreCase);
            foreach (var type in library.Types)
            {
                if (type.Kind is not (ComTypeKind.Enum or ComTypeKind.Module))
                {
                    continue;
                }

                if (!byLibrary && !type.Name.Equals(qualifier, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var member = type.FindMember(name, ComMemberKind.Constant);
                if (member?.Value is not null)
                {
                    return new BoundLiteral(syntax, ReferencedLibraries.ConstantValue(member.Value), type.Kind == ComTypeKind.Enum ? VbaType.Long : libraries.TypeOf(member.Type));
                }
            }
        }

        return null;
    }
}
