using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;
using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// Class modules (MS-VBAL 4.2, 5.2.4.1): the class type and its instance members, the
/// predeclared instance, default members, events, and Implements. ARCHITECTURE.md D18 governs
/// the lifetime of the instances.
/// </summary>
public sealed partial class Binder
{
    /// <summary>
    /// A .cls with a class header (MS-VBAL 4.2). Excel's document modules carry
    /// VB_PredeclaredId = True together with VB_Exposed = True, which is how they are told from a
    /// class that merely has a predeclared instance (ARCHITECTURE.md section 3).
    /// </summary>
    private void DeclareClassModule(SyntaxTree tree, bool predeclared, bool exposed)
    {
        if (manifest.IsDocumentModule(module.Name, predeclared, exposed))
        {
            // The workbook's CodeNames or the manifest decide; the attribute pairing only when neither speaks (ARCHITECTURE.md D19).
            module.Kind = ModuleKind.Document;
            module.DocumentKind = manifest.DocumentKind(module.Name);
            var meType = libraries.ResolveType("Excel", module.DocumentKind) ?? VbaType.Object;
            module.Me = new VariableSymbol("Me", meType, VariableKind.Module)
            {
                IsPublic = true,
                Module = module,
                EmitName = "Me",
            };
            return;
        }

        if (predeclared && exposed && manifest.WorkbookDocuments is not null)
        {
            // The attributes say document module, the workbook beside the folder names no such object: the host could never bind it (VBA0024).
            var attribute = tree.Root.Members.OfType<AttributeStatementSyntax>().First(a => a.Name is IdentifierNameSyntax { Name: var n } && n.Equals("VB_PredeclaredId", StringComparison.OrdinalIgnoreCase));
            Warn(DiagnosticIds.UnboundDocumentModule, attribute.FirstToken()!, $"'{module.Name}' has a document module's attributes, but the workbook has no object with that CodeName; it compiles as a class and its event procedures never run. List it in the manifest's documents map if it is one.");
        }

        module.Kind = ModuleKind.Class;
        module.ClassType = VbaType.ForClass(module);
        module.Me = new VariableSymbol("Me", module.ClassType, VariableKind.Module)
        {
            Module = module,
            EmitName = "this",
        };

        if (predeclared)
        {
            // VB_PredeclaredId = True: the class name reaches an instance created on first use, as As New does (Classes golden).
            module.Instance = new VariableSymbol(module.Name, module.ClassType, VariableKind.Module)
            {
                IsPublic = true,
                IsNew = true,
                Module = module,
                EmitName = "__Instance",
            };
        }
    }

    private void DeclareEvent(EventDeclarationSyntax declaration)
    {
        if (module.Kind != ModuleKind.Class)
        {
            Report(DiagnosticIds.InvalidStatementPlacement, declaration.EventKeyword, "Event declarations are only valid in a class module.");
            return;
        }

        var symbol = new EventSymbol(declaration.Name.NameValue, module, declaration);
        if (module.Events.Any(e => e.Name.Equals(symbol.Name, StringComparison.OrdinalIgnoreCase)))
        {
            Report(DiagnosticIds.AmbiguousName, declaration.Name, $"Ambiguous name detected: '{symbol.Name}'.");
            return;
        }

        module.Events.Add(symbol);
    }

    /// <summary>Implements class (MS-VBAL 5.2.4.2); the named class is resolved once every module is declared.</summary>
    private void DeclareImplements(ImplementsStatementSyntax statement)
    {
        if (module.Kind != ModuleKind.Class)
        {
            Report(DiagnosticIds.InvalidStatementPlacement, statement.ImplementsKeyword, "Implements is only valid in a class module.");
            return;
        }

        pendingImplements.Add((module, statement));
    }

    /// <summary>
    /// The second pass over class declarations, once every module and member exists: the classes an
    /// Implements names, the members marked by VB_UserMemId, the event parameters, and the
    /// procedures that handle a WithEvents source's events or implement an interface's members.
    /// </summary>
    private void ResolveClasses()
    {
        foreach (var (owner, statement) in pendingImplements)
        {
            module = owner;
            var type = BindNamedType(statement.Name);
            if (type.ProjectClass is not { } implemented)
            {
                Report(DiagnosticIds.TypeNotDefined, statement.Name.FirstToken()!, "Implements needs a class module of this project: '" + statement.Name.ToFullString().Trim() + "'.");
                continue;
            }

            implemented.IsImplemented = true;
            if (!owner.Implemented.Contains(implemented))
            {
                owner.Implemented.Add(implemented);
            }
        }

        foreach (var moduleSymbol in project.Modules)
        {
            module = moduleSymbol;
            foreach (var symbol in module.Events)
            {
                var index = 0;
                foreach (var parameterSyntax in symbol.Syntax.Parameters?.Parameters.ToList() ?? [])
                {
                    symbol.Parameters.Add(BindParameter(parameterSyntax, index++));
                }
            }

            if (module.Kind == ModuleKind.Class)
            {
                foreach (var procedureSymbol in module.Procedures)
                {
                    ResolveClassProcedure(procedureSymbol);
                }
            }
            else if (module.Kind == ModuleKind.Document)
            {
                // A document module handles its WithEvents variables' events like a class; its own object's and its controls' are the host's (DeclareEventHandler).
                foreach (var procedureSymbol in module.Procedures)
                {
                    ResolveEventHandler(procedureSymbol);
                }
            }
        }
    }

    /// <summary>
    /// A class procedure's role: the member VB_UserMemId marks as the default member (0) or the
    /// enumerator (-4), the implementation of an interface member (<c>IShape_Area</c>), or the
    /// handler of a WithEvents source's event (<c>source_Changed</c>).
    /// </summary>
    private void ResolveClassProcedure(ProcedureSymbol symbol)
    {
        symbol.UserMemberId = UserMemberId(symbol);
        switch (symbol.UserMemberId)
        {
            case 0:
                module.DefaultMember = module.Members.GetValueOrDefault(symbol.Name);
                break;
            case -4:
                module.EnumMember = symbol;
                break;
        }

        var split = symbol.Name.IndexOf('_', StringComparison.Ordinal);
        while (split > 0 && split < symbol.Name.Length - 1)
        {
            var qualifier = symbol.Name[..split];
            var member = symbol.Name[(split + 1)..];
            if (module.Implemented.FirstOrDefault(i => i.Name.Equals(qualifier, StringComparison.OrdinalIgnoreCase)) is { } implemented)
            {
                symbol.ImplementsInterface = implemented;
                symbol.ImplementsMember = member;
                return;
            }

            if (module.Members.GetValueOrDefault(qualifier) is VariableSymbol { IsWithEvents: true })
            {
                ResolveEventHandler(symbol);
                return;
            }

            split = symbol.Name.IndexOf('_', split + 1);
        }
    }

    /// <summary>
    /// The handler of a WithEvents variable's event (<c>source_Changed</c>, MS-VBAL 5.2.3.1.4). A
    /// project class must declare the event; a library type's source interface may not, in which
    /// case the procedure is an ordinary one, as it is in VBA.
    /// </summary>
    private void ResolveEventHandler(ProcedureSymbol symbol)
    {
        var split = symbol.Name.IndexOf('_', StringComparison.Ordinal);
        while (split > 0 && split < symbol.Name.Length - 1)
        {
            var qualifier = symbol.Name[..split];
            var member = symbol.Name[(split + 1)..];
            if (module.Members.GetValueOrDefault(qualifier) is VariableSymbol { IsWithEvents: true } source)
            {
                if (source.Type.ProjectClass is { } sourceClass)
                {
                    if (!sourceClass.Events.Any(e => e.Name.Equals(member, StringComparison.OrdinalIgnoreCase)))
                    {
                        Report(DiagnosticIds.ProcedureNotDefined, symbol.Syntax.Name, $"'{sourceClass.Name}' has no event named '{member}'.");
                    }
                }
                else if (source.EventInterface?.FindMember(member, ComMemberKind.Method) is null)
                {
                    // Not an event of the source: an ordinary procedure, not a control handler either (DeclareEventHandler).
                    symbol.EventSource = null;
                    symbol.EventName = null;
                    return;
                }

                symbol.EventVariable = source;
                symbol.EventSource = qualifier;
                symbol.EventName = member;
                return;
            }

            split = symbol.Name.IndexOf('_', split + 1);
        }
    }

    /// <summary>The value of an <c>Attribute name.VB_UserMemId = n</c> line inside a procedure (MS-VBAL 5.2.4.1.1).</summary>
    private int? UserMemberId(ProcedureSymbol symbol)
    {
        foreach (var statement in symbol.Syntax.Body)
        {
            if (statement is not AttributeStatementSyntax attribute)
            {
                continue;
            }

            var name = attribute.Name.ToFullString().Trim();
            if (!name.EndsWith(".VB_UserMemId", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (attribute.Values.Nodes.FirstOrDefault() is { } value && FoldConstant(value) is { } folded)
            {
                return Coerce.ToInt32(folded);
            }
        }

        return null;
    }

    /// <summary>A member of a class instance (MS-VBAL 5.6.13): a field, a method, or a property, bound to the class the receiver's type names.</summary>
    private BoundExpression BindClassMember(BoundExpression receiver, ModuleSymbol classModule, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isCallStatement)
    {
        if (!classModule.Members.TryGetValue(name, out var symbol) || !IsVisible(symbol, classModule))
        {
            Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
            return Empty(syntax);
        }

        switch (symbol)
        {
            case VariableSymbol { IsArray: true } array when arguments.Count > 0:
                return BindElement(new BoundField(syntax, receiver, array), arguments, syntax);
            case VariableSymbol field:
                {
                    var access = new BoundField(syntax, receiver, field);
                    return arguments.Count == 0 ? access : BindLateElement(access, arguments, syntax);
                }

            case ProcedureSymbol callee:
                return BindCallOrIndex(callee, arguments, syntax, isCallStatement, receiver);
            case PropertySymbol property:
                return BindPropertyGet(property, arguments, syntax, receiver);
            default:
                Report(DiagnosticIds.TypeMismatch, syntax.FirstToken()!, $"'{name}' is a type; expected: expression.");
                return Empty(syntax);
        }
    }

    /// <summary>The left side of an assignment to a class member: a field, or a property through its Let or Set (MS-VBAL 5.4.3.8, 5.4.3.9).</summary>
    private BoundExpression BindClassTarget(BoundExpression receiver, ModuleSymbol classModule, string name, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax, bool isSet)
    {
        if (!classModule.Members.TryGetValue(name, out var symbol) || !IsVisible(symbol, classModule))
        {
            Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{name}'.");
            return Empty(syntax);
        }

        switch (symbol)
        {
            case PropertySymbol property:
                return BindPropertyTarget(property, syntax, arguments, isSet, receiver);
            case VariableSymbol { IsArray: true } array when arguments.Count > 0:
                return BindElement(new BoundField(syntax, receiver, array), arguments, syntax);
            case VariableSymbol field when arguments.Count == 0:
                return new BoundField(syntax, receiver, field);
            default:
                Report(DiagnosticIds.ExpectedVariable, syntax.FirstToken()!, $"'{name}' cannot be assigned to.");
                return Empty(syntax);
        }
    }

    /// <summary>A member is reachable from outside its class when it is Public; inside, everything is.</summary>
    private bool IsVisible(Symbol symbol, ModuleSymbol owner) => ReferenceEquals(owner, module) || IsPublic(symbol);

    /// <summary>
    /// The value an object of a class type stands for where a value is needed (MS-VBAL 5.6.9.3):
    /// its default member. The runtime resolves it through the object's dispatch, so the binder
    /// only needs the declared type of the result.
    /// </summary>
    private static VbaType DefaultMemberType(ModuleSymbol classModule) => classModule.DefaultMember switch
    {
        PropertySymbol { Get.ReturnType: { } type } => type,
        ProcedureSymbol { ReturnType: { } type } => type,
        VariableSymbol variable => variable.Type,
        _ => VbaType.Variant,
    };

    /// <summary>
    /// obj(arguments) on a class instance: the default member with those arguments (MS-VBAL 5.6.13.2).
    /// </summary>
    private BoundExpression BindClassDefault(BoundExpression receiver, ModuleSymbol classModule, IReadOnlyList<ArgumentSyntax> arguments, SyntaxNode syntax)
    {
        if (classModule.DefaultMember is null)
        {
            Report(DiagnosticIds.VariableNotDefined, syntax.FirstToken()!, $"Method or data member not found: '{classModule.Name}' has no default member.");
            return Empty(syntax);
        }

        return BindClassMember(receiver, classModule, classModule.DefaultMember.Name, arguments, syntax, isCallStatement: false);
    }

    /// <summary>TypeOf expression Is type (MS-VBAL 5.6.9.10): a run-time test of the object's class.</summary>
    private BoundExpression BindTypeOf(TypeOfExpressionSyntax typeOf)
    {
        var value = Convert(BindExpression(typeOf.Expression), VbaType.Variant);
        var type = BindNamedType(typeOf.Type);
        if (!type.IsObject)
        {
            Report(DiagnosticIds.TypeMismatch, typeOf.Type.FirstToken()!, "Expected: class name after Is.");
            return Empty(typeOf);
        }

        return new BoundTypeOf(typeOf, value, type);
    }

    /// <summary>RaiseEvent name(arguments) (MS-VBAL 5.4.2.10): every object that handles the event through a WithEvents variable runs its handler, in subscription order.</summary>
    private BoundStatement BindRaiseEvent(RaiseEventStatementSyntax statement)
    {
        var name = statement.Name.NameValue;
        var symbol = module.Events.FirstOrDefault(e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (symbol is null)
        {
            Report(DiagnosticIds.ProcedureNotDefined, statement.Name, $"Event not defined: '{name}'.");
            return new BoundNop(statement);
        }

        var arguments = statement.Arguments?.Arguments.ToList() ?? [];
        if (arguments.Count != symbol.Parameters.Count)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, statement.Name, $"Wrong number of arguments or invalid property assignment: '{name}'.");
        }

        var values = new List<BoundExpression>();
        for (var i = 0; i < arguments.Count && i < symbol.Parameters.Count; i++)
        {
            var parameter = symbol.Parameters[i];
            var argument = arguments[i];
            var bound = argument.IsOmitted ? new BoundLiteral(statement, Variant.Missing, VbaType.Variant) : BindExpression(argument.Expression!);

            // A ByRef event parameter is written back into the caller's variable when a handler changes it (Classes golden).
            var writeBack = !parameter.IsByVal && argument.ByValKeyword is null && bound.IsLValue;
            values.Add(writeBack ? bound : Convert(bound, parameter.Type));
        }

        return new BoundRaiseEvent(statement, symbol, values);
    }
}
