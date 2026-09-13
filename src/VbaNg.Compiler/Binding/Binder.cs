using System.Globalization;

using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// The binder (ARCHITECTURE.md section 4, step 3): declares every module's members, resolves
/// names and types, and produces the typed bound tree the emitter compiles. Binding happens in
/// two passes, declarations for every module first, so procedures can call each other in any
/// order, then bodies. Every semantic rule cites its MS-VBAL section.
/// </summary>
public sealed partial class Binder
{
    private readonly ProjectSymbol project;
    private readonly List<Diagnostic> diagnostics;
    private readonly Dictionary<string, (ModuleSymbol Module, ConstDeclaratorSyntax Syntax)> pendingConstants = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<ConstDeclaratorSyntax> foldingConstants = [];
    private readonly List<(ModuleSymbol Module, ImplementsStatementSyntax Syntax)> pendingImplements = [];

    /// <summary>The member access currently being bound as the left side of an assignment, so a class member binds to its Let or Set (MS-VBAL 5.4.3.8).</summary>
    private MemberAccessExpressionSyntax? assignmentTarget;

    private bool assignmentIsSet;

    private ModuleSymbol module = null!;
    private ProcedureContext procedure = null!;

    /// <summary>The manifest references the vbang library, which brings the Assert module (ARCHITECTURE.md section 8).</summary>
    private readonly bool vbangReference;

    /// <summary>The type libraries the manifest references, for early binding (ARCHITECTURE.md section 6).</summary>
    private readonly ReferencedLibraries libraries;

    private readonly ProjectManifest manifest;

    private Binder(string projectName, List<Diagnostic> diagnostics, ProjectManifest manifest, ReferencedLibraries libraries)
    {
        project = new ProjectSymbol(projectName);
        this.diagnostics = diagnostics;
        this.libraries = libraries;
        this.manifest = manifest;
        vbangReference = manifest.HasReference(ProjectManifest.VbangLibrary);
    }

    /// <summary>Binds a project; returns null when any error was reported.</summary>
    public static BoundProject? Bind(string projectName, IReadOnlyList<SyntaxTree> trees, List<Diagnostic> diagnostics, ProjectManifest? manifest = null, IReadOnlyList<Runtime.TypeLibraries.ComLibrary>? libraries = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectName);
        ArgumentNullException.ThrowIfNull(trees);
        ArgumentNullException.ThrowIfNull(diagnostics);

        var binder = new Binder(projectName, diagnostics, manifest ?? ProjectManifest.Empty, libraries is null ? ReferencedLibraries.None : new ReferencedLibraries(libraries));
        var before = diagnostics.Count(d => d.IsError);
        foreach (var tree in trees)
        {
            binder.DeclareModule(tree);
        }

        binder.ResolveDeclarations();
        var modules = new List<BoundModule>();
        foreach (var moduleSymbol in binder.project.Modules)
        {
            modules.Add(binder.BindModule(moduleSymbol));
        }

        return diagnostics.Count(d => d.IsError) > before ? null : new BoundProject(binder.project, modules);
    }

    // Declarations (MS-VBAL 5.2).

    private void DeclareModule(SyntaxTree tree)
    {
        var name = tree.Root.Name ?? Path.GetFileNameWithoutExtension(tree.FilePath);
        module = new ModuleSymbol(name, tree) { EmitName = CSharpNames.Identifier(name) };
        if (project.FindModule(name) is not null)
        {
            Report(DiagnosticIds.AmbiguousName, tree.Root.FirstToken() ?? tree.Root.EndOfFileToken, $"Ambiguous name detected: '{name}' (two modules share the name).");
        }

        project.Modules.Add(module);
        DeclareModuleKind(tree);
        foreach (var member in tree.Root.Members)
        {
            switch (member)
            {
                case AttributeStatementSyntax:
                    break;
                case OptionStatementSyntax option:
                    DeclareOption(option);
                    break;
                case DefTypeStatementSyntax defType:
                    DeclareDefType(defType);
                    break;
                case VariableDeclarationSyntax variables:
                    DeclareModuleVariables(variables);
                    break;
                case ConstDeclarationSyntax constants:
                    foreach (var declarator in constants.Declarators)
                    {
                        pendingConstants[module.Name + "." + declarator.Name.NameValue] = (module, declarator);
                    }

                    break;
                case TypeDefinitionSyntax type:
                    DeclareRecord(type);
                    break;
                case EnumDefinitionSyntax enumeration:
                    DeclareEnum(enumeration);
                    break;
                case ProcedureDeclarationSyntax procedureSyntax:
                    DeclareProcedure(procedureSyntax);
                    break;
                case DeclareStatementSyntax declare:
                    DeclareExternal(declare);
                    break;
                case EventDeclarationSyntax declaration:
                    DeclareEvent(declaration);
                    break;
                case ImplementsStatementSyntax implements:
                    DeclareImplements(implements);
                    break;
                case LabelStatementSyntax or BadStatementSyntax:
                    break;
                default:
                    Report(DiagnosticIds.SyntaxError, member.FirstToken()!, "Only declarations are allowed outside procedures.");
                    break;
            }
        }
    }

    private void DeclareOption(OptionStatementSyntax option)
    {
        switch (option.OptionName.Text.ToUpperInvariant())
        {
            case "EXPLICIT":
                module.Options.Explicit = true;
                break;
            case "BASE":
                module.Options.Base = option.Argument?.Text == "1" ? 1 : 0;
                break;
            case "COMPARE":
                module.Options.Compare = option.Argument is { } mode && mode.Text.Equals("Text", StringComparison.OrdinalIgnoreCase) ? CompareMode.Text : CompareMode.Binary;
                break;
            case "PRIVATE":
                module.Options.PrivateModule = true;
                break;
        }
    }

    /// <summary>DefInt A-Z and friends (MS-VBAL 5.2.2): undeclared names starting with those letters take the type.</summary>
    private void DeclareDefType(DefTypeStatementSyntax defType)
    {
        var type = defType.DefKeyword.Kind switch
        {
            SyntaxKind.DefBoolKeyword => VbaType.Boolean,
            SyntaxKind.DefByteKeyword => VbaType.Byte,
            SyntaxKind.DefIntKeyword => VbaType.Integer,
            SyntaxKind.DefLngKeyword => VbaType.Long,
            SyntaxKind.DefLngLngKeyword or SyntaxKind.DefLngPtrKeyword => VbaType.LongLong,
            SyntaxKind.DefCurKeyword => VbaType.Currency,
            SyntaxKind.DefSngKeyword => VbaType.Single,
            SyntaxKind.DefDblKeyword => VbaType.Double,
            SyntaxKind.DefDateKeyword => VbaType.Date,
            SyntaxKind.DefStrKeyword => VbaType.String,
            SyntaxKind.DefObjKeyword => VbaType.Object,
            _ => VbaType.Variant,
        };

        foreach (var range in defType.Ranges)
        {
            var first = char.ToUpperInvariant(range.First.Text[0]);
            var last = range.Last is null ? first : char.ToUpperInvariant(range.Last.Text[0]);
            for (var letter = first; letter <= last; letter++)
            {
                module.DefTypes[letter] = type;
            }
        }
    }

    private void DeclareModuleVariables(VariableDeclarationSyntax declaration)
    {
        var isPublic = declaration.Modifiers.Any(SyntaxKind.PublicKeyword) || declaration.Modifiers.Any(SyntaxKind.GlobalKeyword);
        foreach (var declarator in declaration.Declarators)
        {
            var variable = new VariableSymbol(declarator.Name.NameValue, VbaType.Variant, VariableKind.Module)
            {
                IsPublic = isPublic,
                IsWithEvents = declarator.WithEventsKeyword is not null,
                Module = module,
                Syntax = declarator,
                EmitName = MemberEmitName(declarator.Name.NameValue),
            };
            if (AddMember(variable, declarator.Name))
            {
                module.Variables.Add(variable);
            }
        }
    }

    private void DeclareRecord(TypeDefinitionSyntax type)
    {
        var record = new RecordSymbol(type.Name.NameValue, module, type, MemberEmitName(type.Name.NameValue))
        {
            IsPublic = !type.Modifiers.Any(SyntaxKind.PrivateKeyword),
        };
        if (AddMember(record, type.Name))
        {
            module.Records.Add(record);
        }
    }

    private void DeclareEnum(EnumDefinitionSyntax enumeration)
    {
        var symbol = new EnumSymbol(enumeration.Name.NameValue, module, enumeration)
        {
            IsPublic = !enumeration.Modifiers.Any(SyntaxKind.PrivateKeyword),
            EmitName = MemberEmitName(enumeration.Name.NameValue),
        };
        if (AddMember(symbol, enumeration.Name))
        {
            module.Enums.Add(symbol);
        }
    }

    /// <summary>A file with the VBE's class header is a class module (MS-VBAL 4.2); see <see cref="DeclareClassModule"/> for the kinds.</summary>
    private void DeclareModuleKind(SyntaxTree tree)
    {
        // A form is a class module, but its file carries the form designer's own header where a
        // .cls carries the class one, and the importer writes neither; the extension says so
        // instead (ROADMAP.md D-D).
        if (tree.Root.Header is null && !IsForm(tree))
        {
            return;
        }

        DeclareClassModule(tree, Flag(tree, "VB_PredeclaredId"), Flag(tree, "VB_Exposed"));
    }

    /// <summary>Whether the module is a UserForm, which the alpha compiles as a class without its controls (ROADMAP.md D-D).</summary>
    private static bool IsForm(SyntaxTree tree) =>
        tree.FilePath.EndsWith(".frm", StringComparison.OrdinalIgnoreCase);

    /// <summary>The value of a module-level boolean attribute (MS-VBAL 5.2.4.1.1).</summary>
    private static bool Flag(SyntaxTree tree, string name) =>
        tree.Root.Members.OfType<AttributeStatementSyntax>()
            .Where(a => a.Name is IdentifierNameSyntax { Name: var n } && n.Equals(name, StringComparison.OrdinalIgnoreCase))
            .Select(a => a.Values.Nodes.FirstOrDefault() is LiteralExpressionSyntax { Token: var token } && (token.Kind == SyntaxKind.TrueKeyword || token.Value is true))
            .FirstOrDefault();

    /// <summary>
    /// Worksheet_Change, Workbook_Open, CommandButton1_Click in a document module: the part before
    /// the last underscore names the source, the module's own object when it is the document kind,
    /// otherwise a control on it; the host resolves both at load (ARCHITECTURE.md section 6, "Events").
    /// </summary>
    private void DeclareEventHandler(ProcedureSymbol symbol)
    {
        if (module.Kind != ModuleKind.Document || symbol.Kind != ProcedureKind.Sub)
        {
            return;
        }

        var split = symbol.Name.LastIndexOf('_');
        if (split <= 0 || split == symbol.Name.Length - 1)
        {
            return;
        }

        var source = symbol.Name[..split];
        var eventName = symbol.Name[(split + 1)..];
        if (source.Equals(module.DocumentKind, StringComparison.OrdinalIgnoreCase) && !IsDocumentEvent(eventName))
        {
            return;
        }

        symbol.EventSource = source;
        symbol.EventName = eventName;
    }

    /// <summary>True when the document kind's default source interface has the event, or when no Excel library is referenced to check.</summary>
    private bool IsDocumentEvent(string eventName)
    {
        var coclass = libraries.ResolveType("Excel", module.DocumentKind!)?.ComType;
        var source = coclass?.DefaultSource is { } name ? libraries.FindQualified(name) : null;
        return source is null || source.FindMembers(eventName).Any();
    }

    /// <summary>
    /// The C# name of a module-level member: the escaped name, suffixed when it matches the module's
    /// class name, which C# forbids (CS0542) while VBA allows Sub Main in module Main.
    /// </summary>
    private string MemberEmitName(string name)
    {
        var emitName = CSharpNames.Identifier(name);
        if (emitName != module.EmitName)
        {
            return emitName;
        }

        var candidate = emitName + "_";
        while (module.Members.ContainsKey(candidate))
        {
            candidate += "_";
        }

        return candidate;
    }

    /// <summary>A comment whose first word is <c>@Test</c>, written as <c>'@Test</c> or <c>Rem @Test</c>; anything after the word is free text.</summary>
    private static bool IsTestAnnotation(string comment)
    {
        var body = comment.StartsWith('\'') ? comment[1..] : comment.Length >= 3 ? comment[3..] : string.Empty;
        var trimmed = body.TrimStart();
        var end = trimmed.IndexOfAny([' ', '\t']);
        var word = end < 0 ? trimmed : trimmed[..end];
        return word.Equals("@Test", StringComparison.OrdinalIgnoreCase);
    }

    private void DeclareProcedure(ProcedureDeclarationSyntax syntax)
    {
        var name = syntax.Name.NameValue;
        var kind = syntax.AccessorKeyword?.Kind switch
        {
            SyntaxKind.GetKeyword => ProcedureKind.PropertyGet,
            SyntaxKind.LetKeyword => ProcedureKind.PropertyLet,
            SyntaxKind.SetKeyword => ProcedureKind.PropertySet,
            _ => syntax.IsFunction ? ProcedureKind.Function : ProcedureKind.Sub,
        };

        var symbol = new ProcedureSymbol(name, kind, module, syntax)
        {
            IsPublic = !syntax.Modifiers.Any(SyntaxKind.PrivateKeyword),
            IsStatic = syntax.Modifiers.Any(SyntaxKind.StaticKeyword),
        };
        symbol.EmitName = kind switch
        {
            ProcedureKind.PropertyGet => "Get_" + name,
            ProcedureKind.PropertyLet => "Let_" + name,
            ProcedureKind.PropertySet => "Set_" + name,
            _ => MemberEmitName(name),
        };

        DeclareEventHandler(symbol);
        if (syntax.FirstToken()!.LeadingTrivia.Any(t => t.Kind == SyntaxKind.CommentTrivia && IsTestAnnotation(t.Text)))
        {
            // '@Test marks a test procedure (ARCHITECTURE.md section 8); the runner calls it with no arguments and no result.
            var parameterCount = syntax.Parameters?.Parameters.Count ?? 0;
            if (kind == ProcedureKind.Sub && symbol.IsPublic && parameterCount == 0)
            {
                symbol.IsTest = true;
            }
            else
            {
                Report(DiagnosticIds.InvalidTestProcedure, syntax.Name, $"'@Test applies only to a Public Sub without parameters: '{name}'.");
            }
        }

        if (syntax.IsProperty)
        {
            if (!module.Members.TryGetValue(name, out var existing))
            {
                existing = new PropertySymbol(name, module);
                module.Members[name] = existing;
            }

            if (existing is not PropertySymbol property)
            {
                Report(DiagnosticIds.AmbiguousName, syntax.Name, $"Ambiguous name detected: '{name}'.");
                return;
            }

            var slot = kind switch
            {
                ProcedureKind.PropertyGet => property.Get,
                ProcedureKind.PropertyLet => property.Let,
                _ => property.Set,
            };
            if (slot is not null)
            {
                Report(DiagnosticIds.AmbiguousName, syntax.Name, $"Ambiguous name detected: '{name}'.");
                return;
            }

            switch (kind)
            {
                case ProcedureKind.PropertyGet:
                    property.Get = symbol;
                    break;
                case ProcedureKind.PropertyLet:
                    property.Let = symbol;
                    break;
                default:
                    property.Set = symbol;
                    break;
            }
        }
        else if (!AddMember(symbol, syntax.Name))
        {
            return;
        }

        module.Procedures.Add(symbol);
    }

    private bool AddMember(Symbol symbol, SyntaxToken nameToken)
    {
        if (module.Members.TryAdd(symbol.Name, symbol))
        {
            return true;
        }

        Report(DiagnosticIds.AmbiguousName, nameToken, $"Ambiguous name detected: '{symbol.Name}'.");
        return false;
    }

    /// <summary>Second declaration pass: types, constant values, record fields, enum members, and signatures, once every name exists.</summary>
    private void ResolveDeclarations()
    {
        // Enums before records: a record's array bounds may name enum members (Procedures golden), never the reverse.
        foreach (var moduleSymbol in project.Modules)
        {
            module = moduleSymbol;
            foreach (var enumeration in module.Enums)
            {
                ResolveEnum(enumeration);
            }
        }

        foreach (var moduleSymbol in project.Modules)
        {
            module = moduleSymbol;
            foreach (var record in module.Records)
            {
                ResolveRecord(record);
            }
        }

        foreach (var (module, declarator) in pendingConstants.Values.ToList())
        {
            this.module = module;
            FoldModuleConstant(declarator);
        }

        foreach (var moduleSymbol in project.Modules)
        {
            module = moduleSymbol;
            for (var i = 0; i < module.Variables.Count; i++)
            {
                ResolveVariable(module.Variables[i], (VariableDeclaratorSyntax)module.Variables[i].Syntax!);
            }

            foreach (var procedureSymbol in module.Procedures)
            {
                ResolveSignature(procedureSymbol);
            }

            foreach (var external in module.Externals)
            {
                ResolveExternalSignature(external);
            }
        }

        ResolveClasses();
    }

    private void ResolveRecord(RecordSymbol record)
    {
        foreach (var member in record.Syntax.Members)
        {
            if (member is not TypeMemberSyntax field)
            {
                continue;
            }

            var type = BindType(field.AsClause, field.Name, isArray: field.Bounds is not null);
            var symbol = new VariableSymbol(field.Name.NameValue, type, VariableKind.Field)
            {
                Module = module,
                Syntax = field,
                EmitName = CSharpNames.Identifier(field.Name.NameValue),
            };
            if (field.Bounds is { Bounds.Count: > 0 } bounds)
            {
                symbol.Bounds = ConstantBounds(bounds);
            }

            if (record.FindField(symbol.Name) is not null)
            {
                Report(DiagnosticIds.AmbiguousName, field.Name, $"Ambiguous name detected: '{symbol.Name}'.");
                continue;
            }

            record.Fields.Add(symbol);
        }
    }

    private bool IsEnumMember(Symbol symbol) => symbol is ConstantSymbol constant && module.Enums.Any(e => e.Members.Contains(constant));

    /// <summary>Enum members (MS-VBAL 5.2.3.4): an omitted value is the previous member plus one, starting at 0.</summary>
    private void ResolveEnum(EnumSymbol enumeration)
    {
        long next = 0;
        foreach (var member in enumeration.Syntax.Members)
        {
            if (member is not EnumMemberSyntax item)
            {
                continue;
            }

            var value = next;
            if (item.Value is not null)
            {
                var folded = FoldConstant(item.Value);
                if (folded is { } constant)
                {
                    value = Coerce.ToInt64(constant);
                }
            }

            var symbol = new ConstantSymbol(item.Name.NameValue, enumeration.Type, Variant.FromInt32(unchecked((int)value)))
            {
                IsPublic = enumeration.IsPublic,
                Module = module,
            };
            enumeration.Members.Add(symbol);
            if (!module.Members.TryAdd(symbol.Name, symbol) && !IsEnumMember(module.Members[symbol.Name]))
            {
                // Two enums of one module may share a member name (Procedures golden); only a clash with another kind
                // of member is an error. The first enum keeps the bare name, and a qualified use reaches either.
                Report(DiagnosticIds.AmbiguousName, item.Name, $"Ambiguous name detected: '{symbol.Name}'.");
            }

            next = value + 1;
        }
    }

    /// <summary>Folds a module-level Const, on demand when another constant refers to it (MS-VBAL 5.2.3.2).</summary>
    private ConstantSymbol? FoldModuleConstant(ConstDeclaratorSyntax declarator)
    {
        var name = declarator.Name.NameValue;
        if (module.Members.TryGetValue(name, out var existing))
        {
            return existing as ConstantSymbol;
        }

        if (!foldingConstants.Add(declarator))
        {
            Report(DiagnosticIds.ConstantExpressionRequired, declarator.Name, $"Constant '{name}' refers to itself.");
            return null;
        }

        var value = FoldConstant(declarator.Value) ?? Variant.Empty;
        var type = declarator.AsClause is null
            ? (declarator.Name.TypeSuffix is var suffix && suffix != '\0' ? VbaType.FromSuffix(suffix)! : VbaType.FromVarType(value.Type))
            : BindType(declarator.AsClause, declarator.Name, isArray: false);
        var symbol = new ConstantSymbol(name, type, ConvertConstant(value, type, declarator.Name))
        {
            IsPublic = ((ConstDeclarationSyntax)FindParent(declarator)).Modifiers.Any(SyntaxKind.PublicKeyword) || ((ConstDeclarationSyntax)FindParent(declarator)).Modifiers.Any(SyntaxKind.GlobalKeyword),
            Module = module,
        };
        foldingConstants.Remove(declarator);
        pendingConstants.Remove(module.Name + "." + name);
        if (AddMember(symbol, declarator.Name))
        {
            module.Constants.Add(symbol);
        }

        return symbol;
    }

    private ConstDeclarationSyntax FindParent(ConstDeclaratorSyntax declarator) =>
        module.Tree.Root.Members.OfType<ConstDeclarationSyntax>().First(c => c.Declarators.Contains(declarator));

    private Variant ConvertConstant(Variant value, VbaType type, SyntaxToken at)
    {
        if (type.IsVariant || value.IsEmpty)
        {
            return value;
        }

        try
        {
            return Coerce.ToType(value, type.VarType);
        }
        catch (VbaException ex)
        {
            Report(DiagnosticIds.TypeMismatch, at, ex.Description + ".");
            return value;
        }
    }

    private void ResolveVariable(VariableSymbol variable, VariableDeclaratorSyntax declarator)
    {
        var resolved = DeclareVariableType(declarator, variable.Kind, variable.IsPublic, variable.Module, variable.Procedure);
        variable.Module!.Variables[variable.Module.Variables.IndexOf(variable)] = resolved;
        variable.Module.Members[variable.Name] = resolved;
    }

    /// <summary>Builds the variable a declarator declares: its type from the As clause, suffix, or DefType, and its array bounds.</summary>
    private VariableSymbol DeclareVariableType(VariableDeclaratorSyntax declarator, VariableKind kind, bool isPublic, ModuleSymbol? owner, ProcedureSymbol? ownerProcedure)
    {
        var name = declarator.Name.NameValue;
        var isArray = declarator.Bounds is not null || declarator.AsClause?.ArrayDesignator is not null;
        var type = declarator.AsClause is null
            ? DefaultType(declarator.Name)
            : BindType(declarator.AsClause, declarator.Name, isArray: false);
        if (isArray)
        {
            type = VbaType.ArrayOf(type);
        }

        var variable = new VariableSymbol(name, type, kind)
        {
            IsPublic = isPublic,
            Module = owner,
            Procedure = ownerProcedure,
            Syntax = declarator,
            IsNew = declarator.AsClause?.NewKeyword is not null,
            IsWithEvents = declarator.WithEventsKeyword is not null,
            EmitName = kind == VariableKind.Module ? MemberEmitName(name) : CSharpNames.Identifier(name),
        };
        if (variable.IsWithEvents)
        {
            // The handlers of a WithEvents source are members of the object module that declares it (MS-VBAL 5.2.3.1.4).
            if (owner?.Kind is not (ModuleKind.Class or ModuleKind.Document))
            {
                Report(DiagnosticIds.InvalidStatementPlacement, declarator.WithEventsKeyword!, "Only valid in object module.");
            }
            else if (type.ProjectClass is not null)
            {
                owner.EventSources.Add(variable);
            }
            else if (type.IsCom && type.ComType?.DefaultSource is { } sourceName && libraries.FindQualified(sourceName) is { } source)
            {
                // A library type sources the events of its coclass's default source interface (ARCHITECTURE.md section 6, "Events").
                variable.EventInterface = source;
                owner.EventSources.Add(variable);
            }
            else if (!type.IsObject || type.ComType is not null)
            {
                // A type the binder could not resolve was reported already (VBA0012) and binds as Object; no second error for it.
                Report(DiagnosticIds.TypeMismatch, declarator.Name, "Object does not source automation events.");
            }
        }
        if (declarator.Bounds is { Bounds.Count: > 0 } bounds)
        {
            variable.Bounds = ConstantBounds(bounds);
        }

        if (variable.IsNew && !type.IsObject)
        {
            Report(DiagnosticIds.TypeMismatch, declarator.AsClause!.NewKeyword!, "Expected: class name after New.");
        }

        return variable;
    }

    /// <summary>The type of a name without an As clause: its type suffix, else the DefType of its first letter, else Variant.</summary>
    private VbaType DefaultType(SyntaxToken name)
    {
        var suffix = name.TypeSuffix;
        if (suffix != '\0')
        {
            return VbaType.FromSuffix(suffix)!;
        }

        return module.DefTypes.TryGetValue(char.ToUpperInvariant(name.NameValue[0]), out var type) ? type : VbaType.Variant;
    }

    private (int Lower, int Upper)[] ConstantBounds(ArrayBoundsSyntax bounds)
    {
        var result = new (int, int)[bounds.Bounds.Count];
        for (var i = 0; i < result.Length; i++)
        {
            var bound = bounds.Bounds[i];
            var lower = bound.Lower is null ? module.Options.Base : ConstantInt(bound.Lower);
            var upper = ConstantInt(bound.Upper);
            result[i] = (lower, upper);
        }

        return result;
    }

    private int ConstantInt(ExpressionSyntax expression)
    {
        var value = FoldConstant(expression);
        if (value is null)
        {
            return 0;
        }

        try
        {
            return Coerce.ToInt32(value.Value);
        }
        catch (VbaException ex)
        {
            Report(DiagnosticIds.TypeMismatch, expression.FirstToken()!, ex.Description + ".");
            return 0;
        }
    }

    private void ResolveSignature(ProcedureSymbol symbol)
    {
        var syntax = symbol.Syntax;
        if (!symbol.IsExternal)
        {
            foreach (var parameterSyntax in syntax.Parameters?.Parameters.ToList() ?? [])
            {
                if (parameterSyntax.AsClause?.Type is BuiltinTypeSyntax { Keyword.Kind: SyntaxKind.AnyKeyword } any)
                {
                    Report(DiagnosticIds.TypeNotDefined, any.Keyword, "User-defined type not defined: 'Any' is only valid in a Declare.");
                }
            }
        }

        if (syntax.Parameters is not null)
        {
            var index = 0;
            foreach (var parameterSyntax in syntax.Parameters.Parameters)
            {
                var parameter = BindParameter(parameterSyntax, index++);
                if (symbol.Parameters.Any(p => p.Name.Equals(parameter.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    Report(DiagnosticIds.AmbiguousName, parameterSyntax.Name, $"Ambiguous name detected: '{parameter.Name}'.");
                }

                parameter.Procedure = symbol;
                symbol.Parameters.Add(parameter);
            }
        }

        if (symbol.ReturnsValue)
        {
            var returnType = syntax.AsClause is null ? DefaultType(syntax.Name) : BindType(syntax.AsClause, syntax.Name, isArray: false);
            if (syntax.AsClause?.ArrayDesignator is not null)
            {
                returnType = VbaType.ArrayOf(returnType);
            }

            symbol.ReturnType = returnType;
            symbol.Result = new VariableSymbol(symbol.Name, returnType, VariableKind.Result)
            {
                Module = module,
                Procedure = symbol,
                EmitName = "__result",
            };
        }
    }

    /// <summary>[Optional] [ByVal | ByRef] [ParamArray] name [()] [As type] [= default] (MS-VBAL 5.3.1.5).</summary>
    private ParameterSymbol BindParameter(ParameterSyntax syntax, int index)
    {
        var isArray = syntax.Bounds is not null || syntax.AsClause?.ArrayDesignator is not null;
        var isParamArray = syntax.Modifiers.Any(SyntaxKind.ParamArrayKeyword);
        var type = syntax.AsClause is null ? DefaultType(syntax.Name) : BindType(syntax.AsClause, syntax.Name, isArray: false);
        if (isArray || isParamArray)
        {
            type = VbaType.ArrayOf(isParamArray ? VbaType.Variant : type);
        }

        var parameter = new ParameterSymbol(syntax.Name.NameValue, type)
        {
            IsByVal = syntax.Modifiers.Any(SyntaxKind.ByValKeyword),
            IsOptional = syntax.Modifiers.Any(SyntaxKind.OptionalKeyword),
            IsParamArray = isParamArray,
            Index = index,
            Module = module,
            Syntax = syntax,
            EmitName = CSharpNames.Identifier(syntax.Name.NameValue),
        };
        parameter.EmitByRef = !parameter.IsByVal && !parameter.IsParamArray;

        if (syntax.DefaultValue is LiteralExpressionSyntax { Token.Kind: SyntaxKind.NothingKeyword } && (type.IsObject || type.IsVariant))
        {
            // Optional o As Class = Nothing: the default is Nothing, which is the type's own default (MS-VBAL 5.3.1.5);
            // a Variant's is a Variant holding Nothing (Procedures golden).
            if (type.IsVariant)
            {
                parameter.Default = Runtime.Variant.Nothing;
            }
        }
        else if (syntax.DefaultValue is not null)
        {
            var value = FoldConstant(syntax.DefaultValue);
            if (value is { } constant)
            {
                parameter.Default = type.IsVariant ? constant : ConvertConstant(constant, type, syntax.Name);
            }
        }

        if (parameter.IsByVal && (type.IsArray || type.IsRecord))
        {
            Report(DiagnosticIds.TypeMismatch, syntax.Name, "Array and user-defined type arguments must be passed ByRef.");
        }

        return parameter;
    }

    /// <summary>Resolves an As clause (MS-VBAL 5.6.16.7 type-expression) to a declared type.</summary>
    private VbaType BindType(AsClauseSyntax? asClause, SyntaxToken nameToken, bool isArray)
    {
        if (asClause is null)
        {
            return DefaultType(nameToken);
        }

        var type = BindTypeSyntax(asClause.Type);
        return isArray ? VbaType.ArrayOf(type) : type;
    }

    private VbaType BindTypeSyntax(TypeSyntax typeSyntax)
    {
        switch (typeSyntax)
        {
            case BuiltinTypeSyntax builtin:
                {
                    var type = VbaType.FromKeyword(builtin.Keyword.Kind);
                    if (type is null)
                    {
                        if (builtin.Keyword.Kind == SyntaxKind.AnyKeyword)
                        {
                            // As Any (MS-VBAL 5.2.3.5): only a Declare's parameter may have it, which ResolveSignature checks.
                            return VbaType.Any;
                        }

                        if (builtin.Keyword.Kind == SyntaxKind.DecimalKeyword)
                        {
                            Report(DiagnosticIds.TypeMismatch, builtin.Keyword, "Decimal values exist only inside Variants; declare the variable As Variant.");
                        }
                        else
                        {
                            Report(DiagnosticIds.TypeNotDefined, builtin.Keyword, $"User-defined type not defined: '{builtin.Keyword.Text}'.");
                        }

                        return VbaType.Variant;
                    }

                    if (builtin.AsteriskToken is not null && builtin.Length is not null)
                    {
                        return VbaType.FixedString(ConstantInt(builtin.Length));
                    }

                    return type;
                }

            case NamedTypeSyntax named:
                return BindNamedType(named.Name);
            default:
                Report(DiagnosticIds.TypeNotDefined, typeSyntax.FirstToken()!, "User-defined type not defined.");
                return VbaType.Variant;
        }
    }

    private VbaType BindNamedType(ExpressionSyntax name)
    {
        switch (name)
        {
            case IdentifierNameSyntax simple:
                return LookupType(simple.Name, null, simple.Identifier);
            case MemberAccessExpressionSyntax { Expression: IdentifierNameSyntax qualifier } qualified:
                return LookupType(qualified.Name.Name, qualifier.Name, qualified.Name.Identifier);
            default:
                Report(DiagnosticIds.TypeNotDefined, name.FirstToken()!, "User-defined type not defined: '" + name.ToFullString().Trim() + "'.");
                return VbaType.Variant;
        }
    }

    /// <summary>Type names resolve to this module's records and enums, then other modules' public ones, then the runtime's classes.</summary>
    private VbaType LookupType(string name, string? qualifier, SyntaxToken at)
    {
        if (name.Equals("Object", StringComparison.OrdinalIgnoreCase))
        {
            return VbaType.Object;
        }

        if (name.Equals("Any", StringComparison.OrdinalIgnoreCase) && qualifier is null)
        {
            // As Any is only meaningful in a Declare; ResolveSignature rejects it anywhere else.
            return VbaType.Any;
        }

        if (StandardLibrary.IsEnumName(name) && (qualifier is null || qualifier.Equals("VBA", StringComparison.OrdinalIgnoreCase)))
        {
            // The enums of the VBA library (VbVarType, VbMsgBoxStyle, and the rest) are Long (MS-VBAL 6.1.1).
            return VbaType.Long;
        }

        if (qualifier is null || qualifier.Equals(module.Name, StringComparison.OrdinalIgnoreCase))
        {
            if (module.Members.TryGetValue(name, out var own))
            {
                if (own is RecordSymbol record)
                {
                    return record.Type;
                }

                if (own is EnumSymbol enumeration)
                {
                    return enumeration.Type;
                }
            }
        }

        foreach (var other in project.Modules)
        {
            if (ReferenceEquals(other, module) || (qualifier is not null && !qualifier.Equals(other.Name, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            if (other.Members.TryGetValue(name, out var found))
            {
                if (found is RecordSymbol { IsPublic: true } record)
                {
                    return record.Type;
                }

                if (found is EnumSymbol { IsPublic: true } enumeration)
                {
                    return enumeration.Type;
                }
            }
        }

        // A class module of this project names a type (MS-VBAL 5.2.4.1.3).
        if (project.FindModule(name) is { Kind: ModuleKind.Class, ClassType: { } classType } && (qualifier is null || qualifier.Equals(project.Name, StringComparison.OrdinalIgnoreCase)))
        {
            return classType;
        }

        if (qualifier is null || qualifier.Equals("VBA", StringComparison.OrdinalIgnoreCase))
        {
            if (name.Equals("Collection", StringComparison.OrdinalIgnoreCase))
            {
                return VbaType.Collection;
            }

            if (name.Equals("ErrObject", StringComparison.OrdinalIgnoreCase))
            {
                return VbaType.ErrObject;
            }
        }

        // stdole is referenced by every VBA project; its object interfaces name any object (MS-VBAL 5.6.16.7).
        if ((name.Equals("IUnknown", StringComparison.OrdinalIgnoreCase) || name.Equals("IDispatch", StringComparison.OrdinalIgnoreCase))
            && (qualifier is null || qualifier.Equals("stdole", StringComparison.OrdinalIgnoreCase)))
        {
            return VbaType.Object;
        }

        // Referenced type libraries, in manifest order (MS-VBAL 5.6.16.7, the referenced library step).
        if (!string.Equals(qualifier, "VBA", StringComparison.OrdinalIgnoreCase) && libraries.ResolveType(qualifier, name) is { } libraryType)
        {
            return libraryType;
        }

        Report(DiagnosticIds.TypeNotDefined, at, $"User-defined type not defined: '{(qualifier is null ? name : qualifier + "." + name)}'.");
        return VbaType.Variant;
    }

    // Procedure bodies (MS-VBAL 5.3, 5.4).

    private BoundModule BindModule(ModuleSymbol moduleSymbol)
    {
        module = moduleSymbol;
        var procedures = new List<BoundProcedure>();
        foreach (var symbol in module.Procedures)
        {
            procedures.Add(BindProcedure(symbol));
        }

        return new BoundModule(module, procedures);
    }

    private BoundProcedure BindProcedure(ProcedureSymbol symbol)
    {
        procedure = new ProcedureContext(symbol);
        foreach (var parameter in symbol.Parameters)
        {
            procedure.Locals[parameter.Name] = parameter;
        }

        if (symbol.Result is not null)
        {
            procedure.Locals[symbol.Name] = symbol.Result;
        }

        if (IsForm(module.Tree))
        {
            return new BoundProcedure(symbol, FormBody(symbol), procedure.DeclaredLocals, [], false, false);
        }

        CollectLabels(symbol.Syntax.Body);
        var body = BindBlock(symbol.Syntax.Body, symbol.Syntax);
        return new BoundProcedure(symbol, body, procedure.DeclaredLocals, procedure.Labels.Values.ToList(), procedure.UsesDispatch, procedure.UsesGoSub);
    }

    /// <summary>
    /// What a form's procedure does in the alpha (ROADMAP.md D-D): raise, naming the gap. The form
    /// has no controls and cannot be shown, so a body that read them would read Nothing; the
    /// signatures still compile, which is what lets the rest of such a workbook build and run.
    /// </summary>
    private BoundBlock FormBody(ProcedureSymbol symbol)
    {
        SyntaxNode syntax = symbol.Syntax;
        var raise = new BoundMember(
            syntax,
            new BoundErrObject(syntax),
            "Raise",
            [
                new BoundLiteral(syntax, 438, VbaType.Long),
                new BoundLiteral(syntax, module.Name, VbaType.String),
                new BoundLiteral(syntax, $"UserForms are not supported yet: {module.Name}.{symbol.Name} cannot run, and the form's controls are Nothing.", VbaType.String),
            ],
            VbaType.Variant);
        return new BoundBlock(syntax, [new BoundExpressionStatement(syntax, raise)]);
    }

    private void CollectLabels(SyntaxList<StatementSyntax> statements)
    {
        foreach (var statement in statements)
        {
            switch (statement)
            {
                case LabelStatementSyntax label:
                    {
                        var name = label.IsLineNumber ? label.Label.Text : label.Label.NameValue;
                        if (!procedure.Labels.TryAdd(name, new LabelSymbol(name, label.IsLineNumber, label)))
                        {
                            Report(DiagnosticIds.DuplicateLabel, label.Label, $"Duplicate label: '{name}'.");
                        }

                        break;
                    }

                case IfBlockSyntax ifBlock:
                    CollectLabels(ifBlock.Statements);
                    foreach (var elseIf in ifBlock.ElseIfBlocks)
                    {
                        CollectLabels(elseIf.Statements);
                    }

                    if (ifBlock.ElseBlock is not null)
                    {
                        CollectLabels(ifBlock.ElseBlock.Statements);
                    }

                    break;
                case SelectCaseBlockSyntax select:
                    foreach (var block in select.CaseBlocks)
                    {
                        CollectLabels(block.Statements);
                    }

                    break;
                case ForBlockSyntax forBlock:
                    CollectLabels(forBlock.Statements);
                    break;
                case ForEachBlockSyntax forEach:
                    CollectLabels(forEach.Statements);
                    break;
                case DoLoopBlockSyntax doLoop:
                    CollectLabels(doLoop.Statements);
                    break;
                case WhileBlockSyntax whileBlock:
                    CollectLabels(whileBlock.Statements);
                    break;
                case WithBlockSyntax with:
                    CollectLabels(with.Statements);
                    break;
            }
        }
    }

    private BoundBlock BindBlock(SyntaxList<StatementSyntax> statements, SyntaxNode syntax)
    {
        var bound = new List<BoundStatement>(statements.Count);
        foreach (var statement in statements)
        {
            bound.Add(BindStatement(statement));
        }

        return new BoundBlock(syntax, bound);
    }

    private BoundStatement BindStatement(StatementSyntax statement)
    {
        switch (statement)
        {
            case AttributeStatementSyntax or OptionStatementSyntax or DefTypeStatementSyntax:
                return new BoundNop(statement);
            case VariableDeclarationSyntax declaration:
                return BindLocalDeclaration(declaration);
            case ConstDeclarationSyntax constants:
                BindLocalConstants(constants);
                return new BoundNop(statement);
            case AssignmentStatementSyntax assignment:
                return BindAssignment(assignment);
            case CallStatementSyntax call:
                return BindCallStatement(call);
            case ReDimStatementSyntax redim:
                return BindReDim(redim);
            case EraseStatementSyntax erase:
                return new BoundErase(erase, erase.Targets.Select(BindEraseTarget).ToList());
            case MidStatementSyntax mid:
                return BindMid(mid);
            case IfBlockSyntax ifBlock:
                return BindIf(ifBlock);
            case SingleLineIfStatementSyntax singleIf:
                {
                    var branches = new List<BoundBranch> { new(BindCondition(singleIf.Condition), BindBlock(singleIf.Statements, singleIf)) };
                    var elseBody = singleIf.ElseClause is null ? null : BindBlock(singleIf.ElseClause.Statements, singleIf.ElseClause);
                    return new BoundIf(singleIf, branches, elseBody);
                }

            case SelectCaseBlockSyntax select:
                return BindSelect(select);
            case ForBlockSyntax forBlock:
                return BindFor(forBlock);
            case ForEachBlockSyntax forEach:
                return BindForEach(forEach);
            case DoLoopBlockSyntax doLoop:
                return BindDoLoop(doLoop);
            case WhileBlockSyntax whileBlock:
                {
                    procedure.LoopKinds.Push(ExitKind.Do);
                    var body = BindBlock(whileBlock.Statements, whileBlock);
                    procedure.LoopKinds.Pop();
                    return new BoundDoLoop(whileBlock, BindCondition(whileBlock.Condition), false, body, null, false, whileBlock);
                }

            case WithBlockSyntax with:
                return BindWith(with);
            case GoToStatementSyntax jump:
                {
                    procedure.UsesDispatch = true;
                    var label = LookupLabel(jump.Label);
                    if (jump.Kind == SyntaxKind.GoSubStatement)
                    {
                        procedure.UsesGoSub = true;
                        return new BoundGoSub(jump, label);
                    }

                    return new BoundGoTo(jump, label);
                }

            case KeywordStatementSyntax { Kind: SyntaxKind.ReturnStatement } returnStatement:
                procedure.UsesDispatch = true;
                procedure.UsesGoSub = true;
                return new BoundReturn(returnStatement);
            case KeywordStatementSyntax { Kind: SyntaxKind.EndStatement } end:
                return new BoundEnd(end);
            case KeywordStatementSyntax { Kind: SyntaxKind.StopStatement } stop:
                return new BoundStop(stop);
            case OnGoToStatementSyntax onGoTo:
                {
                    procedure.UsesDispatch = true;
                    var isGoSub = onGoTo.JumpKeyword.Kind == SyntaxKind.GoSubKeyword;
                    procedure.UsesGoSub |= isGoSub;
                    var labels = onGoTo.Labels.Select(LookupLabel).ToList();
                    return new BoundOnGoTo(onGoTo, Convert(BindExpression(onGoTo.Expression), VbaType.Variant), labels, isGoSub);
                }

            case LabelStatementSyntax label:
                procedure.UsesDispatch = true;
                return new BoundLabel(label, procedure.Labels[label.IsLineNumber ? label.Label.Text : label.Label.NameValue]);
            case OnErrorStatementSyntax onError:
                return BindOnError(onError);
            case ResumeStatementSyntax resume:
                procedure.UsesDispatch = true;
                if (resume.NextKeyword is not null)
                {
                    return new BoundResume(resume, ResumeKind.Next, null);
                }

                return resume.Label is null ? new BoundResume(resume, ResumeKind.Retry, null) : new BoundResume(resume, ResumeKind.Label, LookupLabel(resume.Label));
            case ErrorStatementSyntax error:
                return new BoundErrorStatement(error, Convert(BindExpression(error.Number), VbaType.Variant));
            case ExitStatementSyntax exit:
                return BindExit(exit);
            case DebugPrintStatementSyntax print:
                return BindDebugPrint(print);
            case LSetRSetStatementSyntax lset:
                return BindLSet(lset);
            case RaiseEventStatementSyntax raise:
                return BindRaiseEvent(raise);
            case OpenStatementSyntax or CloseStatementSyntax or SeekStatementSyntax or LockStatementSyntax or LineInputStatementSyntax
                or WidthStatementSyntax or FileOutputStatementSyntax or InputStatementSyntax or PutGetStatementSyntax or NameStatementSyntax
                or KeywordStatementSyntax { Kind: SyntaxKind.ResetStatement }:
                return BindFileStatement(statement);
            case BadStatementSyntax:
                return new BoundNop(statement);
            default:
                Report(DiagnosticIds.NotSupported, statement.FirstToken()!, $"{statement.Kind} is not supported yet.");
                return new BoundNop(statement);
        }
    }

    private BoundStatement BindLocalDeclaration(VariableDeclarationSyntax declaration)
    {
        var isStatic = declaration.Modifiers.Any(SyntaxKind.StaticKeyword) || procedure.Symbol.IsStatic;
        var variables = new List<VariableSymbol>();
        foreach (var declarator in declaration.Declarators)
        {
            if (declarator.WithEventsKeyword is not null)
            {
                // MS-VBAL 5.2.3.1.4: WithEvents is a module-level declaration.
                Report(DiagnosticIds.InvalidStatementPlacement, declarator.WithEventsKeyword, "WithEvents is not valid inside a procedure.");
                continue;
            }

            var variable = DeclareVariableType(declarator, isStatic ? VariableKind.Static : VariableKind.Local, false, module, procedure.Symbol);
            if (!procedure.Locals.TryAdd(variable.Name, variable))
            {
                Report(DiagnosticIds.AmbiguousName, declarator.Name, $"Duplicate declaration in current scope: '{variable.Name}'.");
                continue;
            }

            if (isStatic)
            {
                variable.EmitName = "_s_" + procedure.Symbol.EmitName.TrimStart('@') + "_" + variable.Name;
                module.StaticLocals.Add(variable);
            }
            else
            {
                procedure.DeclaredLocals.Add(variable);
            }

            variables.Add(variable);
        }

        return isStatic ? new BoundNop(declaration) : new BoundLocalDeclaration(declaration, variables);
    }

    private void BindLocalConstants(ConstDeclarationSyntax constants)
    {
        foreach (var declarator in constants.Declarators)
        {
            var value = FoldConstant(declarator.Value) ?? Variant.Empty;
            var type = declarator.AsClause is null
                ? (declarator.Name.TypeSuffix != '\0' ? VbaType.FromSuffix(declarator.Name.TypeSuffix)! : VbaType.FromVarType(value.Type))
                : BindType(declarator.AsClause, declarator.Name, isArray: false);
            var symbol = new ConstantSymbol(declarator.Name.NameValue, type, ConvertConstant(value, type, declarator.Name)) { Module = module };
            if (!procedure.Locals.TryAdd(symbol.Name, symbol))
            {
                Report(DiagnosticIds.AmbiguousName, declarator.Name, $"Duplicate declaration in current scope: '{symbol.Name}'.");
            }
        }
    }

    private BoundAssignment BindAssignment(AssignmentStatementSyntax assignment)
    {
        var target = BindTarget(assignment.Target, assignment.IsSet);
        var value = BindExpression(assignment.Value);
        if (assignment.IsSet)
        {
            // MS-VBAL 5.4.3.9 Set statement: both sides are object references or Variants.
            if (!target.Type.IsObject && !target.Type.IsVariant)
            {
                Report(DiagnosticIds.ObjectRequired, assignment.Target.FirstToken()!, "Object required.");
            }

            if (!value.Type.IsObject && !value.Type.IsVariant)
            {
                Report(DiagnosticIds.ObjectRequired, assignment.Value.FirstToken()!, "Object required.");
            }

            return new BoundAssignment(assignment, target, value, isSet: true);
        }

        // A Let into an object variable with a default member assigns through it (MS-VBAL 5.4.3.8, Classes golden).
        if (target is BoundVariable or BoundField && target.Type.ProjectClass is { DefaultMember: not null } defaulted)
        {
            var member = BindClassTarget(target, defaulted, defaulted.DefaultMember!.Name, [], assignment.Target, isSet: false);
            var accepted = member is BoundMember { Target: BoundCall let } ? let.Procedure.Parameters[^1].Type : member.Type;
            return new BoundAssignment(assignment, member, Convert(value, accepted), isSet: false);
        }

        // MS-VBAL 5.4.3.8 Let statement: the value is let-coerced to the target's declared type (5.5.1).
        if (target is BoundMember)
        {
            return new BoundAssignment(assignment, target, Convert(value, VbaType.Variant), isSet: false);
        }

        if (target is BoundComCall comCall)
        {
            // A COM property put let-coerces to the accessor's declared value type; Range("A1") = 5 goes to the default member.
            var accepts = comCall.PutMember is not null || comCall.PutRefMember is not null ? comCall.PutValueType ?? VbaType.Variant : VbaType.Variant;
            return new BoundAssignment(assignment, target, Convert(value, accepts.IsObject ? VbaType.Variant : accepts), isSet: false);
        }

        return new BoundAssignment(assignment, target, Convert(value, target.Type), isSet: false);
    }

    /// <summary>Binds the left side of an assignment: a variable, array element, record field, property, or late-bound member.</summary>
    private BoundExpression BindTarget(ExpressionSyntax syntax, bool isSet)
    {
        var outerTarget = assignmentTarget;
        var outerIsSet = assignmentIsSet;
        assignmentTarget = syntax as MemberAccessExpressionSyntax ?? (syntax as IndexExpressionSyntax)?.Expression as MemberAccessExpressionSyntax;
        assignmentIsSet = isSet;
        try
        {
            return BindTargetCore(syntax, isSet);
        }
        finally
        {
            assignmentTarget = outerTarget;
            assignmentIsSet = outerIsSet;
        }
    }

    private BoundExpression BindTargetCore(ExpressionSyntax syntax, bool isSet)
    {
        switch (syntax)
        {
            case IdentifierNameSyntax name:
                {
                    var symbol = Lookup(name.Name);
                    switch (symbol)
                    {
                        case VariableSymbol variable:
                            return new BoundVariable(name, variable);
                        case null when project.FindModule(name.Name) is { Kind: ModuleKind.Class, Instance: { } instance }:
                            return new BoundVariable(name, instance);
                        case PropertySymbol property:
                            return BindPropertyTarget(property, name, [], isSet);
                        case DocumentMemberSymbol documentMember:
                            return BindComMember(new BoundVariable(name, documentMember.Me), documentMember.Me.Type.ComInterface!, documentMember.Name, [], name, isCallStatement: false);
                        case ComGlobalSymbol global:
                            return BindComGlobal(global, [], name, isCallStatement: false);
                        case null when name.Identifier.Kind == SyntaxKind.ForeignNameToken && BindForeignName(name.Identifier, name, asTarget: true) is { } evaluated:
                            return evaluated;
                        case null:
                            return new BoundVariable(name, ImplicitLocal(name.Identifier));
                        case ProcedureSymbol { ReturnsValue: true } function when ReferenceEquals(function, procedure.Symbol):
                            return new BoundVariable(name, function.Result!);
                        default:
                            Report(DiagnosticIds.ExpectedVariable, name.Identifier, $"'{name.Name}' cannot be assigned to.");
                            return new BoundVariable(name, ImplicitLocal(name.Identifier, force: true));
                    }
                }

            case IndexExpressionSyntax index:
                {
                    if (index.Expression is IdentifierNameSyntax { Name: var callee } && Lookup(callee) is PropertySymbol property)
                    {
                        return BindPropertyTarget(property, index, index.Arguments.Arguments.ToList(), isSet);
                    }

                    if (index.Expression is IdentifierNameSyntax { Name: var receiverName } receiverSyntax
                        && Lookup(receiverName) is VariableSymbol { Type.ProjectClass: { DefaultMember: { } defaultMember } defaultClass } classVariable)
                    {
                        // s(key) = value assigns through the class's default member, early-bound (MS-VBAL 5.6.13.2; Classes golden).
                        return BindClassTarget(new BoundVariable(receiverSyntax, classVariable), defaultClass, defaultMember.Name, index.Arguments.Arguments.ToList(), index, isSet);
                    }

                    var bound = BindIndex(index);
                    if (!bound.IsLValue)
                    {
                        Report(DiagnosticIds.ExpectedVariable, syntax.FirstToken()!, "Expected: variable or array element.");
                    }

                    return bound;
                }

            case MemberAccessExpressionSyntax member:
                {
                    var bound = BindMember(member, null, asTarget: true);
                    if (!bound.IsLValue)
                    {
                        Report(DiagnosticIds.ExpectedVariable, syntax.FirstToken()!, "Expected: variable or property.");
                    }

                    return bound;
                }

            default:
                Report(DiagnosticIds.ExpectedVariable, syntax.FirstToken()!, "Expected: variable.");
                return new BoundLiteral(syntax, Variant.Empty, VbaType.Variant);
        }
    }

    /// <summary>x = value where x is a property: a call to Property Let (or Set) with the value as its last argument.</summary>
    private BoundExpression BindPropertyTarget(PropertySymbol property, SyntaxNode syntax, IReadOnlyList<ArgumentSyntax> indexArguments, bool isSet, BoundExpression? receiver = null)
    {
        var accessor = isSet ? property.Set : property.Let;
        if (accessor is null)
        {
            Report(DiagnosticIds.WrongNumberOfArguments, syntax.FirstToken()!, $"Wrong number of arguments or invalid property assignment: '{property.Name}'.");
            return new BoundLiteral(syntax, Variant.Empty, VbaType.Variant);
        }

        // Bound as a call with one argument short; the emitter appends the assigned value.
        var call = BindProcedureCall(accessor, indexArguments, syntax, asStatement: true, valueParameterPending: true, receiver: receiver);
        return new BoundMember(syntax, call, property.Name, [], accessor.Parameters[^1].Type);
    }

    private BoundStatement BindCallStatement(CallStatementSyntax call)
    {
        var expression = call.Expression;
        var arguments = call.Arguments?.Arguments.ToList() ?? [];
        var byValParenthesized = false;
        if (expression is IndexExpressionSyntax index && call.Arguments is null)
        {
            // Foo(a) without Call: the parenthesized list is the argument list, but a single
            // argument written this way is a parenthesized expression, passed ByVal (MS-VBAL 5.6.13.1).
            arguments = index.Arguments.Arguments.ToList();
            byValParenthesized = call.CallKeyword is null && arguments.Count == 1;
            expression = index.Expression;
        }

        BoundExpression bound;
        switch (expression)
        {
            case IdentifierNameSyntax name:
                bound = BindNameCall(name, arguments, call, byValParenthesized);
                break;
            case MemberAccessExpressionSyntax member:
                bound = BindMember(member, arguments, isCallStatement: true);
                break;
            default:
                Report(DiagnosticIds.ProcedureNotDefined, expression.FirstToken()!, "Expected: Sub or Function.");
                return new BoundNop(call);
        }

        return new BoundExpressionStatement(call, bound);
    }

    private BoundStatement BindReDim(ReDimStatementSyntax redim)
    {
        var statements = new List<BoundStatement>();
        foreach (var declarator in redim.Declarators)
        {
            BoundExpression target;
            if (declarator.Target is IdentifierNameSyntax name && Lookup(name.Name) is null)
            {
                // ReDim declares the array when the name is new (MS-VBAL 5.4.3.3).
                var elementType = declarator.AsClause is null ? DefaultType(name.Identifier) : BindType(declarator.AsClause, name.Identifier, isArray: false);
                var variable = ImplicitLocal(name.Identifier, force: true, type: VbaType.ArrayOf(elementType));
                target = new BoundVariable(name, variable);
            }
            else
            {
                target = BindTarget(declarator.Target, isSet: false);
            }

            var declaredElement = target.Type.IsArray ? target.Type.ElementType! : declarator.AsClause is null ? VbaType.Variant : BindType(declarator.AsClause, declarator.Target.FirstToken()!, isArray: false);
            if (!target.Type.IsArray && !target.Type.IsVariant)
            {
                Report(DiagnosticIds.ExpectedArray, declarator.Target.FirstToken()!, "Expected array.");
            }

            var bounds = new List<BoundBound>();
            foreach (var bound in declarator.Bounds.Bounds)
            {
                var lower = bound.Lower is null ? null : Convert(BindExpression(bound.Lower), VbaType.Long);
                bounds.Add(new BoundBound(lower, Convert(BindExpression(bound.Upper), VbaType.Long)));
            }

            statements.Add(new BoundReDim(declarator, target, bounds, redim.PreserveKeyword is not null, declaredElement));
        }

        return statements.Count == 1 ? statements[0] : new BoundBlock(redim, statements);
    }

    private BoundExpression BindEraseTarget(ExpressionSyntax syntax)
    {
        var target = BindTarget(syntax, isSet: false);
        if (!target.Type.IsArray && !target.Type.IsVariant)
        {
            Report(DiagnosticIds.ExpectedArray, syntax.FirstToken()!, "Expected array.");
        }

        return target;
    }

    private BoundMidAssignment BindMid(MidStatementSyntax mid)
    {
        if (mid.MidKeyword.Text.StartsWith("MidB", StringComparison.OrdinalIgnoreCase))
        {
            Report(DiagnosticIds.NotSupported, mid.MidKeyword, "MidB needs strings with an odd byte count.");
        }

        var target = BindTarget(mid.Target, isSet: false);
        var start = Convert(BindExpression(mid.StartIndex), VbaType.Long);
        var length = mid.Length is null ? null : Convert(BindExpression(mid.Length), VbaType.Long);
        var value = Convert(BindExpression(mid.Value), VbaType.String);
        return new BoundMidAssignment(mid, target, start, length, value);
    }

    private BoundIf BindIf(IfBlockSyntax ifBlock)
    {
        var branches = new List<BoundBranch> { new(BindCondition(ifBlock.Condition), BindBlock(ifBlock.Statements, ifBlock)) };
        foreach (var elseIf in ifBlock.ElseIfBlocks)
        {
            branches.Add(new BoundBranch(BindCondition(elseIf.Condition), BindBlock(elseIf.Statements, elseIf)));
        }

        var elseBody = ifBlock.ElseBlock is null ? null : BindBlock(ifBlock.ElseBlock.Statements, ifBlock.ElseBlock);
        return new BoundIf(ifBlock, branches, elseBody);
    }

    /// <summary>A condition is any expression; Null counts as False (MS-VBAL 5.4.2.8: Null is treated as False).</summary>
    private BoundExpression BindCondition(ExpressionSyntax syntax) => Convert(BindExpression(syntax), VbaType.Variant);

    /// <summary>LSet and RSet (MS-VBAL 5.4.3.6, 5.4.3.7): a String or fixed-length String target with a value coerced to String, or LSet between two records; anything else is a type mismatch.</summary>
    private BoundStatement BindLSet(LSetRSetStatementSyntax lset)
    {
        var isRSet = lset.Kind == SyntaxKind.RSetStatement;
        var target = BindTarget(lset.Target, isSet: false);
        var value = BindExpression(lset.Value);
        if (target.Type.IsRecord)
        {
            if (isRSet || !value.Type.IsRecord)
            {
                Report(DiagnosticIds.TypeMismatch, lset.Keyword, isRSet ? "RSet is not allowed between user-defined types." : "LSet between a user-defined type and a value of another type.");
                return new BoundNop(lset);
            }

            return new BoundLSet(lset, target, value, isRSet);
        }

        if (!target.Type.IsString && target.Type.Kind != TypeKind.FixedString && !target.Type.IsVariant)
        {
            Report(DiagnosticIds.TypeMismatch, lset.Keyword, "LSet and RSet need a String or a user-defined type as the target.");
            return new BoundNop(lset);
        }

        return new BoundLSet(lset, target, Convert(value, VbaType.String), isRSet);
    }

    private BoundSelect BindSelect(SelectCaseBlockSyntax select)
    {
        var value = Convert(BindExpression(select.Expression), VbaType.Variant);
        var selector = NewTemporary("select", VbaType.Variant);
        var cases = new List<BoundCaseBlock>();
        BoundBlock? elseBody = null;
        foreach (var block in select.CaseBlocks)
        {
            if (block.IsCaseElse)
            {
                elseBody = BindBlock(block.Statements, block);
                continue;
            }

            var clauses = new List<BoundCaseClause>();
            foreach (var clause in block.Clauses)
            {
                switch (clause)
                {
                    case ValueCaseClauseSyntax single:
                        clauses.Add(new BoundCaseClause(CaseClauseKind.Value, BinaryKind.Equal, Convert(BindExpression(single.Value), VbaType.Variant), null));
                        break;
                    case RangeCaseClauseSyntax range:
                        clauses.Add(new BoundCaseClause(CaseClauseKind.Range, BinaryKind.Equal, Convert(BindExpression(range.From), VbaType.Variant), Convert(BindExpression(range.To), VbaType.Variant)));
                        break;
                    case IsCaseClauseSyntax comparison:
                        clauses.Add(new BoundCaseClause(CaseClauseKind.Comparison, ComparisonKind(comparison.OperatorToken.Kind), Convert(BindExpression(comparison.Value), VbaType.Variant), null));
                        break;
                }
            }

            cases.Add(new BoundCaseBlock(clauses, BindBlock(block.Statements, block)));
        }

        return new BoundSelect(select, selector, value, cases, elseBody);
    }

    private static BinaryKind ComparisonKind(SyntaxKind kind) => kind switch
    {
        SyntaxKind.LessThanGreaterThanToken => BinaryKind.NotEqual,
        SyntaxKind.LessThanToken => BinaryKind.LessThan,
        SyntaxKind.GreaterThanToken => BinaryKind.GreaterThan,
        SyntaxKind.LessThanEqualsToken => BinaryKind.LessThanOrEqual,
        SyntaxKind.GreaterThanEqualsToken => BinaryKind.GreaterThanOrEqual,
        _ => BinaryKind.Equal,
    };

    private BoundFor BindFor(ForBlockSyntax forBlock)
    {
        var counter = BindTarget(forBlock.Variable, isSet: false);
        if (!counter.Type.IsNumeric && !counter.Type.IsVariant)
        {
            Report(DiagnosticIds.TypeMismatch, forBlock.Variable.FirstToken()!, "For counter variable must be numeric.");
        }

        // The start, limit, and step are Let-coerced to a typed counter's type before the loop (MS-VBAL 5.4.2.3): For i = 1 To "3" runs three times and To Null raises 94 (ControlFlow golden); a Variant counter keeps them as they are.
        var boundType = counter.Type.IsVariant ? VbaType.Variant : counter.Type;
        var from = Convert(BindExpression(forBlock.From), counter.Type);
        var to = Convert(BindExpression(forBlock.To), boundType);
        var step = forBlock.StepClause is null ? null : Convert(BindExpression(forBlock.StepClause.Value), boundType);
        // A counter of a type the runtime steps natively keeps its limit and step in slots of that type (ROADMAP.md D-J, M7 E).
        var slotType = counter is BoundVariable && IsNativeCounter(counter.Type) ? counter.Type : VbaType.Variant;
        var limit = NewTemporary("limit", slotType);
        var increment = NewTemporary("step", slotType);
        procedure.LoopKinds.Push(ExitKind.For);
        var body = BindBlock(forBlock.Statements, forBlock);
        procedure.LoopKinds.Pop();
        return new BoundFor(forBlock, counter, from, to, step, limit, increment, body, (SyntaxNode?)forBlock.Next ?? forBlock);
    }

    private static bool IsNativeCounter(VbaType type) =>
        type.Kind == TypeKind.Builtin && type.VarType is VarType.Byte or VarType.Integer or VarType.Long or VarType.LongLong or VarType.Double;

    private BoundForEach BindForEach(ForEachBlockSyntax forEach)
    {
        var variable = BindTarget(forEach.Variable, isSet: false);
        if (!variable.Type.IsVariant && !variable.Type.IsObject)
        {
            Report(DiagnosticIds.TypeMismatch, forEach.Variable.FirstToken()!, "For Each control variable must be Variant or Object.");
        }

        var collection = Convert(BindExpression(forEach.Collection), VbaType.Variant);
        var enumerator = NewTemporary("each", VbaType.Variant);
        enumerator.EmitType = "global::VbaNg.Runtime.ForEachEnumerator";
        enumerator.EmitDefault = "null!";
        procedure.LoopKinds.Push(ExitKind.For);
        var body = BindBlock(forEach.Statements, forEach);
        procedure.LoopKinds.Pop();
        return new BoundForEach(forEach, variable, collection, enumerator, body, (SyntaxNode?)forEach.Next ?? forEach);
    }

    private BoundDoLoop BindDoLoop(DoLoopBlockSyntax doLoop)
    {
        var top = doLoop.TopCondition is null ? null : BindCondition(doLoop.TopCondition.Condition);
        procedure.LoopKinds.Push(ExitKind.Do);
        var body = BindBlock(doLoop.Statements, doLoop);
        procedure.LoopKinds.Pop();
        var bottom = doLoop.BottomCondition is null ? null : BindCondition(doLoop.BottomCondition.Condition);
        return new BoundDoLoop(doLoop, top, doLoop.TopCondition?.IsUntil ?? false, body, bottom, doLoop.BottomCondition?.IsUntil ?? false, doLoop);
    }

    private BoundWith BindWith(WithBlockSyntax with)
    {
        var target = BindExpression(with.Expression);
        var type = target.Type.IsRecord || target.Type.IsObject ? target.Type : VbaType.Variant;
        if (!target.Type.IsRecord && !target.Type.IsObject && !target.Type.IsVariant)
        {
            Report(DiagnosticIds.ObjectRequired, with.Expression.FirstToken()!, "Object required.");
        }

        var temporary = NewTemporary("with", type);

        // A With over a record with storage of its own works on that storage; a record value (a call's result) is copied into the temporary instead.
        temporary.IsAlias = type.IsRecord && target.IsLValue;
        procedure.WithTargets.Push(temporary);
        var body = BindBlock(with.Statements, with);
        procedure.WithTargets.Pop();
        return new BoundWith(with, temporary, Convert(target, type), body);
    }

    private BoundOnError BindOnError(OnErrorStatementSyntax onError)
    {
        procedure.UsesDispatch = true;
        if (onError.IsResumeNext)
        {
            return new BoundOnError(onError, OnErrorMode.ResumeNext, null);
        }

        var target = onError.Target!;
        if (target is LiteralExpressionSyntax { Token: { Kind: SyntaxKind.IntegerLiteralToken, Text: "0" } })
        {
            return new BoundOnError(onError, OnErrorMode.Disable, null);
        }

        if (target is UnaryExpressionSyntax { OperatorToken.Kind: SyntaxKind.MinusToken, Operand: LiteralExpressionSyntax { Token.Text: "1" } })
        {
            return new BoundOnError(onError, OnErrorMode.Dismiss, null);
        }

        return new BoundOnError(onError, OnErrorMode.GoToLabel, LookupLabel(target));
    }

    private BoundExit BindExit(ExitStatementSyntax exit)
    {
        switch (exit.BlockKeyword.Kind)
        {
            case SyntaxKind.ForKeyword:
                if (!procedure.LoopKinds.Contains(ExitKind.For))
                {
                    Report(DiagnosticIds.InvalidStatementPlacement, exit.ExitKeyword, "Exit For not within For...Next.");
                }

                return new BoundExit(exit, ExitKind.For);
            case SyntaxKind.DoKeyword:
                if (!procedure.LoopKinds.Contains(ExitKind.Do))
                {
                    Report(DiagnosticIds.InvalidStatementPlacement, exit.ExitKeyword, "Exit Do not within Do...Loop.");
                }

                return new BoundExit(exit, ExitKind.Do);
            default:
                // VBA does not check which Exit leaves a procedure either: Exit Sub, Exit Function,
                // and Exit Property are interchangeable there, which the VBE accepts in all five
                // pairings and stdVBA's sources rely on (docs/vba-quirks.md; R3). Each one leaves
                // the procedure it stands in, which is the single operation VBA has.
                return new BoundExit(exit, ExitKind.Procedure);
        }
    }

    private BoundDebugPrint BindDebugPrint(DebugPrintStatementSyntax print) => new(print, BindOutputItems(print.Items));

    private LabelSymbol LookupLabel(SyntaxToken token) => LookupLabel(token.Kind == SyntaxKind.IntegerLiteralToken ? token.Text : token.NameValue, token);

    private LabelSymbol LookupLabel(ExpressionSyntax expression) => expression switch
    {
        IdentifierNameSyntax name => LookupLabel(name.Name, name.Identifier),
        LiteralExpressionSyntax literal => LookupLabel(literal.Token.Text, literal.Token),
        _ => LookupLabel(expression.ToFullString().Trim(), expression.FirstToken()!),
    };

    private LabelSymbol LookupLabel(string name, SyntaxToken at)
    {
        if (procedure.Labels.TryGetValue(name, out var label))
        {
            return label;
        }

        Report(DiagnosticIds.LabelNotDefined, at, $"Label not defined: '{name}'.");
        label = new LabelSymbol(name, false, new LabelStatementSyntax(at, null));
        procedure.Labels[name] = label;
        return label;
    }

    private VariableSymbol NewTemporary(string purpose, VbaType type)
    {
        var temporary = new VariableSymbol(purpose, type, VariableKind.Temporary)
        {
            Module = module,
            Procedure = procedure.Symbol,
            EmitName = "__" + purpose + (++procedure.TemporaryCount).ToString(CultureInfo.InvariantCulture),
        };
        procedure.DeclaredLocals.Add(temporary);
        return temporary;
    }

    /// <summary>An undeclared name: an error under Option Explicit (MS-VBAL 5.2.1.2), otherwise a new local of the DefType or Variant.</summary>
    private VariableSymbol ImplicitLocal(SyntaxToken nameToken, bool force = false, VbaType? type = null)
    {
        var name = nameToken.NameValue;
        if (procedure.Locals.TryGetValue(name, out var existing) && existing is VariableSymbol variable)
        {
            return variable;
        }

        if (module.Options.Explicit && !force)
        {
            Report(DiagnosticIds.VariableNotDefined, nameToken, $"Variable not defined: '{name}'.");
        }

        variable = new VariableSymbol(name, type ?? DefaultType(nameToken), procedure.Symbol.IsStatic ? VariableKind.Static : VariableKind.Local)
        {
            Module = module,
            Procedure = procedure.Symbol,
            EmitName = CSharpNames.Identifier(name),
        };
        procedure.Locals[name] = variable;
        if (variable.Kind == VariableKind.Static)
        {
            variable.EmitName = "_s_" + procedure.Symbol.EmitName.TrimStart('@') + "_" + name;
            module.StaticLocals.Add(variable);
        }
        else
        {
            procedure.DeclaredLocals.Add(variable);
        }

        return variable;
    }

    private void Report(string id, SyntaxToken token, string message)
    {
        var (line, column) = module.Tree.GetLinePosition(token);
        diagnostics.Add(new Diagnostic(id, DiagnosticSeverity.Error, message, module.FilePath, line, column));
    }

    private void Warn(string id, SyntaxToken token, string message)
    {
        var (line, column) = module.Tree.GetLinePosition(token);
        diagnostics.Add(new Diagnostic(id, DiagnosticSeverity.Warning, message, module.FilePath, line, column));
    }

    /// <summary>Per-procedure binding state: locals, labels, the loop and With nesting, and what the emitter must know.</summary>
    private sealed class ProcedureContext(ProcedureSymbol symbol)
    {
        public ProcedureSymbol Symbol { get; } = symbol;

        public Dictionary<string, Symbol> Locals { get; } = new(StringComparer.OrdinalIgnoreCase);

        public List<VariableSymbol> DeclaredLocals { get; } = [];

        public Dictionary<string, LabelSymbol> Labels { get; } = new(StringComparer.OrdinalIgnoreCase);

        public Stack<ExitKind> LoopKinds { get; } = new();

        public Stack<VariableSymbol> WithTargets { get; } = new();

        public int TemporaryCount { get; set; }

        public bool UsesDispatch { get; set; }

        public bool UsesGoSub { get; set; }
    }
}
