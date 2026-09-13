using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>Anything a name can resolve to (MS-VBAL 5.6.10 simple name expressions).</summary>
public abstract class Symbol
{
    protected Symbol(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Name = name;
    }

    /// <summary>The name as first declared; lookups are case-insensitive.</summary>
    public string Name { get; }

    public override string ToString() => Name;
}

/// <summary>A project: its modules, in file order.</summary>
public sealed class ProjectSymbol(string name) : Symbol(name)
{
    public List<ModuleSymbol> Modules { get; } = [];

    public ModuleSymbol? FindModule(string name) =>
        Modules.FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>The Option statements of a module (MS-VBAL 5.2.1).</summary>
public sealed class ModuleOptions
{
    public bool Explicit { get; set; }

    public int Base { get; set; }

    public CompareMode Compare { get; set; }

    public bool PrivateModule { get; set; }
}

public enum ModuleKind
{
    /// <summary>A .bas file.</summary>
    Standard,

    /// <summary>A class module with a predeclared instance bound to a workbook object (ARCHITECTURE.md section 3).</summary>
    Document,

    /// <summary>A class module (MS-VBAL 4.2): instances are created with New, or reached through a predeclared instance.</summary>
    Class,
}

/// <summary>A standard or document module and its members (MS-VBAL 4.2, 5.1).</summary>
public sealed class ModuleSymbol(string name, SyntaxTree tree) : Symbol(name)
{
    public SyntaxTree Tree { get; } = tree;

    public ModuleKind Kind { get; set; }

    /// <summary>Workbook, Worksheet, or Chart for a document module.</summary>
    public string? DocumentKind { get; set; }

    /// <summary>The bound object of a document module, as <c>Me</c> names it; the host sets it at load. In a class module, the instance itself.</summary>
    public VariableSymbol? Me { get; set; }

    /// <summary>The type of a class module's instances (MS-VBAL 5.2.4.1.3).</summary>
    public VbaType? ClassType { get; set; }

    /// <summary>The predeclared instance of a class module with VB_PredeclaredId = True, reached by the class name; created on first use, like As New.</summary>
    public VariableSymbol? Instance { get; set; }

    /// <summary>The member marked <c>VB_UserMemId = 0</c>: what the object stands for where a value is needed (MS-VBAL 5.6.9.3).</summary>
    public Symbol? DefaultMember { get; set; }

    /// <summary>The member marked <c>VB_UserMemId = -4</c>: the enumerator For Each walks.</summary>
    public ProcedureSymbol? EnumMember { get; set; }

    /// <summary>Another class module names this one in an Implements statement, so this class is emitted as a C# interface plus its own implementation.</summary>
    public bool IsImplemented { get; set; }

    /// <summary>The class modules this one implements (MS-VBAL 5.2.4.2), in declaration order.</summary>
    public List<ModuleSymbol> Implemented { get; } = [];

    /// <summary>The events this class declares (MS-VBAL 5.2.4.1.2).</summary>
    public List<EventSymbol> Events { get; } = [];

    /// <summary>The WithEvents variables of this module: the source objects whose events its procedures handle (MS-VBAL 5.2.3.1.4).</summary>
    public List<VariableSymbol> EventSources { get; } = [];

    /// <summary>The C# type name of a class module's own implementation; the interface takes the plain name when another class implements it.</summary>
    public string ClassEmitName => IsImplemented ? EmitName + "__Impl" : EmitName;

    public string FilePath => Tree.FilePath;

    /// <summary>The C# class name; the module name, escaped when it is a C# keyword.</summary>
    public string EmitName { get; set; } = name;

    public ModuleOptions Options { get; } = new();

    /// <summary>DefType letters (MS-VBAL 5.2.2): the declared type of undeclared names by first letter.</summary>
    public Dictionary<char, VbaType> DefTypes { get; } = [];

    /// <summary>Every module-level name: variables, constants, procedures, properties, records, enums, and enum members.</summary>
    public Dictionary<string, Symbol> Members { get; } = new(StringComparer.OrdinalIgnoreCase);

    public List<VariableSymbol> Variables { get; } = [];

    public List<ConstantSymbol> Constants { get; } = [];

    public List<ProcedureSymbol> Procedures { get; } = [];

    public List<RecordSymbol> Records { get; } = [];

    public List<EnumSymbol> Enums { get; } = [];

    /// <summary>Static locals of the module's procedures, emitted as fields of the module class.</summary>
    public List<VariableSymbol> StaticLocals { get; } = [];

    /// <summary>The Declare statements of the module (MS-VBAL 5.2.3.5): procedures that live in a DLL and have no body here.</summary>
    public List<ProcedureSymbol> Externals { get; } = [];
}

public enum VariableKind
{
    Local,
    Parameter,
    Module,

    /// <summary>A Static local (MS-VBAL 5.4.3.1): lives in the module, scoped to its procedure.</summary>
    Static,

    /// <summary>The implicit result variable of a Function or Property Get (MS-VBAL 5.3.1.8).</summary>
    Result,

    /// <summary>A field of a user-defined type.</summary>
    Field,

    /// <summary>A compiler temporary: With targets, loop limits, copy-back arguments.</summary>
    Temporary,
}

/// <summary>A variable: local, module-level, static, parameter, procedure result, record field, or temporary.</summary>
public class VariableSymbol(string name, VbaType type, VariableKind kind) : Symbol(name)
{
    public VbaType Type { get; } = type;

    public VariableKind Kind { get; } = kind;

    public bool IsPublic { get; set; }

    public ModuleSymbol? Module { get; set; }

    public ProcedureSymbol? Procedure { get; set; }

    /// <summary>The constant bounds of a fixed-size array (Dim a(1 To 3)); null for scalars and dynamic arrays.</summary>
    public (int Lower, int Upper)[]? Bounds { get; set; }

    /// <summary>
    /// A fixed-size array member of a user-defined type whose elements lie inline in the record,
    /// as VBA lays them out (ROADMAP.md M7 C6): elements of a type an array's storage holds as
    /// itself, records included (C9). An array of fixed-length strings lies inline too, as its
    /// characters, which the emitter reaches by index rather than through a descriptor (C10).
    /// </summary>
    public bool IsEmbeddedArray => Kind == VariableKind.Field && Bounds is not null && Type.ElementType is { } element
        && (element.IsRecord || element.IsEnum || element.IsObject || element.IsVariant || element.IsVariableString
            || (element.Kind == TypeKind.Builtin && element.VarType is Runtime.VarType.Byte or Runtime.VarType.Integer or Runtime.VarType.Long or Runtime.VarType.LongLong
                or Runtime.VarType.Single or Runtime.VarType.Double or Runtime.VarType.Currency or Runtime.VarType.Date or Runtime.VarType.Boolean));

    /// <summary>Declared As New: the object is created on first use (MS-VBAL 5.2.3.1.1 As New).</summary>
    public bool IsNew { get; set; }

    /// <summary>Declared WithEvents: assigning the variable subscribes this module's handlers to the source's events (MS-VBAL 5.2.3.1.4).</summary>
    public bool IsWithEvents { get; set; }

    /// <summary>For a WithEvents variable of a library type: the default source interface of its coclass, whose events the handlers take (ARCHITECTURE.md section 6, "Events").</summary>
    public Runtime.TypeLibraries.ComType? EventInterface { get; set; }

    public bool IsArray => Type.IsArray;

    /// <summary>The C# identifier; distinct from every other name in scope.</summary>
    public string EmitName { get; set; } = name;

    /// <summary>For compiler temporaries whose C# type is not a VBA type (a For Each enumerator): the C# type and initializer to declare them with.</summary>
    public string? EmitType { get; set; }

    public string? EmitDefault { get; set; }

    /// <summary>A With temporary standing for a record variable, field, or element rather than a copy of it (MS-VBAL 5.4.2.7): a ref to that storage, owning nothing (ROADMAP.md M7 C3).</summary>
    public bool IsAlias { get; set; }

    public SyntaxNode? Syntax { get; set; }
}

/// <summary>A procedure parameter (MS-VBAL 5.3.1.5).</summary>
public sealed class ParameterSymbol(string name, VbaType type) : VariableSymbol(name, type, VariableKind.Parameter)
{
    public bool IsByVal { get; set; }

    public bool IsOptional { get; set; }

    public bool IsParamArray { get; set; }

    /// <summary>The parameter belongs to a Declare, so its argument travels to a DLL (MS-VBAL 5.2.3.5).</summary>
    public bool IsExternal { get; set; }

    /// <summary>The generated parameter takes a reference: every ByRef parameter, and the string buffer of a Declare that the callee may write into.</summary>
    public bool EmitByRef { get; set; }

    /// <summary>The default of an Optional parameter, a constant; null means the type's default, or Missing for a Variant.</summary>
    public Variant? Default { get; set; }

    public int Index { get; set; }
}

/// <summary>A Const (MS-VBAL 5.2.3.2, 5.4.3.2) or an enum member: a name for a compile-time value.</summary>
public sealed class ConstantSymbol(string name, VbaType type, Variant value) : Symbol(name)
{
    public VbaType Type { get; } = type;

    public Variant Value { get; } = value;

    public bool IsPublic { get; set; }

    public ModuleSymbol? Module { get; set; }
}

public enum ProcedureKind
{
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

/// <summary>A Sub, Function, or Property procedure of a standard module (MS-VBAL 5.3.1).</summary>
public sealed class ProcedureSymbol(string name, ProcedureKind kind, ModuleSymbol module, ProcedureDeclarationSyntax syntax) : Symbol(name)
{
    public ProcedureKind Kind { get; } = kind;

    public ModuleSymbol Module { get; } = module;

    public ProcedureDeclarationSyntax Syntax { get; } = syntax;

    public List<ParameterSymbol> Parameters { get; } = [];

    /// <summary>The declared result type of a Function or Property Get; null for Subs and Let/Set.</summary>
    public VbaType? ReturnType { get; set; }

    public bool IsPublic { get; set; }

    /// <summary>Declared Static: every local keeps its value between calls (MS-VBAL 5.3.1.2).</summary>
    public bool IsStatic { get; set; }

    /// <summary>A test procedure: a Public Sub without parameters under a <c>'@Test</c> comment annotation (ARCHITECTURE.md section 8).</summary>
    public bool IsTest { get; set; }

    /// <summary>For an event procedure in a document module (Worksheet_Change, CommandButton1_Click): the source object's name.</summary>
    public string? EventSource { get; set; }

    /// <summary>For an event procedure: the event's name on the source interface.</summary>
    public string? EventName { get; set; }

    /// <summary>For the handler of an event of a WithEvents variable (<c>source_Changed</c>): the variable; null for the handlers a document module's own object and its controls raise.</summary>
    public VariableSymbol? EventVariable { get; set; }

    /// <summary>The dispid a <c>VB_UserMemId</c> attribute gives the member: 0 for the default member, -4 for the enumerator.</summary>
    public int? UserMemberId { get; set; }

    /// <summary>For a procedure that implements an interface member (<c>IShape_Area</c>): the interface and its member.</summary>
    public ModuleSymbol? ImplementsInterface { get; set; }

    public string? ImplementsMember { get; set; }

    /// <summary>The implicit result variable of a Function or Property Get.</summary>
    public VariableSymbol? Result { get; set; }

    public bool ReturnsValue => Kind is ProcedureKind.Function or ProcedureKind.PropertyGet;

    /// <summary>The C# method name: the VBA name, prefixed for property accessors.</summary>
    public string EmitName { get; set; } = name;

    /// <summary>For a Declare (MS-VBAL 5.2.3.5): the library the procedure lives in.</summary>
    public string? ExternalLibrary { get; set; }

    /// <summary>For a Declare: the entry point, which is the Alias when there is one and the declared name otherwise.</summary>
    public string? ExternalEntryPoint { get; set; }

    public bool IsExternal => ExternalLibrary is not null;

    /// <summary>A Declare of VBE7's VarPtr under another name (VarPtrArray, VarPtrStringArray): the runtime answers it with the address of what it is handed, as VBE7 does (ARCHITECTURE.md section 5, "Arrays").</summary>
    public bool IsRuntimeVarPtr =>
        ExternalLibrary is { } library
        && (library.Equals("VBE7", StringComparison.OrdinalIgnoreCase) || library.Equals("VBE7.DLL", StringComparison.OrdinalIgnoreCase))
        && string.Equals(ExternalEntryPoint, "VarPtr", StringComparison.OrdinalIgnoreCase);

    /// <summary>A Declare of VBE7's (or msvbvm60's) rtcCallByName, which stdVBA calls for speed: the runtime answers it as CallByName, so it works where VBE7 is not loaded too (Declares golden).</summary>
    public bool IsRuntimeCallByName =>
        ExternalLibrary is { } library
        && (library.Equals("VBE7", StringComparison.OrdinalIgnoreCase) || library.Equals("VBE7.DLL", StringComparison.OrdinalIgnoreCase)
            || library.Equals("msvbvm60", StringComparison.OrdinalIgnoreCase) || library.Equals("msvbvm60.dll", StringComparison.OrdinalIgnoreCase))
        && string.Equals(ExternalEntryPoint, "rtcCallByName", StringComparison.OrdinalIgnoreCase);

    public int RequiredArguments => Parameters.Count(p => !p.IsOptional && !p.IsParamArray);

    public ParameterSymbol? ParamArray => Parameters.LastOrDefault(p => p.IsParamArray);
}

/// <summary>The Get, Let, and Set procedures that share a property name (MS-VBAL 5.3.1.4).</summary>
public sealed class PropertySymbol(string name, ModuleSymbol module) : Symbol(name)
{
    public ModuleSymbol Module { get; } = module;

    public ProcedureSymbol? Get { get; set; }

    public ProcedureSymbol? Let { get; set; }

    public ProcedureSymbol? Set { get; set; }

    public bool IsPublic => (Get?.IsPublic ?? false) || (Let?.IsPublic ?? false) || (Set?.IsPublic ?? false);
}

/// <summary>A user-defined type (MS-VBAL 5.2.3.3), emitted as a class with copy semantics.</summary>
public sealed class RecordSymbol : Symbol
{
    public RecordSymbol(string name, ModuleSymbol module, TypeDefinitionSyntax syntax, string? emitName = null)
        : base(name)
    {
        Module = module;
        Syntax = syntax;
        EmitName = emitName ?? name;
        Type = VbaType.ForRecord(this, "global::" + module.EmitName + "." + EmitName);
    }

    public ModuleSymbol Module { get; }

    public TypeDefinitionSyntax Syntax { get; }

    public List<VariableSymbol> Fields { get; } = [];

    public bool IsPublic { get; set; }

    public VbaType Type { get; }

    public string EmitName { get; }

    public VariableSymbol? FindField(string name) => Fields.FirstOrDefault(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>An Enum (MS-VBAL 5.2.3.4): named Long constants, also visible unqualified.</summary>
public sealed class EnumSymbol : Symbol
{
    public EnumSymbol(string name, ModuleSymbol module, EnumDefinitionSyntax syntax)
        : base(name)
    {
        Module = module;
        Syntax = syntax;
        EmitName = name;
        Type = VbaType.ForEnum(this);
    }

    public ModuleSymbol Module { get; }

    public EnumDefinitionSyntax Syntax { get; }

    public List<ConstantSymbol> Members { get; } = [];

    public bool IsPublic { get; set; }

    public VbaType Type { get; }

    /// <summary>The nested C# class name; the enum name unless that is the module's own name.</summary>
    public string EmitName { get; set; }
}

/// <summary>An event a class module declares (MS-VBAL 5.2.4.1.2): a name and the parameters RaiseEvent passes.</summary>
public sealed class EventSymbol(string name, ModuleSymbol module, EventDeclarationSyntax syntax) : Symbol(name)
{
    public ModuleSymbol Module { get; } = module;

    public EventDeclarationSyntax Syntax { get; } = syntax;

    public List<ParameterSymbol> Parameters { get; } = [];
}

/// <summary>A statement label or line number (MS-VBAL 5.4.1.1).</summary>
public sealed class LabelSymbol(string name, bool isLineNumber, LabelStatementSyntax syntax) : Symbol(name)
{
    public bool IsLineNumber { get; } = isLineNumber;

    public LabelStatementSyntax Syntax { get; } = syntax;

    /// <summary>The line number value; 0 for named labels.</summary>
    public int Number => IsLineNumber ? int.Parse(Name, System.Globalization.CultureInfo.InvariantCulture) : 0;

    /// <summary>Set by lowering: the index of the flat statement the label marks.</summary>
    public int Target { get; set; } = -1;

    /// <summary>The C# label used in structured procedures; assigned by the emitter.</summary>
    public string EmitName { get; set; } = name;
}

/// <summary>
/// A function or statement of the VBA standard library (MS-VBAL 6.1), described by the C# it
/// compiles to. <see cref="Templates"/> holds one template per argument count starting at
/// <see cref="MinArguments"/>; the last one serves every larger count. See <see cref="StandardLibrary"/>
/// for the placeholder syntax.
/// </summary>
public sealed class IntrinsicSymbol(string name, VbaType returnType, VbaType resultType, int minArguments, int maxArguments, IReadOnlyList<string> templates) : Symbol(name)
{
    /// <summary>The declared type of the call expression in VBA terms.</summary>
    public VbaType ReturnType { get; } = returnType;

    /// <summary>The C# type the template evaluates to; the emitter converts from this to the declared type.</summary>
    public VbaType ResultType { get; } = resultType;

    public int MinArguments { get; } = minArguments;

    public int MaxArguments { get; } = maxArguments;

    public IReadOnlyList<string> Templates { get; } = templates;

    /// <summary>A statement such as Randomize: usable only as a call statement.</summary>
    public bool IsSub { get; init; }

    /// <summary>The milestone that brings this function, for the VBA0002 message; null when it is available.</summary>
    public string? Pending { get; init; }

    public string TemplateFor(int argumentCount)
    {
        var index = Math.Clamp(argumentCount - MinArguments, 0, Templates.Count - 1);
        return Templates[index];
    }
}
