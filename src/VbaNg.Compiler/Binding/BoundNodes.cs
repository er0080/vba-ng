using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>A node of the bound tree: the syntax it came from, with names resolved and every expression typed.</summary>
public abstract class BoundNode(SyntaxNode syntax)
{
    public SyntaxNode Syntax { get; } = syntax;
}

public abstract class BoundExpression(SyntaxNode syntax, VbaType type) : BoundNode(syntax)
{
    /// <summary>The declared type of the expression (MS-VBAL 5.6 "declared type").</summary>
    public VbaType Type { get; } = type;

    /// <summary>True for variables, array elements, and record fields: the expression names storage.</summary>
    public virtual bool IsLValue => false;
}

public sealed class BoundLiteral(SyntaxNode syntax, Variant value, VbaType type) : BoundExpression(syntax, type)
{
    public Variant Value { get; } = value;
}

public sealed class BoundVariable(SyntaxNode syntax, VariableSymbol variable) : BoundExpression(syntax, variable.Type)
{
    public VariableSymbol Variable { get; } = variable;

    public override bool IsLValue => true;
}

public sealed class BoundConstant(SyntaxNode syntax, ConstantSymbol constant) : BoundExpression(syntax, constant.Type)
{
    public ConstantSymbol Constant { get; } = constant;
}

/// <summary>An array element, or an index applied to a Variant or Object bound at run time (MS-VBAL 5.6.13).</summary>
public sealed class BoundElement(SyntaxNode syntax, BoundExpression array, IReadOnlyList<BoundExpression> indices, VbaType type) : BoundExpression(syntax, type)
{
    public BoundExpression Array { get; } = array;

    public IReadOnlyList<BoundExpression> Indices { get; } = indices;

    public override bool IsLValue => true;
}

/// <summary>A field of a user-defined type value.</summary>
public sealed class BoundField(SyntaxNode syntax, BoundExpression record, VariableSymbol field) : BoundExpression(syntax, field.Type)
{
    public BoundExpression Record { get; } = record;

    public VariableSymbol Field { get; } = field;

    public override bool IsLValue => true;
}

public enum ArgumentMode
{
    /// <summary>The value is passed; the callee's changes stay in the callee.</summary>
    ByVal,

    /// <summary>The variable itself is passed (C# ref).</summary>
    ByRef,

    /// <summary>A temporary of the parameter's type is passed and copied back after the call (MS-VBAL 5.6.13.1: a Variant variable to a typed ByRef parameter, or an array element).</summary>
    ByRefCopyBack,
}

/// <summary>One argument of a call, matched to its parameter. A null value is an omitted Optional argument.</summary>
public sealed class BoundArgument(ParameterSymbol parameter, BoundExpression? value, ArgumentMode mode)
{
    public ParameterSymbol Parameter { get; } = parameter;

    public BoundExpression? Value { get; } = value;

    public ArgumentMode Mode { get; } = mode;
}

/// <summary>A call to a procedure of the project (MS-VBAL 5.4.2.1, 5.6.13).</summary>
public sealed class BoundCall(SyntaxNode syntax, ProcedureSymbol procedure, IReadOnlyList<BoundArgument> arguments, IReadOnlyList<BoundExpression> paramArrayArguments, VbaType type) : BoundExpression(syntax, type)
{
    /// <summary>The instance a class module's member is called on (MS-VBAL 5.6.13); null for a standard module's procedure.</summary>
    public BoundExpression? Receiver { get; init; }

    public ProcedureSymbol Procedure { get; } = procedure;

    /// <summary>One argument per declared parameter, ParamArray excluded.</summary>
    public IReadOnlyList<BoundArgument> Arguments { get; } = arguments;

    public IReadOnlyList<BoundExpression> ParamArrayArguments { get; } = paramArrayArguments;
}

/// <summary>A call to a standard-library function; null entries are omitted arguments.</summary>
public sealed class BoundIntrinsicCall(SyntaxNode syntax, IntrinsicSymbol intrinsic, IReadOnlyList<BoundExpression?> arguments, VbaType type) : BoundExpression(syntax, type)
{
    public IntrinsicSymbol Intrinsic { get; } = intrinsic;

    public IReadOnlyList<BoundExpression?> Arguments { get; } = arguments;

    /// <summary>The $ form (Left$, Error$): the result is a String and a Null result raises error 94.</summary>
    public bool DollarForm { get; init; }
}

/// <summary>Array(...) (MS-VBAL 6.1.2.6): the Option Base of the module decides the lower bound.</summary>
public sealed class BoundArrayFunction(SyntaxNode syntax, IReadOnlyList<BoundExpression> items, int optionBase) : BoundExpression(syntax, VbaType.Variant)
{
    public IReadOnlyList<BoundExpression> Items { get; } = items;

    public int OptionBase { get; } = optionBase;
}

/// <summary>Len or LenB of a variable with a declared non-String type: the size of the type in bytes (Strings golden).</summary>
/// <summary>Len or LenB of a user-defined type variable: the size of its file layout or of its memory layout, as an Integer (FileSystem golden).</summary>
public sealed class BoundRecordLength(SyntaxNode syntax, BoundExpression record, bool bytes) : BoundExpression(syntax, VbaType.Integer)
{
    public BoundExpression Record { get; } = record;

    public bool Bytes { get; } = bytes;
}

/// <summary>Which pointer a pointer function yields (ARCHITECTURE.md section 5, "Pointers").</summary>
public enum PointerKind
{
    /// <summary>VarPtr: the address of a variable's storage.</summary>
    Variable,

    /// <summary>StrPtr: the BSTR a String holds.</summary>
    String,

    /// <summary>ObjPtr: the interface pointer an object reference is.</summary>
    Object,
}

/// <summary>VarPtr, StrPtr, or ObjPtr of an operand, as a LongPtr (ROADMAP.md M7 E1).</summary>
public sealed class BoundPointer(SyntaxNode syntax, PointerKind kind, BoundExpression operand) : BoundExpression(syntax, VbaType.LongLong)
{
    public PointerKind Kind { get; } = kind;

    public BoundExpression Operand { get; } = operand;
}

public sealed class BoundLenOfDeclared(SyntaxNode syntax, VbaType declared, bool bytes) : BoundExpression(syntax, VbaType.Integer)
{
    public VbaType Declared { get; } = declared;

    /// <summary>The element LenB measures when it is one of a fixed-length String array's, read for its subscripts' sake; null for a variable.</summary>
    public BoundExpression? Operand { get; init; }

    public bool Bytes { get; } = bytes;
}

/// <summary>A property or method of an object: a Collection, the Err object, or something bound late on a Variant or Object.</summary>
public sealed class BoundMember(SyntaxNode syntax, BoundExpression target, string member, IReadOnlyList<BoundExpression> arguments, VbaType type) : BoundExpression(syntax, type)
{
    public BoundExpression Target { get; } = target;

    public string Member { get; } = member;

    public IReadOnlyList<BoundExpression> Arguments { get; } = arguments;

    /// <summary>obj.[name] on a late-bound receiver: the member of that name if the object has one, else its Evaluate of the text (MS-VBAL 3.3.5.3).</summary>
    public bool IsForeign { get; init; }

    /// <summary>The names of the trailing named arguments (MS-VBAL 5.6.13.1); <see cref="Arguments"/> holds the positional ones first, then these in order.</summary>
    public IReadOnlyList<string> NamedArguments { get; init; } = [];

    /// <summary>Members can be assigned through Let and Set; whether one can is decided at run time for late-bound targets.</summary>
    public override bool IsLValue => true;
}

/// <summary>The Err object (MS-VBAL 6.1.3.1).</summary>
public sealed class BoundErrObject(SyntaxNode syntax) : BoundExpression(syntax, VbaType.ErrObject);

/// <summary>Erl: the last executed line number.</summary>
public sealed class BoundErl(SyntaxNode syntax) : BoundExpression(syntax, VbaType.Long);

public enum UnaryKind
{
    Negate,
    Plus,
    Not,
}

public sealed class BoundUnary(SyntaxNode syntax, UnaryKind kind, BoundExpression operand, VbaType type) : BoundExpression(syntax, type)
{
    public UnaryKind Kind { get; } = kind;

    public BoundExpression Operand { get; } = operand;
}

public enum BinaryKind
{
    Add,
    Subtract,
    Multiply,
    Divide,
    IntegerDivide,
    Modulo,
    Power,
    Concatenate,
    Equal,
    NotEqual,
    LessThan,
    GreaterThan,
    LessThanOrEqual,
    GreaterThanOrEqual,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

public sealed class BoundBinary(SyntaxNode syntax, BinaryKind kind, BoundExpression left, BoundExpression right, DeclaredTypes declared, VbaType type) : BoundExpression(syntax, type)
{
    public BinaryKind Kind { get; } = kind;

    public BoundExpression Left { get; } = left;

    public BoundExpression Right { get; } = right;

    /// <summary>Which operands are declared Variant, which decides whether an overflow widens (MS-VBAL 5.6.9.3).</summary>
    public DeclaredTypes Declared { get; } = declared;
}

/// <summary>Let-coercion of a value to a declared type (MS-VBAL 5.5.1).</summary>
public sealed class BoundConversion(SyntaxNode syntax, BoundExpression operand, VbaType type) : BoundExpression(syntax, type)
{
    public BoundExpression Operand { get; } = operand;
}

/// <summary>New Collection (MS-VBAL 5.6.11).</summary>
public sealed class BoundNewObject(SyntaxNode syntax, VbaType type) : BoundExpression(syntax, type);

/// <summary>AddressOf procedure (MS-VBAL 5.6.16.8): the address of a procedure of this project, for a callback a DLL will call.</summary>
public sealed class BoundAddressOf(SyntaxNode syntax, ProcedureSymbol procedure) : BoundExpression(syntax, VbaType.LongLong)
{
    public ProcedureSymbol Procedure { get; } = procedure;
}

/// <summary>TypeOf expression Is type (MS-VBAL 5.6.9.10): whether the object is of that class.</summary>
public sealed class BoundTypeOf(SyntaxNode syntax, BoundExpression value, VbaType tested) : BoundExpression(syntax, VbaType.Boolean)
{
    public BoundExpression Value { get; } = value;

    public VbaType Tested { get; } = tested;
}

/// <summary>RaiseEvent name(arguments) (MS-VBAL 5.4.2.10); an argument that is an l-value receives what a ByRef handler parameter writes back.</summary>
public sealed class BoundRaiseEvent(SyntaxNode syntax, EventSymbol raised, IReadOnlyList<BoundExpression> arguments) : BoundStatement(syntax)
{
    public EventSymbol Raised { get; } = raised;

    public IReadOnlyList<BoundExpression> Arguments { get; } = arguments;
}

/// <summary>A parenthesized expression: its value, passed ByVal when it is an argument (MS-VBAL 5.6.13.1).</summary>
public sealed class BoundParenthesized(SyntaxNode syntax, BoundExpression operand) : BoundExpression(syntax, operand.Type)
{
    public BoundExpression Operand { get; } = operand;
}

public abstract class BoundStatement(SyntaxNode syntax) : BoundNode(syntax);

public sealed class BoundBlock(SyntaxNode syntax, IReadOnlyList<BoundStatement> statements) : BoundStatement(syntax)
{
    public IReadOnlyList<BoundStatement> Statements { get; } = statements;
}

/// <summary>A statement with no run-time effect: Attribute lines, Option lines inside procedures, and the like.</summary>
public sealed class BoundNop(SyntaxNode syntax) : BoundStatement(syntax);

/// <summary>Dim of locals: the variables come into existence with their default values (MS-VBAL 5.4.3.1).</summary>
public sealed class BoundLocalDeclaration(SyntaxNode syntax, IReadOnlyList<VariableSymbol> variables) : BoundStatement(syntax)
{
    public IReadOnlyList<VariableSymbol> Variables { get; } = variables;
}

public sealed class BoundAssignment(SyntaxNode syntax, BoundExpression target, BoundExpression value, bool isSet) : BoundStatement(syntax)
{
    public BoundExpression Target { get; } = target;

    /// <summary>The value, already converted to the target's declared type for Let.</summary>
    public BoundExpression Value { get; } = value;

    public bool IsSet { get; } = isSet;
}

/// <summary>A call statement (MS-VBAL 5.4.2.1); a returned value is discarded.</summary>
public sealed class BoundExpressionStatement(SyntaxNode syntax, BoundExpression expression) : BoundStatement(syntax)
{
    public BoundExpression Expression { get; } = expression;
}

public sealed record BoundBound(BoundExpression? Lower, BoundExpression Upper);

public sealed class BoundReDim(SyntaxNode syntax, BoundExpression target, IReadOnlyList<BoundBound> bounds, bool preserve, VbaType elementType) : BoundStatement(syntax)
{
    public BoundExpression Target { get; } = target;

    public IReadOnlyList<BoundBound> Bounds { get; } = bounds;

    public bool Preserve { get; } = preserve;

    /// <summary>The element type of the array a ReDim creates in a Variant.</summary>
    public VbaType ElementType { get; } = elementType;
}

public sealed class BoundErase(SyntaxNode syntax, IReadOnlyList<BoundExpression> targets) : BoundStatement(syntax)
{
    public IReadOnlyList<BoundExpression> Targets { get; } = targets;
}

public sealed class BoundMidAssignment(SyntaxNode syntax, BoundExpression target, BoundExpression start, BoundExpression? length, BoundExpression value) : BoundStatement(syntax)
{
    public BoundExpression Target { get; } = target;

    public BoundExpression Start { get; } = start;

    public BoundExpression? Length { get; } = length;

    public BoundExpression Value { get; } = value;
}

public sealed record BoundBranch(BoundExpression Condition, BoundBlock Body);

public sealed class BoundIf(SyntaxNode syntax, IReadOnlyList<BoundBranch> branches, BoundBlock? elseBody) : BoundStatement(syntax)
{
    public IReadOnlyList<BoundBranch> Branches { get; } = branches;

    public BoundBlock? ElseBody { get; } = elseBody;
}

public enum CaseClauseKind
{
    Value,
    Range,
    Comparison,
}

public sealed record BoundCaseClause(CaseClauseKind Kind, BinaryKind Operator, BoundExpression First, BoundExpression? Second);

public sealed record BoundCaseBlock(IReadOnlyList<BoundCaseClause> Clauses, BoundBlock Body);

public sealed class BoundSelect(SyntaxNode syntax, VariableSymbol selector, BoundExpression value, IReadOnlyList<BoundCaseBlock> cases, BoundBlock? elseBody) : BoundStatement(syntax)
{
    /// <summary>The temporary holding the Select expression's value, evaluated once (MS-VBAL 5.4.2.10).</summary>
    public VariableSymbol Selector { get; } = selector;

    public BoundExpression Value { get; } = value;

    public IReadOnlyList<BoundCaseBlock> Cases { get; } = cases;

    public BoundBlock? ElseBody { get; } = elseBody;
}

/// <summary>For counter = from To to [Step step] (MS-VBAL 5.4.2.3). The limit and step are evaluated once into temporaries.</summary>
public sealed class BoundFor(SyntaxNode syntax, BoundExpression counter, BoundExpression from, BoundExpression to, BoundExpression? step, VariableSymbol limit, VariableSymbol increment, BoundBlock body, SyntaxNode nextSyntax) : BoundStatement(syntax)
{
    public BoundExpression Counter { get; } = counter;

    public BoundExpression From { get; } = from;

    public BoundExpression To { get; } = to;

    public BoundExpression? Step { get; } = step;

    public VariableSymbol Limit { get; } = limit;

    public VariableSymbol Increment { get; } = increment;

    public BoundBlock Body { get; } = body;

    /// <summary>The Next statement, whose line the increment is attributed to.</summary>
    public SyntaxNode NextSyntax { get; } = nextSyntax;
}

public sealed class BoundForEach(SyntaxNode syntax, BoundExpression variable, BoundExpression collection, VariableSymbol enumerator, BoundBlock body, SyntaxNode nextSyntax) : BoundStatement(syntax)
{
    public BoundExpression Variable { get; } = variable;

    public BoundExpression Collection { get; } = collection;

    public VariableSymbol Enumerator { get; } = enumerator;

    public BoundBlock Body { get; } = body;

    public SyntaxNode NextSyntax { get; } = nextSyntax;
}

/// <summary>Do ... Loop with optional While/Until at either end (MS-VBAL 5.4.2.2); While ... Wend is a Do While (5.4.2.6).</summary>
public sealed class BoundDoLoop(SyntaxNode syntax, BoundExpression? topCondition, bool topIsUntil, BoundBlock body, BoundExpression? bottomCondition, bool bottomIsUntil, SyntaxNode loopSyntax) : BoundStatement(syntax)
{
    public BoundExpression? TopCondition { get; } = topCondition;

    public bool TopIsUntil { get; } = topIsUntil;

    public BoundBlock Body { get; } = body;

    public BoundExpression? BottomCondition { get; } = bottomCondition;

    public bool BottomIsUntil { get; } = bottomIsUntil;

    public SyntaxNode LoopSyntax { get; } = loopSyntax;
}

/// <summary>With target ... End With (MS-VBAL 5.4.2.7): the target is evaluated once into a temporary.</summary>
public sealed class BoundWith(SyntaxNode syntax, VariableSymbol temporary, BoundExpression target, BoundBlock body) : BoundStatement(syntax)
{
    public VariableSymbol Temporary { get; } = temporary;

    public BoundExpression Target { get; } = target;

    public BoundBlock Body { get; } = body;
}

public sealed class BoundGoTo(SyntaxNode syntax, LabelSymbol label) : BoundStatement(syntax)
{
    public LabelSymbol Label { get; } = label;
}

public sealed class BoundGoSub(SyntaxNode syntax, LabelSymbol label) : BoundStatement(syntax)
{
    public LabelSymbol Label { get; } = label;
}

public sealed class BoundReturn(SyntaxNode syntax) : BoundStatement(syntax);

/// <summary>On expression GoTo | GoSub labels (MS-VBAL 5.4.2.13).</summary>
public sealed class BoundOnGoTo(SyntaxNode syntax, BoundExpression selector, IReadOnlyList<LabelSymbol> labels, bool isGoSub) : BoundStatement(syntax)
{
    public BoundExpression Selector { get; } = selector;

    public IReadOnlyList<LabelSymbol> Labels { get; } = labels;

    public bool IsGoSub { get; } = isGoSub;
}

public sealed class BoundLabel(SyntaxNode syntax, LabelSymbol label) : BoundStatement(syntax)
{
    public LabelSymbol Label { get; } = label;
}

public enum OnErrorMode
{
    ResumeNext,

    /// <summary>On Error GoTo 0: no handler.</summary>
    Disable,

    /// <summary>On Error GoTo -1: dismiss the active error, keep the handler.</summary>
    Dismiss,

    GoToLabel,
}

public sealed class BoundOnError(SyntaxNode syntax, OnErrorMode mode, LabelSymbol? label) : BoundStatement(syntax)
{
    public OnErrorMode Mode { get; } = mode;

    public LabelSymbol? Label { get; } = label;
}

public enum ResumeKind
{
    /// <summary>Resume: retry the statement that raised the error.</summary>
    Retry,
    Next,
    Label,
}

public sealed class BoundResume(SyntaxNode syntax, ResumeKind kind, LabelSymbol? label) : BoundStatement(syntax)
{
    public ResumeKind Kind { get; } = kind;

    public LabelSymbol? Label { get; } = label;
}

public enum ExitKind
{
    Procedure,
    For,
    Do,
}

public sealed class BoundExit(SyntaxNode syntax, ExitKind kind) : BoundStatement(syntax)
{
    public ExitKind Kind { get; } = kind;
}

/// <summary>Error number (MS-VBAL 5.4.4.3).</summary>
/// <summary>LSet target = value or RSet target = value (MS-VBAL 5.4.3.6, 5.4.3.7): a String target takes the value justified in its current length; an LSet between records copies the memory image.</summary>
public sealed class BoundLSet(SyntaxNode syntax, BoundExpression target, BoundExpression value, bool isRSet) : BoundStatement(syntax)
{
    public BoundExpression Target { get; } = target;

    public BoundExpression Value { get; } = value;

    public bool IsRSet { get; } = isRSet;
}

public sealed class BoundErrorStatement(SyntaxNode syntax, BoundExpression number) : BoundStatement(syntax)
{
    public BoundExpression Number { get; } = number;
}

public sealed class BoundEnd(SyntaxNode syntax) : BoundStatement(syntax);

public sealed class BoundStop(SyntaxNode syntax) : BoundStatement(syntax);

public enum PrintItemKind
{
    Value,
    Spc,
    Tab,
}

/// <summary>One item of a Debug.Print list (MS-VBAL 5.4.5.6 output-list): a value, Spc(n), or Tab[(n)], followed by ";" or "," or the end of the line.</summary>
public sealed record BoundPrintItem(PrintItemKind Kind, BoundExpression? Value, char? Separator);

public sealed class BoundDebugPrint(SyntaxNode syntax, IReadOnlyList<BoundPrintItem> items) : BoundStatement(syntax)
{
    public IReadOnlyList<BoundPrintItem> Items { get; } = items;
}

/// <summary>Open pathname For mode [Access access] [lock] As #filenumber [Len = reclength] (MS-VBAL 5.4.5.1); the clauses name runtime enum members.</summary>
public sealed class BoundOpen(SyntaxNode syntax, BoundExpression pathName, string mode, string access, string lockMode, BoundExpression fileNumber, BoundExpression? recordLength) : BoundStatement(syntax)
{
    public BoundExpression PathName { get; } = pathName;

    public string Mode { get; } = mode;

    public string Access { get; } = access;

    public string LockMode { get; } = lockMode;

    public BoundExpression FileNumber { get; } = fileNumber;

    public BoundExpression? RecordLength { get; } = recordLength;
}

public enum FileStatementKind
{
    Close,
    Reset,
    Seek,
    Lock,
    Unlock,
    Width,
    Name,
}

/// <summary>Close, Reset, Seek, Lock, Unlock, Width, and Name (MS-VBAL 5.4.5.1 to 5.4.5.5): the file number where the statement has one, and its other operands in order, Missing for a left-out one.</summary>
public sealed class BoundFileStatement(SyntaxNode syntax, FileStatementKind kind, BoundExpression? fileNumber, IReadOnlyList<BoundExpression?> operands) : BoundStatement(syntax)
{
    public FileStatementKind Kind { get; } = kind;

    public BoundExpression? FileNumber { get; } = fileNumber;

    public IReadOnlyList<BoundExpression?> Operands { get; } = operands;
}

/// <summary>Print # or Write # (MS-VBAL 5.4.5.6, 5.4.5.7).</summary>
public sealed class BoundFileOutput(SyntaxNode syntax, bool isWrite, BoundExpression fileNumber, IReadOnlyList<BoundPrintItem> items) : BoundStatement(syntax)
{
    public bool IsWrite { get; } = isWrite;

    public BoundExpression FileNumber { get; } = fileNumber;

    public IReadOnlyList<BoundPrintItem> Items { get; } = items;
}

/// <summary>Input # or Line Input # into assignable targets (MS-VBAL 5.4.5.8, 5.4.5.4).</summary>
public sealed class BoundFileInput(SyntaxNode syntax, bool isLineInput, BoundExpression fileNumber, IReadOnlyList<BoundExpression> targets) : BoundStatement(syntax)
{
    public bool IsLineInput { get; } = isLineInput;

    public BoundExpression FileNumber { get; } = fileNumber;

    public IReadOnlyList<BoundExpression> Targets { get; } = targets;
}

/// <summary>Get # or Put # of a scalar, a string, a record, or an array (MS-VBAL 5.4.5.9); the record number is Missing when left out.</summary>
public sealed class BoundFileRecord(SyntaxNode syntax, bool isPut, BoundExpression fileNumber, BoundExpression recordNumber, BoundExpression variable) : BoundStatement(syntax)
{
    public bool IsPut { get; } = isPut;

    public BoundExpression FileNumber { get; } = fileNumber;

    public BoundExpression RecordNumber { get; } = recordNumber;

    public BoundExpression Variable { get; } = variable;
}

/// <summary>A bound procedure and what the emitter needs to know about it.</summary>
public sealed class BoundProcedure(ProcedureSymbol symbol, BoundBlock body, IReadOnlyList<VariableSymbol> locals, IReadOnlyList<LabelSymbol> labels, bool usesDispatch, bool usesGoSub)
{
    public ProcedureSymbol Symbol { get; } = symbol;

    public BoundBlock Body { get; } = body;

    /// <summary>Locals and temporaries, declaration order; Static locals live in the module instead.</summary>
    public IReadOnlyList<VariableSymbol> Locals { get; } = locals;

    public IReadOnlyList<LabelSymbol> Labels { get; } = labels;

    /// <summary>The procedure uses On Error, Resume, GoTo, GoSub, labels, or Erl and compiles to the dispatch loop (ARCHITECTURE.md section 4).</summary>
    public bool UsesDispatch { get; } = usesDispatch;

    public bool UsesGoSub { get; } = usesGoSub;
}

public sealed class BoundModule(ModuleSymbol symbol, IReadOnlyList<BoundProcedure> procedures)
{
    public ModuleSymbol Symbol { get; } = symbol;

    public IReadOnlyList<BoundProcedure> Procedures { get; } = procedures;
}

public sealed class BoundProject(ProjectSymbol symbol, IReadOnlyList<BoundModule> modules)
{
    public ProjectSymbol Symbol { get; } = symbol;

    public IReadOnlyList<BoundModule> Modules { get; } = modules;
}

/// <summary>
/// A member of a COM object bound by dispid at compile time (ARCHITECTURE.md section 6): a
/// method call or property read on a target of a library type. Assignable when the property has
/// a put accessor, or when its result is an object whose default member has one, as in
/// <c>Range("A1") = 5</c>.
/// </summary>
public sealed class BoundComCall(
    SyntaxNode syntax,
    BoundExpression target,
    Runtime.TypeLibraries.ComMember member,
    IReadOnlyList<BoundExpression> arguments,
    Runtime.InvokeKind kind,
    VbaType type,
    Runtime.TypeLibraries.ComMember? putMember,
    Runtime.TypeLibraries.ComMember? putRefMember,
    Runtime.TypeLibraries.ComMember? defaultPut,
    VbaType? putValueType) : BoundExpression(syntax, type)
{
    public BoundExpression Target { get; } = target;

    public Runtime.TypeLibraries.ComMember Member { get; } = member;

    /// <summary>Positional arguments, Missing for omitted optional ones between supplied ones.</summary>
    public IReadOnlyList<BoundExpression> Arguments { get; } = arguments;

    public Runtime.InvokeKind Kind { get; } = kind;

    /// <summary>The property's Let accessor, when it has one.</summary>
    public Runtime.TypeLibraries.ComMember? PutMember { get; } = putMember;

    /// <summary>The property's Set accessor, when it has one.</summary>
    public Runtime.TypeLibraries.ComMember? PutRefMember { get; } = putRefMember;

    /// <summary>The Let accessor of the result's default member, for assignments to the object the call returns.</summary>
    public Runtime.TypeLibraries.ComMember? DefaultPut { get; } = defaultPut;

    /// <summary>The declared type of the value a put accepts, for let-coercion before the call.</summary>
    public VbaType? PutValueType { get; } = putValueType;

    public override bool IsLValue => PutMember is not null || PutRefMember is not null || DefaultPut is not null;
}

/// <summary>A library's application object (Excel's Global): the implicit target of its members used unqualified.</summary>
public sealed class BoundAppObject(SyntaxNode syntax, Runtime.TypeLibraries.ComLibrary library, Runtime.TypeLibraries.ComType appObject, VbaType type) : BoundExpression(syntax, type)
{
    public Runtime.TypeLibraries.ComLibrary Library { get; } = library;

    public Runtime.TypeLibraries.ComType AppObject { get; } = appObject;
}
