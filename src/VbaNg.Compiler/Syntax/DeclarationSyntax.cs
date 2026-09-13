namespace VbaNg.Compiler.Syntax;

/// <summary>Base of every statement and declaration; module members and procedure bodies are lists of these.</summary>
public abstract class StatementSyntax : SyntaxNode
{
    protected StatementSyntax(SyntaxKind kind)
        : base(kind)
    {
    }
}

/// <summary>A whole module (MS-VBAL 4.2, 5.1): optional raw header, then attributes, declarations, and procedures.</summary>
public sealed class ModuleSyntax : SyntaxNode
{
    public ModuleSyntax(ModuleHeaderSyntax? header, SyntaxList<StatementSyntax> members, SyntaxToken endOfFileToken)
        : base(SyntaxKind.Module)
    {
        Header = header;
        Members = members ?? throw new ArgumentNullException(nameof(members));
        EndOfFileToken = endOfFileToken ?? throw new ArgumentNullException(nameof(endOfFileToken));
    }

    public ModuleHeaderSyntax? Header { get; }

    public SyntaxList<StatementSyntax> Members { get; }

    public SyntaxToken EndOfFileToken { get; }

    public IEnumerable<ProcedureDeclarationSyntax> Procedures => Members.OfType<ProcedureDeclarationSyntax>();

    /// <summary>The VB_Name attribute value, or null when the module has none.</summary>
    public string? Name =>
        Members.OfType<AttributeStatementSyntax>()
            .Where(a => a.Name is IdentifierNameSyntax { Name: var n } && n.Equals("VB_Name", StringComparison.OrdinalIgnoreCase))
            .Select(a => a.Values.Nodes.FirstOrDefault() is LiteralExpressionSyntax { Token.Value: string value } ? value : null)
            .FirstOrDefault();

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(Header), Members.AsChildren(), One(EndOfFileToken));
}

/// <summary>The VERSION line and BEGIN...END block the VBE writes before a class or form module (MS-VBAL 4.2); not VBA grammar.</summary>
public sealed class ModuleHeaderSyntax : SyntaxNode
{
    public ModuleHeaderSyntax(SyntaxTokenList lines)
        : base(SyntaxKind.ModuleHeader)
    {
        Lines = lines ?? throw new ArgumentNullException(nameof(lines));
    }

    public SyntaxTokenList Lines { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Lines.AsChildren();
}

/// <summary>"Attribute" name "=" values (MS-VBAL 4.2, 5.2.4.1); also written after Sub lines and variable declarations.</summary>
public sealed class AttributeStatementSyntax : StatementSyntax
{
    public AttributeStatementSyntax(SyntaxToken attributeKeyword, ExpressionSyntax name, SyntaxToken equalsToken, SeparatedSyntaxList<ExpressionSyntax> values)
        : base(SyntaxKind.AttributeStatement)
    {
        AttributeKeyword = attributeKeyword ?? throw new ArgumentNullException(nameof(attributeKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        Values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public SyntaxToken AttributeKeyword { get; }

    /// <summary>VB_Name, or a dotted name such as Main.VB_ProcData.VB_Invoke_Func.</summary>
    public ExpressionSyntax Name { get; }

    public SyntaxToken EqualsToken { get; }

    public SeparatedSyntaxList<ExpressionSyntax> Values { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(AttributeKeyword), One(Name), One(EqualsToken), Values.AsChildren());
}

/// <summary>Option Explicit, Option Base n, Option Compare mode, Option Private Module (MS-VBAL 5.2.1).</summary>
public sealed class OptionStatementSyntax : StatementSyntax
{
    public OptionStatementSyntax(SyntaxToken optionKeyword, SyntaxToken optionName, SyntaxToken? argument)
        : base(SyntaxKind.OptionStatement)
    {
        OptionKeyword = optionKeyword ?? throw new ArgumentNullException(nameof(optionKeyword));
        OptionName = optionName ?? throw new ArgumentNullException(nameof(optionName));
        Argument = argument;
    }

    public SyntaxToken OptionKeyword { get; }

    /// <summary>Explicit, Base, Compare, or Private.</summary>
    public SyntaxToken OptionName { get; }

    /// <summary>The 0 or 1 of Option Base, the Binary/Text/Database of Option Compare, the Module of Option Private.</summary>
    public SyntaxToken? Argument { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(OptionKeyword, OptionName, Argument);
}

/// <summary>DefInt A-Z and friends (MS-VBAL 5.2.2).</summary>
public sealed class DefTypeStatementSyntax : StatementSyntax
{
    public DefTypeStatementSyntax(SyntaxToken defKeyword, SeparatedSyntaxList<LetterRangeSyntax> ranges)
        : base(SyntaxKind.DefTypeStatement)
    {
        DefKeyword = defKeyword ?? throw new ArgumentNullException(nameof(defKeyword));
        Ranges = ranges ?? throw new ArgumentNullException(nameof(ranges));
    }

    public SyntaxToken DefKeyword { get; }

    public SeparatedSyntaxList<LetterRangeSyntax> Ranges { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(DefKeyword), Ranges.AsChildren());
}

/// <summary>A letter or letter range in a DefType statement.</summary>
public sealed class LetterRangeSyntax : SyntaxNode
{
    public LetterRangeSyntax(SyntaxToken first, SyntaxToken? minusToken, SyntaxToken? last)
        : base(SyntaxKind.LetterRange)
    {
        First = first ?? throw new ArgumentNullException(nameof(first));
        MinusToken = minusToken;
        Last = last;
    }

    public SyntaxToken First { get; }

    public SyntaxToken? MinusToken { get; }

    public SyntaxToken? Last { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(First, MinusToken, Last);
}

/// <summary>Dim, Static, Private, Public, Global variable declarations (MS-VBAL 5.2.3.1, 5.4.3.1).</summary>
public sealed class VariableDeclarationSyntax : StatementSyntax
{
    public VariableDeclarationSyntax(SyntaxTokenList modifiers, SeparatedSyntaxList<VariableDeclaratorSyntax> declarators)
        : base(SyntaxKind.VariableDeclaration)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        Declarators = declarators ?? throw new ArgumentNullException(nameof(declarators));
    }

    /// <summary>Dim, Static, Private, Public, or Global, possibly followed by Shared (ignored by VBA).</summary>
    public SyntaxTokenList Modifiers { get; }

    public SeparatedSyntaxList<VariableDeclaratorSyntax> Declarators { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Modifiers.AsChildren(), Declarators.AsChildren());
}

/// <summary>[WithEvents] name [(bounds)] [As [New] type].</summary>
public sealed class VariableDeclaratorSyntax : SyntaxNode
{
    public VariableDeclaratorSyntax(SyntaxToken? withEventsKeyword, SyntaxToken name, ArrayBoundsSyntax? bounds, AsClauseSyntax? asClause)
        : base(SyntaxKind.VariableDeclarator)
    {
        WithEventsKeyword = withEventsKeyword;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Bounds = bounds;
        AsClause = asClause;
    }

    public SyntaxToken? WithEventsKeyword { get; }

    public SyntaxToken Name { get; }

    public ArrayBoundsSyntax? Bounds { get; }

    public AsClauseSyntax? AsClause { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(WithEventsKeyword, Name, Bounds, AsClause);
}

/// <summary>[Public | Private] Const declarators (MS-VBAL 5.2.3.2, 5.4.3.2).</summary>
public sealed class ConstDeclarationSyntax : StatementSyntax
{
    public ConstDeclarationSyntax(SyntaxTokenList modifiers, SyntaxToken constKeyword, SeparatedSyntaxList<ConstDeclaratorSyntax> declarators)
        : base(SyntaxKind.ConstDeclaration)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        ConstKeyword = constKeyword ?? throw new ArgumentNullException(nameof(constKeyword));
        Declarators = declarators ?? throw new ArgumentNullException(nameof(declarators));
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken ConstKeyword { get; }

    public SeparatedSyntaxList<ConstDeclaratorSyntax> Declarators { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Modifiers.AsChildren(), One(ConstKeyword), Declarators.AsChildren());
}

/// <summary>name [As type] = expression.</summary>
public sealed class ConstDeclaratorSyntax : SyntaxNode
{
    public ConstDeclaratorSyntax(SyntaxToken name, AsClauseSyntax? asClause, SyntaxToken equalsToken, ExpressionSyntax value)
        : base(SyntaxKind.ConstDeclarator)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        AsClause = asClause;
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public SyntaxToken Name { get; }

    public AsClauseSyntax? AsClause { get; }

    public SyntaxToken EqualsToken { get; }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Name, AsClause, EqualsToken, Value);
}

/// <summary>"(" [bound, ...] ")" after an array name (MS-VBAL 5.2.3.1.3); empty for a dynamic array.</summary>
public sealed class ArrayBoundsSyntax : SyntaxNode
{
    public ArrayBoundsSyntax(SyntaxToken openParenToken, SeparatedSyntaxList<BoundSyntax> bounds, SyntaxToken closeParenToken)
        : base(SyntaxKind.ArrayBounds)
    {
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        Bounds = bounds ?? throw new ArgumentNullException(nameof(bounds));
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
    }

    public SyntaxToken OpenParenToken { get; }

    public SeparatedSyntaxList<BoundSyntax> Bounds { get; }

    public SyntaxToken CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(OpenParenToken), Bounds.AsChildren(), One(CloseParenToken));
}

/// <summary>[lower To] upper.</summary>
public sealed class BoundSyntax : SyntaxNode
{
    public BoundSyntax(ExpressionSyntax? lower, SyntaxToken? toKeyword, ExpressionSyntax upper)
        : base(SyntaxKind.Bound)
    {
        Lower = lower;
        ToKeyword = toKeyword;
        Upper = upper ?? throw new ArgumentNullException(nameof(upper));
    }

    public ExpressionSyntax? Lower { get; }

    public SyntaxToken? ToKeyword { get; }

    public ExpressionSyntax Upper { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Lower, ToKeyword, Upper);
}

/// <summary>"As" [New] type [()] (MS-VBAL 5.2.3.1.1, 5.3.1.4). The array designator only follows a function or property return type.</summary>
public sealed class AsClauseSyntax : SyntaxNode
{
    public AsClauseSyntax(SyntaxToken asKeyword, SyntaxToken? newKeyword, TypeSyntax type, ArrayDesignatorSyntax? arrayDesignator)
        : base(SyntaxKind.AsClause)
    {
        AsKeyword = asKeyword ?? throw new ArgumentNullException(nameof(asKeyword));
        NewKeyword = newKeyword;
        Type = type ?? throw new ArgumentNullException(nameof(type));
        ArrayDesignator = arrayDesignator;
    }

    public SyntaxToken AsKeyword { get; }

    public SyntaxToken? NewKeyword { get; }

    public TypeSyntax Type { get; }

    public ArrayDesignatorSyntax? ArrayDesignator { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(AsKeyword, NewKeyword, Type, ArrayDesignator);
}

/// <summary>"(" ")" marking an array return type.</summary>
public sealed class ArrayDesignatorSyntax : SyntaxNode
{
    public ArrayDesignatorSyntax(SyntaxToken openParenToken, SyntaxToken closeParenToken)
        : base(SyntaxKind.ArrayDesignator)
    {
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
    }

    public SyntaxToken OpenParenToken { get; }

    public SyntaxToken CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(OpenParenToken, CloseParenToken);
}

/// <summary>Base of type references in As clauses (MS-VBAL 5.6.16.7 type-expression).</summary>
public abstract class TypeSyntax : SyntaxNode
{
    protected TypeSyntax(SyntaxKind kind)
        : base(kind)
    {
    }
}

/// <summary>A reserved type name, optionally "String * length" for a fixed-length string (MS-VBAL 5.2.3.1.1).</summary>
public sealed class BuiltinTypeSyntax : TypeSyntax
{
    public BuiltinTypeSyntax(SyntaxToken keyword, SyntaxToken? asteriskToken, ExpressionSyntax? length)
        : base(SyntaxKind.BuiltinType)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        AsteriskToken = asteriskToken;
        Length = length;
    }

    public SyntaxToken Keyword { get; }

    public SyntaxToken? AsteriskToken { get; }

    public ExpressionSyntax? Length { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, AsteriskToken, Length);
}

/// <summary>A program-defined or library type, possibly qualified (Excel.Range).</summary>
public sealed class NamedTypeSyntax : TypeSyntax
{
    public NamedTypeSyntax(ExpressionSyntax name)
        : base(SyntaxKind.NamedType)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    public ExpressionSyntax Name { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Name);
}

/// <summary>[Public | Private] Type name ... End Type (MS-VBAL 5.2.3.3).</summary>
public sealed class TypeDefinitionSyntax : StatementSyntax
{
    public TypeDefinitionSyntax(SyntaxTokenList modifiers, SyntaxToken typeKeyword, SyntaxToken name, SyntaxList<StatementSyntax> members, EndBlockStatementSyntax endType)
        : base(SyntaxKind.TypeDefinition)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        TypeKeyword = typeKeyword ?? throw new ArgumentNullException(nameof(typeKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Members = members ?? throw new ArgumentNullException(nameof(members));
        EndType = endType ?? throw new ArgumentNullException(nameof(endType));
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken TypeKeyword { get; }

    public SyntaxToken Name { get; }

    public SyntaxList<StatementSyntax> Members { get; }

    public EndBlockStatementSyntax EndType { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Modifiers.AsChildren(), One(TypeKeyword), One(Name), Members.AsChildren(), One(EndType));
}

/// <summary>A field of a user-defined type: name [(bounds)] As type.</summary>
public sealed class TypeMemberSyntax : StatementSyntax
{
    public TypeMemberSyntax(SyntaxToken name, ArrayBoundsSyntax? bounds, AsClauseSyntax asClause)
        : base(SyntaxKind.TypeMember)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Bounds = bounds;
        AsClause = asClause ?? throw new ArgumentNullException(nameof(asClause));
    }

    public SyntaxToken Name { get; }

    public ArrayBoundsSyntax? Bounds { get; }

    public AsClauseSyntax AsClause { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Name, Bounds, AsClause);
}

/// <summary>[Public | Private] Enum name ... End Enum (MS-VBAL 5.2.3.4).</summary>
public sealed class EnumDefinitionSyntax : StatementSyntax
{
    public EnumDefinitionSyntax(SyntaxTokenList modifiers, SyntaxToken enumKeyword, SyntaxToken name, SyntaxList<StatementSyntax> members, EndBlockStatementSyntax endEnum)
        : base(SyntaxKind.EnumDefinition)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        EnumKeyword = enumKeyword ?? throw new ArgumentNullException(nameof(enumKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Members = members ?? throw new ArgumentNullException(nameof(members));
        EndEnum = endEnum ?? throw new ArgumentNullException(nameof(endEnum));
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken EnumKeyword { get; }

    public SyntaxToken Name { get; }

    public SyntaxList<StatementSyntax> Members { get; }

    public EndBlockStatementSyntax EndEnum { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Modifiers.AsChildren(), One(EnumKeyword), One(Name), Members.AsChildren(), One(EndEnum));
}

/// <summary>An enum member: name [= expression].</summary>
public sealed class EnumMemberSyntax : StatementSyntax
{
    public EnumMemberSyntax(SyntaxToken name, SyntaxToken? equalsToken, ExpressionSyntax? value)
        : base(SyntaxKind.EnumMember)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        EqualsToken = equalsToken;
        Value = value;
    }

    public SyntaxToken Name { get; }

    public SyntaxToken? EqualsToken { get; }

    public ExpressionSyntax? Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Name, EqualsToken, Value);
}

/// <summary>[Public | Private] Declare [PtrSafe] Sub | Function name [CDecl] Lib "lib" [Alias "alias"] [(params)] [As type] (MS-VBAL 5.2.3.5).</summary>
public sealed class DeclareStatementSyntax : StatementSyntax
{
    public DeclareStatementSyntax(
        SyntaxTokenList modifiers,
        SyntaxToken declareKeyword,
        SyntaxToken? ptrSafeKeyword,
        SyntaxToken procedureKeyword,
        SyntaxToken name,
        SyntaxToken? cdeclKeyword,
        SyntaxToken libKeyword,
        SyntaxToken libraryName,
        SyntaxToken? aliasKeyword,
        SyntaxToken? aliasName,
        ParameterListSyntax? parameters,
        AsClauseSyntax? asClause)
        : base(SyntaxKind.DeclareStatement)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        DeclareKeyword = declareKeyword ?? throw new ArgumentNullException(nameof(declareKeyword));
        PtrSafeKeyword = ptrSafeKeyword;
        ProcedureKeyword = procedureKeyword ?? throw new ArgumentNullException(nameof(procedureKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        CDeclKeyword = cdeclKeyword;
        LibKeyword = libKeyword ?? throw new ArgumentNullException(nameof(libKeyword));
        LibraryName = libraryName ?? throw new ArgumentNullException(nameof(libraryName));
        AliasKeyword = aliasKeyword;
        AliasName = aliasName;
        Parameters = parameters;
        AsClause = asClause;
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken DeclareKeyword { get; }

    public SyntaxToken? PtrSafeKeyword { get; }

    /// <summary>Sub or Function.</summary>
    public SyntaxToken ProcedureKeyword { get; }

    public SyntaxToken Name { get; }

    public SyntaxToken? CDeclKeyword { get; }

    public SyntaxToken LibKeyword { get; }

    public SyntaxToken LibraryName { get; }

    public SyntaxToken? AliasKeyword { get; }

    public SyntaxToken? AliasName { get; }

    public ParameterListSyntax? Parameters { get; }

    public AsClauseSyntax? AsClause { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(
            Modifiers.AsChildren(),
            One(DeclareKeyword),
            One(PtrSafeKeyword),
            One(ProcedureKeyword),
            One(Name),
            One(CDeclKeyword),
            One(LibKeyword),
            One(LibraryName),
            One(AliasKeyword),
            One(AliasName),
            One(Parameters),
            One(AsClause));
}

/// <summary>[Public] Event name [(params)] (MS-VBAL 5.2.4.3).</summary>
public sealed class EventDeclarationSyntax : StatementSyntax
{
    public EventDeclarationSyntax(SyntaxTokenList modifiers, SyntaxToken eventKeyword, SyntaxToken name, ParameterListSyntax? parameters)
        : base(SyntaxKind.EventDeclaration)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        EventKeyword = eventKeyword ?? throw new ArgumentNullException(nameof(eventKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Parameters = parameters;
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken EventKeyword { get; }

    public SyntaxToken Name { get; }

    public ParameterListSyntax? Parameters { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Modifiers.AsChildren(), One(EventKeyword), One(Name), One(Parameters));
}

/// <summary>Implements class (MS-VBAL 5.2.4.2).</summary>
public sealed class ImplementsStatementSyntax : StatementSyntax
{
    public ImplementsStatementSyntax(SyntaxToken implementsKeyword, ExpressionSyntax name)
        : base(SyntaxKind.ImplementsStatement)
    {
        ImplementsKeyword = implementsKeyword ?? throw new ArgumentNullException(nameof(implementsKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    public SyntaxToken ImplementsKeyword { get; }

    public ExpressionSyntax Name { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(ImplementsKeyword, Name);
}

/// <summary>Sub, Function, and Property Get/Let/Set declarations with their bodies (MS-VBAL 5.3.1).</summary>
public sealed class ProcedureDeclarationSyntax : StatementSyntax
{
    public ProcedureDeclarationSyntax(
        SyntaxTokenList modifiers,
        SyntaxToken procedureKeyword,
        SyntaxToken? accessorKeyword,
        SyntaxToken name,
        ParameterListSyntax? parameters,
        AsClauseSyntax? asClause,
        SyntaxList<StatementSyntax> body,
        EndBlockStatementSyntax endStatement)
        : base(SyntaxKind.ProcedureDeclaration)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        ProcedureKeyword = procedureKeyword ?? throw new ArgumentNullException(nameof(procedureKeyword));
        AccessorKeyword = accessorKeyword;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Parameters = parameters;
        AsClause = asClause;
        Body = body ?? throw new ArgumentNullException(nameof(body));
        EndStatement = endStatement ?? throw new ArgumentNullException(nameof(endStatement));
    }

    /// <summary>Public, Private, Friend, Static, in source order.</summary>
    public SyntaxTokenList Modifiers { get; }

    /// <summary>Sub, Function, or the identifier Property.</summary>
    public SyntaxToken ProcedureKeyword { get; }

    /// <summary>Get, Let, or Set for properties.</summary>
    public SyntaxToken? AccessorKeyword { get; }

    public SyntaxToken Name { get; }

    public ParameterListSyntax? Parameters { get; }

    public AsClauseSyntax? AsClause { get; }

    public SyntaxList<StatementSyntax> Body { get; }

    public EndBlockStatementSyntax EndStatement { get; }

    public bool IsSub => ProcedureKeyword.Kind == SyntaxKind.SubKeyword;

    public bool IsFunction => ProcedureKeyword.Kind == SyntaxKind.FunctionKeyword;

    public bool IsProperty => AccessorKeyword is not null;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(
            Modifiers.AsChildren(),
            One(ProcedureKeyword),
            One(AccessorKeyword),
            One(Name),
            One(Parameters),
            One(AsClause),
            Body.AsChildren(),
            One(EndStatement));
}

/// <summary>"(" parameters ")" (MS-VBAL 5.3.1.5).</summary>
public sealed class ParameterListSyntax : SyntaxNode
{
    public ParameterListSyntax(SyntaxToken openParenToken, SeparatedSyntaxList<ParameterSyntax> parameters, SyntaxToken closeParenToken)
        : base(SyntaxKind.ParameterList)
    {
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        Parameters = parameters ?? throw new ArgumentNullException(nameof(parameters));
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
    }

    public SyntaxToken OpenParenToken { get; }

    public SeparatedSyntaxList<ParameterSyntax> Parameters { get; }

    public SyntaxToken CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(OpenParenToken), Parameters.AsChildren(), One(CloseParenToken));
}

/// <summary>[Optional] [ByVal | ByRef] [ParamArray] name [()] [As type] [= default] (MS-VBAL 5.3.1.5).</summary>
public sealed class ParameterSyntax : SyntaxNode
{
    public ParameterSyntax(SyntaxTokenList modifiers, SyntaxToken name, ArrayBoundsSyntax? bounds, AsClauseSyntax? asClause, SyntaxToken? equalsToken, ExpressionSyntax? defaultValue)
        : base(SyntaxKind.Parameter)
    {
        Modifiers = modifiers ?? throw new ArgumentNullException(nameof(modifiers));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Bounds = bounds;
        AsClause = asClause;
        EqualsToken = equalsToken;
        DefaultValue = defaultValue;
    }

    public SyntaxTokenList Modifiers { get; }

    public SyntaxToken Name { get; }

    public ArrayBoundsSyntax? Bounds { get; }

    public AsClauseSyntax? AsClause { get; }

    public SyntaxToken? EqualsToken { get; }

    public ExpressionSyntax? DefaultValue { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Modifiers.AsChildren(), One(Name), One(Bounds), One(AsClause), One(EqualsToken), One(DefaultValue));
}

/// <summary>"End" Sub | Function | Property | If | Select | With | Type | Enum, or the single token EndIf.</summary>
public sealed class EndBlockStatementSyntax : StatementSyntax
{
    public EndBlockStatementSyntax(SyntaxToken endKeyword, SyntaxToken? blockKeyword)
        : base(SyntaxKind.EndBlockStatement)
    {
        EndKeyword = endKeyword ?? throw new ArgumentNullException(nameof(endKeyword));
        BlockKeyword = blockKeyword;
    }

    public SyntaxToken EndKeyword { get; }

    public SyntaxToken? BlockKeyword { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(EndKeyword, BlockKeyword);
}
