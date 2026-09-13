namespace VbaNg.Compiler.Syntax;

/// <summary>Base of expressions (MS-VBAL 5.6).</summary>
public abstract class ExpressionSyntax : SyntaxNode
{
    protected ExpressionSyntax(SyntaxKind kind)
        : base(kind)
    {
    }
}

/// <summary>A literal token: number, string, date, True, False, Nothing, Empty, Null (MS-VBAL 5.6.5).</summary>
public sealed class LiteralExpressionSyntax : ExpressionSyntax
{
    public LiteralExpressionSyntax(SyntaxToken token)
        : base(SyntaxKind.LiteralExpression)
    {
        Token = token ?? throw new ArgumentNullException(nameof(token));
    }

    public SyntaxToken Token { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Token);
}

/// <summary>
/// A simple name (MS-VBAL 5.6.10): an identifier, typed name, foreign name, Me, or a keyword
/// used as a name after "." or "!" (unrestricted-name).
/// </summary>
public sealed class IdentifierNameSyntax : ExpressionSyntax
{
    public IdentifierNameSyntax(SyntaxToken identifier)
        : base(SyntaxKind.IdentifierName)
    {
        Identifier = identifier ?? throw new ArgumentNullException(nameof(identifier));
    }

    public SyntaxToken Identifier { get; }

    public string Name => Identifier.NameValue;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Identifier);
}

/// <summary>expression "." name (MS-VBAL 5.6.12); with a null expression, the "." member of a With block (5.6.14).</summary>
public sealed class MemberAccessExpressionSyntax : ExpressionSyntax
{
    public MemberAccessExpressionSyntax(ExpressionSyntax? expression, SyntaxToken dotToken, IdentifierNameSyntax name)
        : base(SyntaxKind.MemberAccessExpression)
    {
        Expression = expression;
        DotToken = dotToken ?? throw new ArgumentNullException(nameof(dotToken));
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    public ExpressionSyntax? Expression { get; }

    public SyntaxToken DotToken { get; }

    public IdentifierNameSyntax Name { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Expression, DotToken, Name);
}

/// <summary>expression "!" name (MS-VBAL 5.6.15 dictionary access); a null expression means the With block object.</summary>
public sealed class DictionaryAccessExpressionSyntax : ExpressionSyntax
{
    public DictionaryAccessExpressionSyntax(ExpressionSyntax? expression, SyntaxToken bangToken, IdentifierNameSyntax name)
        : base(SyntaxKind.DictionaryAccessExpression)
    {
        Expression = expression;
        BangToken = bangToken ?? throw new ArgumentNullException(nameof(bangToken));
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    public ExpressionSyntax? Expression { get; }

    public SyntaxToken BangToken { get; }

    public IdentifierNameSyntax Name { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Expression, BangToken, Name);
}

/// <summary>expression "(" arguments ")" (MS-VBAL 5.6.13): a call, an array index, or a default member access; binding decides.</summary>
public sealed class IndexExpressionSyntax : ExpressionSyntax
{
    public IndexExpressionSyntax(ExpressionSyntax expression, ArgumentListSyntax arguments)
        : base(SyntaxKind.IndexExpression)
    {
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        Arguments = arguments ?? throw new ArgumentNullException(nameof(arguments));
    }

    public ExpressionSyntax Expression { get; }

    public ArgumentListSyntax Arguments { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Expression, Arguments);
}

/// <summary>An argument list (MS-VBAL 5.6.13.1), parenthesized or, in a call statement, bare.</summary>
public sealed class ArgumentListSyntax : SyntaxNode
{
    public ArgumentListSyntax(SyntaxToken? openParenToken, SeparatedSyntaxList<ArgumentSyntax> arguments, SyntaxToken? closeParenToken)
        : base(SyntaxKind.ArgumentList)
    {
        OpenParenToken = openParenToken;
        Arguments = arguments ?? throw new ArgumentNullException(nameof(arguments));
        CloseParenToken = closeParenToken;
    }

    public SyntaxToken? OpenParenToken { get; }

    public SeparatedSyntaxList<ArgumentSyntax> Arguments { get; }

    public SyntaxToken? CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(OpenParenToken), Arguments.AsChildren(), One(CloseParenToken));
}

/// <summary>
/// One argument (MS-VBAL 5.6.13.1): positional or named ("name := value"), optionally ByVal, and
/// possibly omitted (all parts null) as in "f(1, , 3)".
/// </summary>
public sealed class ArgumentSyntax : SyntaxNode
{
    public ArgumentSyntax(SyntaxToken? name, SyntaxToken? colonEqualsToken, SyntaxToken? byValKeyword, ExpressionSyntax? expression)
        : base(SyntaxKind.Argument)
    {
        Name = name;
        ColonEqualsToken = colonEqualsToken;
        ByValKeyword = byValKeyword;
        Expression = expression;
    }

    public SyntaxToken? Name { get; }

    public SyntaxToken? ColonEqualsToken { get; }

    public SyntaxToken? ByValKeyword { get; }

    public ExpressionSyntax? Expression { get; }

    public bool IsNamed => Name is not null;

    public bool IsOmitted => Expression is null && Name is null;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Name, ColonEqualsToken, ByValKeyword, Expression);
}

/// <summary>"(" expression ")" (MS-VBAL 5.6.6). Parentheses also force ByVal in argument position.</summary>
public sealed class ParenthesizedExpressionSyntax : ExpressionSyntax
{
    public ParenthesizedExpressionSyntax(SyntaxToken openParenToken, ExpressionSyntax expression, SyntaxToken closeParenToken)
        : base(SyntaxKind.ParenthesizedExpression)
    {
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
    }

    public SyntaxToken OpenParenToken { get; }

    public ExpressionSyntax Expression { get; }

    public SyntaxToken CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(OpenParenToken, Expression, CloseParenToken);
}

/// <summary>Unary "-", "+", or "Not" (MS-VBAL 5.6.9.3, 5.6.9.7).</summary>
public sealed class UnaryExpressionSyntax : ExpressionSyntax
{
    public UnaryExpressionSyntax(SyntaxToken operatorToken, ExpressionSyntax operand)
        : base(SyntaxKind.UnaryExpression)
    {
        OperatorToken = operatorToken ?? throw new ArgumentNullException(nameof(operatorToken));
        Operand = operand ?? throw new ArgumentNullException(nameof(operand));
    }

    public SyntaxToken OperatorToken { get; }

    public ExpressionSyntax Operand { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(OperatorToken, Operand);
}

/// <summary>A binary operator expression (MS-VBAL 5.6.9). All binary operators are left-associative.</summary>
public sealed class BinaryExpressionSyntax : ExpressionSyntax
{
    public BinaryExpressionSyntax(ExpressionSyntax left, SyntaxToken operatorToken, ExpressionSyntax right)
        : base(SyntaxKind.BinaryExpression)
    {
        Left = left ?? throw new ArgumentNullException(nameof(left));
        OperatorToken = operatorToken ?? throw new ArgumentNullException(nameof(operatorToken));
        Right = right ?? throw new ArgumentNullException(nameof(right));
    }

    public ExpressionSyntax Left { get; }

    public SyntaxToken OperatorToken { get; }

    public ExpressionSyntax Right { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Left, OperatorToken, Right);
}

/// <summary>"New" type (MS-VBAL 5.6.11).</summary>
public sealed class NewExpressionSyntax : ExpressionSyntax
{
    public NewExpressionSyntax(SyntaxToken newKeyword, ExpressionSyntax type)
        : base(SyntaxKind.NewExpression)
    {
        NewKeyword = newKeyword ?? throw new ArgumentNullException(nameof(newKeyword));
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    public SyntaxToken NewKeyword { get; }

    public ExpressionSyntax Type { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(NewKeyword, Type);
}

/// <summary>"TypeOf" expression "Is" type (MS-VBAL 5.6.16.5).</summary>
public sealed class TypeOfExpressionSyntax : ExpressionSyntax
{
    public TypeOfExpressionSyntax(SyntaxToken typeOfKeyword, ExpressionSyntax expression, SyntaxToken isKeyword, ExpressionSyntax type)
        : base(SyntaxKind.TypeOfExpression)
    {
        TypeOfKeyword = typeOfKeyword ?? throw new ArgumentNullException(nameof(typeOfKeyword));
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        IsKeyword = isKeyword ?? throw new ArgumentNullException(nameof(isKeyword));
        Type = type ?? throw new ArgumentNullException(nameof(type));
    }

    public SyntaxToken TypeOfKeyword { get; }

    public ExpressionSyntax Expression { get; }

    public SyntaxToken IsKeyword { get; }

    public ExpressionSyntax Type { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(TypeOfKeyword, Expression, IsKeyword, Type);
}

/// <summary>"AddressOf" procedure (MS-VBAL 5.6.16.8).</summary>
public sealed class AddressOfExpressionSyntax : ExpressionSyntax
{
    public AddressOfExpressionSyntax(SyntaxToken addressOfKeyword, ExpressionSyntax procedure)
        : base(SyntaxKind.AddressOfExpression)
    {
        AddressOfKeyword = addressOfKeyword ?? throw new ArgumentNullException(nameof(addressOfKeyword));
        Procedure = procedure ?? throw new ArgumentNullException(nameof(procedure));
    }

    public SyntaxToken AddressOfKeyword { get; }

    public ExpressionSyntax Procedure { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(AddressOfKeyword, Procedure);
}
