namespace VbaNg.Compiler.Syntax;

/// <summary>A statement the parser could not understand; its tokens are kept so the tree still round-trips.</summary>
public sealed class BadStatementSyntax : StatementSyntax
{
    public BadStatementSyntax(SyntaxTokenList tokens)
        : base(SyntaxKind.BadStatement)
    {
        Tokens = tokens ?? throw new ArgumentNullException(nameof(tokens));
    }

    public SyntaxTokenList Tokens { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Tokens.AsChildren();
}

/// <summary>A label definition first on a line: name ":" or a line number [":"] (MS-VBAL 5.4.1.1).</summary>
public sealed class LabelStatementSyntax : StatementSyntax
{
    public LabelStatementSyntax(SyntaxToken label, SyntaxToken? colonToken)
        : base(SyntaxKind.LabelStatement)
    {
        Label = label ?? throw new ArgumentNullException(nameof(label));
        ColonToken = colonToken;
    }

    public SyntaxToken Label { get; }

    public SyntaxToken? ColonToken { get; }

    public bool IsLineNumber => Label.Kind == SyntaxKind.IntegerLiteralToken;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Label, ColonToken);
}

/// <summary>
/// A procedure call statement (MS-VBAL 5.4.2.1): "Call" expression, or an expression followed by
/// bare arguments. Without Call and without bare arguments, the expression alone is the call
/// (for example "Foo(1)" or "obj.Method").
/// </summary>
public sealed class CallStatementSyntax : StatementSyntax
{
    public CallStatementSyntax(SyntaxToken? callKeyword, ExpressionSyntax expression, ArgumentListSyntax? arguments)
        : base(SyntaxKind.CallStatement)
    {
        CallKeyword = callKeyword;
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        Arguments = arguments;
    }

    public SyntaxToken? CallKeyword { get; }

    public ExpressionSyntax Expression { get; }

    /// <summary>Arguments written without parentheses, or null.</summary>
    public ArgumentListSyntax? Arguments { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(CallKeyword, Expression, Arguments);
}

/// <summary>[Let] target = value (MS-VBAL 5.4.3.8) or Set target = value (5.4.3.9), told apart by <see cref="SyntaxNode.Kind"/>.</summary>
public sealed class AssignmentStatementSyntax : StatementSyntax
{
    public AssignmentStatementSyntax(SyntaxKind kind, SyntaxToken? keyword, ExpressionSyntax target, SyntaxToken equalsToken, ExpressionSyntax value)
        : base(kind)
    {
        Keyword = keyword;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Let or Set, or null for the implicit Let form.</summary>
    public SyntaxToken? Keyword { get; }

    public ExpressionSyntax Target { get; }

    public SyntaxToken EqualsToken { get; }

    public ExpressionSyntax Value { get; }

    public bool IsSet => Kind == SyntaxKind.SetAssignmentStatement;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, Target, EqualsToken, Value);
}

/// <summary>ReDim [Preserve] declarators (MS-VBAL 5.4.3.3).</summary>
public sealed class ReDimStatementSyntax : StatementSyntax
{
    public ReDimStatementSyntax(SyntaxToken reDimKeyword, SyntaxToken? preserveKeyword, SeparatedSyntaxList<ReDimDeclaratorSyntax> declarators)
        : base(SyntaxKind.ReDimStatement)
    {
        ReDimKeyword = reDimKeyword ?? throw new ArgumentNullException(nameof(reDimKeyword));
        PreserveKeyword = preserveKeyword;
        Declarators = declarators ?? throw new ArgumentNullException(nameof(declarators));
    }

    public SyntaxToken ReDimKeyword { get; }

    public SyntaxToken? PreserveKeyword { get; }

    public SeparatedSyntaxList<ReDimDeclaratorSyntax> Declarators { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(ReDimKeyword), One(PreserveKeyword), Declarators.AsChildren());
}

/// <summary>target (bounds) [As type] in a ReDim statement.</summary>
public sealed class ReDimDeclaratorSyntax : SyntaxNode
{
    public ReDimDeclaratorSyntax(ExpressionSyntax target, ArrayBoundsSyntax bounds, AsClauseSyntax? asClause)
        : base(SyntaxKind.ReDimDeclarator)
    {
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Bounds = bounds ?? throw new ArgumentNullException(nameof(bounds));
        AsClause = asClause;
    }

    public ExpressionSyntax Target { get; }

    public ArrayBoundsSyntax Bounds { get; }

    public AsClauseSyntax? AsClause { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Target, Bounds, AsClause);
}

/// <summary>Erase arrays (MS-VBAL 5.4.3.4).</summary>
public sealed class EraseStatementSyntax : StatementSyntax
{
    public EraseStatementSyntax(SyntaxToken eraseKeyword, SeparatedSyntaxList<ExpressionSyntax> targets)
        : base(SyntaxKind.EraseStatement)
    {
        EraseKeyword = eraseKeyword ?? throw new ArgumentNullException(nameof(eraseKeyword));
        Targets = targets ?? throw new ArgumentNullException(nameof(targets));
    }

    public SyntaxToken EraseKeyword { get; }

    public SeparatedSyntaxList<ExpressionSyntax> Targets { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(EraseKeyword), Targets.AsChildren());
}

/// <summary>Mid | MidB | Mid$ | MidB$ "(" target "," start ["," length] ")" "=" value (MS-VBAL 5.4.3.5).</summary>
public sealed class MidStatementSyntax : StatementSyntax
{
    public MidStatementSyntax(
        SyntaxToken midKeyword,
        SyntaxToken openParenToken,
        ExpressionSyntax target,
        SyntaxToken firstCommaToken,
        ExpressionSyntax start,
        SyntaxToken? secondCommaToken,
        ExpressionSyntax? length,
        SyntaxToken closeParenToken,
        SyntaxToken equalsToken,
        ExpressionSyntax value)
        : base(SyntaxKind.MidStatement)
    {
        MidKeyword = midKeyword ?? throw new ArgumentNullException(nameof(midKeyword));
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        Target = target ?? throw new ArgumentNullException(nameof(target));
        FirstCommaToken = firstCommaToken ?? throw new ArgumentNullException(nameof(firstCommaToken));
        StartIndex = start ?? throw new ArgumentNullException(nameof(start));
        SecondCommaToken = secondCommaToken;
        Length = length;
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public SyntaxToken MidKeyword { get; }

    public SyntaxToken OpenParenToken { get; }

    public ExpressionSyntax Target { get; }

    public SyntaxToken FirstCommaToken { get; }

    public ExpressionSyntax StartIndex { get; }

    public SyntaxToken? SecondCommaToken { get; }

    public ExpressionSyntax? Length { get; }

    public SyntaxToken CloseParenToken { get; }

    public SyntaxToken EqualsToken { get; }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(MidKeyword, OpenParenToken, Target, FirstCommaToken, StartIndex, SecondCommaToken, Length, CloseParenToken, EqualsToken, Value);
}

/// <summary>LSet target = value or RSet target = value (MS-VBAL 5.4.3.6, 5.4.3.7).</summary>
public sealed class LSetRSetStatementSyntax : StatementSyntax
{
    public LSetRSetStatementSyntax(SyntaxKind kind, SyntaxToken keyword, ExpressionSyntax target, SyntaxToken equalsToken, ExpressionSyntax value)
        : base(kind)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        Target = target ?? throw new ArgumentNullException(nameof(target));
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public SyntaxToken Keyword { get; }

    public ExpressionSyntax Target { get; }

    public SyntaxToken EqualsToken { get; }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, Target, EqualsToken, Value);
}

/// <summary>A multi-line If block (MS-VBAL 5.4.2.9).</summary>
public sealed class IfBlockSyntax : StatementSyntax
{
    public IfBlockSyntax(
        SyntaxToken ifKeyword,
        ExpressionSyntax condition,
        SyntaxToken thenKeyword,
        SyntaxList<StatementSyntax> statements,
        SyntaxList<ElseIfBlockSyntax> elseIfBlocks,
        ElseBlockSyntax? elseBlock,
        EndBlockStatementSyntax endIf)
        : base(SyntaxKind.IfBlock)
    {
        IfKeyword = ifKeyword ?? throw new ArgumentNullException(nameof(ifKeyword));
        Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        ThenKeyword = thenKeyword ?? throw new ArgumentNullException(nameof(thenKeyword));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        ElseIfBlocks = elseIfBlocks ?? throw new ArgumentNullException(nameof(elseIfBlocks));
        ElseBlock = elseBlock;
        EndIf = endIf ?? throw new ArgumentNullException(nameof(endIf));
    }

    public SyntaxToken IfKeyword { get; }

    public ExpressionSyntax Condition { get; }

    public SyntaxToken ThenKeyword { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public SyntaxList<ElseIfBlockSyntax> ElseIfBlocks { get; }

    public ElseBlockSyntax? ElseBlock { get; }

    public EndBlockStatementSyntax EndIf { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(IfKeyword), One(Condition), One(ThenKeyword), Statements.AsChildren(), ElseIfBlocks.AsChildren(), One(ElseBlock), One(EndIf));
}

/// <summary>ElseIf condition Then statements.</summary>
public sealed class ElseIfBlockSyntax : SyntaxNode
{
    public ElseIfBlockSyntax(SyntaxToken elseIfKeyword, ExpressionSyntax condition, SyntaxToken thenKeyword, SyntaxList<StatementSyntax> statements)
        : base(SyntaxKind.ElseIfBlock)
    {
        ElseIfKeyword = elseIfKeyword ?? throw new ArgumentNullException(nameof(elseIfKeyword));
        Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        ThenKeyword = thenKeyword ?? throw new ArgumentNullException(nameof(thenKeyword));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
    }

    public SyntaxToken ElseIfKeyword { get; }

    public ExpressionSyntax Condition { get; }

    public SyntaxToken ThenKeyword { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(ElseIfKeyword), One(Condition), One(ThenKeyword), Statements.AsChildren());
}

/// <summary>Else statements.</summary>
public sealed class ElseBlockSyntax : SyntaxNode
{
    public ElseBlockSyntax(SyntaxToken elseKeyword, SyntaxList<StatementSyntax> statements)
        : base(SyntaxKind.ElseBlock)
    {
        ElseKeyword = elseKeyword ?? throw new ArgumentNullException(nameof(elseKeyword));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
    }

    public SyntaxToken ElseKeyword { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(ElseKeyword), Statements.AsChildren());
}

/// <summary>If condition Then statements [Else statements] on one logical line (MS-VBAL 5.4.2.8).</summary>
public sealed class SingleLineIfStatementSyntax : StatementSyntax
{
    public SingleLineIfStatementSyntax(SyntaxToken ifKeyword, ExpressionSyntax condition, SyntaxToken thenKeyword, SyntaxList<StatementSyntax> statements, SingleLineElseClauseSyntax? elseClause)
        : base(SyntaxKind.SingleLineIfStatement)
    {
        IfKeyword = ifKeyword ?? throw new ArgumentNullException(nameof(ifKeyword));
        Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        ThenKeyword = thenKeyword ?? throw new ArgumentNullException(nameof(thenKeyword));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        ElseClause = elseClause;
    }

    public SyntaxToken IfKeyword { get; }

    public ExpressionSyntax Condition { get; }

    public SyntaxToken ThenKeyword { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public SingleLineElseClauseSyntax? ElseClause { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(IfKeyword), One(Condition), One(ThenKeyword), Statements.AsChildren(), One(ElseClause));
}

/// <summary>Else statements on the same line as a single-line If.</summary>
public sealed class SingleLineElseClauseSyntax : SyntaxNode
{
    public SingleLineElseClauseSyntax(SyntaxToken elseKeyword, SyntaxList<StatementSyntax> statements)
        : base(SyntaxKind.SingleLineElseClause)
    {
        ElseKeyword = elseKeyword ?? throw new ArgumentNullException(nameof(elseKeyword));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
    }

    public SyntaxToken ElseKeyword { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(ElseKeyword), Statements.AsChildren());
}

/// <summary>Select Case expression ... End Select (MS-VBAL 5.4.2.10).</summary>
public sealed class SelectCaseBlockSyntax : StatementSyntax
{
    public SelectCaseBlockSyntax(SyntaxToken selectKeyword, SyntaxToken caseKeyword, ExpressionSyntax expression, SyntaxList<CaseBlockSyntax> caseBlocks, EndBlockStatementSyntax endSelect)
        : base(SyntaxKind.SelectCaseBlock)
    {
        SelectKeyword = selectKeyword ?? throw new ArgumentNullException(nameof(selectKeyword));
        CaseKeyword = caseKeyword ?? throw new ArgumentNullException(nameof(caseKeyword));
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        CaseBlocks = caseBlocks ?? throw new ArgumentNullException(nameof(caseBlocks));
        EndSelect = endSelect ?? throw new ArgumentNullException(nameof(endSelect));
    }

    public SyntaxToken SelectKeyword { get; }

    public SyntaxToken CaseKeyword { get; }

    public ExpressionSyntax Expression { get; }

    public SyntaxList<CaseBlockSyntax> CaseBlocks { get; }

    public EndBlockStatementSyntax EndSelect { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(SelectKeyword), One(CaseKeyword), One(Expression), CaseBlocks.AsChildren(), One(EndSelect));
}

/// <summary>Case clauses statements, or Case Else statements.</summary>
public sealed class CaseBlockSyntax : SyntaxNode
{
    public CaseBlockSyntax(SyntaxToken caseKeyword, SyntaxToken? elseKeyword, SeparatedSyntaxList<CaseClauseSyntax> clauses, SyntaxList<StatementSyntax> statements)
        : base(SyntaxKind.CaseBlock)
    {
        CaseKeyword = caseKeyword ?? throw new ArgumentNullException(nameof(caseKeyword));
        ElseKeyword = elseKeyword;
        Clauses = clauses ?? throw new ArgumentNullException(nameof(clauses));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
    }

    public SyntaxToken CaseKeyword { get; }

    public SyntaxToken? ElseKeyword { get; }

    public SeparatedSyntaxList<CaseClauseSyntax> Clauses { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public bool IsCaseElse => ElseKeyword is not null;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(CaseKeyword), One(ElseKeyword), Clauses.AsChildren(), Statements.AsChildren());
}

/// <summary>Base of the three Case clause forms (MS-VBAL 5.4.2.10 range-clause).</summary>
public abstract class CaseClauseSyntax : SyntaxNode
{
    protected CaseClauseSyntax(SyntaxKind kind)
        : base(kind)
    {
    }
}

/// <summary>Case expression.</summary>
public sealed class ValueCaseClauseSyntax : CaseClauseSyntax
{
    public ValueCaseClauseSyntax(ExpressionSyntax value)
        : base(SyntaxKind.ValueCaseClause)
    {
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Value);
}

/// <summary>Case from To to.</summary>
public sealed class RangeCaseClauseSyntax : CaseClauseSyntax
{
    public RangeCaseClauseSyntax(ExpressionSyntax from, SyntaxToken toKeyword, ExpressionSyntax to)
        : base(SyntaxKind.RangeCaseClause)
    {
        From = from ?? throw new ArgumentNullException(nameof(from));
        ToKeyword = toKeyword ?? throw new ArgumentNullException(nameof(toKeyword));
        To = to ?? throw new ArgumentNullException(nameof(to));
    }

    public ExpressionSyntax From { get; }

    public SyntaxToken ToKeyword { get; }

    public ExpressionSyntax To { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(From, ToKeyword, To);
}

/// <summary>Case Is operator expression.</summary>
public sealed class IsCaseClauseSyntax : CaseClauseSyntax
{
    public IsCaseClauseSyntax(SyntaxToken isKeyword, SyntaxToken operatorToken, ExpressionSyntax value)
        : base(SyntaxKind.IsCaseClause)
    {
        IsKeyword = isKeyword ?? throw new ArgumentNullException(nameof(isKeyword));
        OperatorToken = operatorToken ?? throw new ArgumentNullException(nameof(operatorToken));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public SyntaxToken IsKeyword { get; }

    public SyntaxToken OperatorToken { get; }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(IsKeyword, OperatorToken, Value);
}

/// <summary>For variable = from To to [Step step] ... Next (MS-VBAL 5.4.2.3). Next is null when a later "Next a, b" closes this loop.</summary>
public sealed class ForBlockSyntax : StatementSyntax
{
    public ForBlockSyntax(
        SyntaxToken forKeyword,
        ExpressionSyntax variable,
        SyntaxToken equalsToken,
        ExpressionSyntax from,
        SyntaxToken toKeyword,
        ExpressionSyntax to,
        StepClauseSyntax? stepClause,
        SyntaxList<StatementSyntax> statements,
        NextStatementSyntax? next)
        : base(SyntaxKind.ForBlock)
    {
        ForKeyword = forKeyword ?? throw new ArgumentNullException(nameof(forKeyword));
        Variable = variable ?? throw new ArgumentNullException(nameof(variable));
        EqualsToken = equalsToken ?? throw new ArgumentNullException(nameof(equalsToken));
        From = from ?? throw new ArgumentNullException(nameof(from));
        ToKeyword = toKeyword ?? throw new ArgumentNullException(nameof(toKeyword));
        To = to ?? throw new ArgumentNullException(nameof(to));
        StepClause = stepClause;
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        Next = next;
    }

    public SyntaxToken ForKeyword { get; }

    public ExpressionSyntax Variable { get; }

    public SyntaxToken EqualsToken { get; }

    public ExpressionSyntax From { get; }

    public SyntaxToken ToKeyword { get; }

    public ExpressionSyntax To { get; }

    public StepClauseSyntax? StepClause { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public NextStatementSyntax? Next { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(ForKeyword), One(Variable), One(EqualsToken), One(From), One(ToKeyword), One(To), One(StepClause), Statements.AsChildren(), One(Next));
}

/// <summary>Step expression.</summary>
public sealed class StepClauseSyntax : SyntaxNode
{
    public StepClauseSyntax(SyntaxToken stepKeyword, ExpressionSyntax value)
        : base(SyntaxKind.StepClause)
    {
        StepKeyword = stepKeyword ?? throw new ArgumentNullException(nameof(stepKeyword));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    public SyntaxToken StepKeyword { get; }

    public ExpressionSyntax Value { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(StepKeyword, Value);
}

/// <summary>For Each variable In collection ... Next (MS-VBAL 5.4.2.4).</summary>
public sealed class ForEachBlockSyntax : StatementSyntax
{
    public ForEachBlockSyntax(
        SyntaxToken forKeyword,
        SyntaxToken eachKeyword,
        ExpressionSyntax variable,
        SyntaxToken inKeyword,
        ExpressionSyntax collection,
        SyntaxList<StatementSyntax> statements,
        NextStatementSyntax? next)
        : base(SyntaxKind.ForEachBlock)
    {
        ForKeyword = forKeyword ?? throw new ArgumentNullException(nameof(forKeyword));
        EachKeyword = eachKeyword ?? throw new ArgumentNullException(nameof(eachKeyword));
        Variable = variable ?? throw new ArgumentNullException(nameof(variable));
        InKeyword = inKeyword ?? throw new ArgumentNullException(nameof(inKeyword));
        Collection = collection ?? throw new ArgumentNullException(nameof(collection));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        Next = next;
    }

    public SyntaxToken ForKeyword { get; }

    public SyntaxToken EachKeyword { get; }

    public ExpressionSyntax Variable { get; }

    public SyntaxToken InKeyword { get; }

    public ExpressionSyntax Collection { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public NextStatementSyntax? Next { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(ForKeyword), One(EachKeyword), One(Variable), One(InKeyword), One(Collection), Statements.AsChildren(), One(Next));
}

/// <summary>Next [variable, ...]; several variables close several loops.</summary>
public sealed class NextStatementSyntax : StatementSyntax
{
    public NextStatementSyntax(SyntaxToken nextKeyword, SeparatedSyntaxList<ExpressionSyntax> variables)
        : base(SyntaxKind.NextStatement)
    {
        NextKeyword = nextKeyword ?? throw new ArgumentNullException(nameof(nextKeyword));
        Variables = variables ?? throw new ArgumentNullException(nameof(variables));
    }

    public SyntaxToken NextKeyword { get; }

    public SeparatedSyntaxList<ExpressionSyntax> Variables { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(NextKeyword), Variables.AsChildren());
}

/// <summary>Do [While | Until cond] ... Loop [While | Until cond] (MS-VBAL 5.4.2.2).</summary>
public sealed class DoLoopBlockSyntax : StatementSyntax
{
    public DoLoopBlockSyntax(SyntaxToken doKeyword, WhileOrUntilClauseSyntax? topCondition, SyntaxList<StatementSyntax> statements, SyntaxToken loopKeyword, WhileOrUntilClauseSyntax? bottomCondition)
        : base(SyntaxKind.DoLoopBlock)
    {
        DoKeyword = doKeyword ?? throw new ArgumentNullException(nameof(doKeyword));
        TopCondition = topCondition;
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        LoopKeyword = loopKeyword ?? throw new ArgumentNullException(nameof(loopKeyword));
        BottomCondition = bottomCondition;
    }

    public SyntaxToken DoKeyword { get; }

    public WhileOrUntilClauseSyntax? TopCondition { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public SyntaxToken LoopKeyword { get; }

    public WhileOrUntilClauseSyntax? BottomCondition { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(DoKeyword), One(TopCondition), Statements.AsChildren(), One(LoopKeyword), One(BottomCondition));
}

/// <summary>While condition or Until condition.</summary>
public sealed class WhileOrUntilClauseSyntax : SyntaxNode
{
    public WhileOrUntilClauseSyntax(SyntaxToken keyword, ExpressionSyntax condition)
        : base(SyntaxKind.WhileOrUntilClause)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        Condition = condition ?? throw new ArgumentNullException(nameof(condition));
    }

    public SyntaxToken Keyword { get; }

    public ExpressionSyntax Condition { get; }

    public bool IsUntil => Keyword.Kind == SyntaxKind.UntilKeyword;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, Condition);
}

/// <summary>While condition ... Wend (MS-VBAL 5.4.2.6).</summary>
public sealed class WhileBlockSyntax : StatementSyntax
{
    public WhileBlockSyntax(SyntaxToken whileKeyword, ExpressionSyntax condition, SyntaxList<StatementSyntax> statements, SyntaxToken wendKeyword)
        : base(SyntaxKind.WhileBlock)
    {
        WhileKeyword = whileKeyword ?? throw new ArgumentNullException(nameof(whileKeyword));
        Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        WendKeyword = wendKeyword ?? throw new ArgumentNullException(nameof(wendKeyword));
    }

    public SyntaxToken WhileKeyword { get; }

    public ExpressionSyntax Condition { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public SyntaxToken WendKeyword { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(WhileKeyword), One(Condition), Statements.AsChildren(), One(WendKeyword));
}

/// <summary>With expression ... End With (MS-VBAL 5.4.2.7).</summary>
public sealed class WithBlockSyntax : StatementSyntax
{
    public WithBlockSyntax(SyntaxToken withKeyword, ExpressionSyntax expression, SyntaxList<StatementSyntax> statements, EndBlockStatementSyntax endWith)
        : base(SyntaxKind.WithBlock)
    {
        WithKeyword = withKeyword ?? throw new ArgumentNullException(nameof(withKeyword));
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        Statements = statements ?? throw new ArgumentNullException(nameof(statements));
        EndWith = endWith ?? throw new ArgumentNullException(nameof(endWith));
    }

    public SyntaxToken WithKeyword { get; }

    public ExpressionSyntax Expression { get; }

    public SyntaxList<StatementSyntax> Statements { get; }

    public EndBlockStatementSyntax EndWith { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(WithKeyword), One(Expression), Statements.AsChildren(), One(EndWith));
}

/// <summary>GoTo label or GoSub label (MS-VBAL 5.4.2.11, 5.4.2.12), told apart by <see cref="SyntaxNode.Kind"/>.</summary>
public sealed class GoToStatementSyntax : StatementSyntax
{
    public GoToStatementSyntax(SyntaxKind kind, SyntaxToken keyword, SyntaxToken label)
        : base(kind)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        Label = label ?? throw new ArgumentNullException(nameof(label));
    }

    public SyntaxToken Keyword { get; }

    /// <summary>An identifier or a line number.</summary>
    public SyntaxToken Label { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, Label);
}

/// <summary>Return, End, Stop, Reset: statements made of one token, told apart by <see cref="SyntaxNode.Kind"/>.</summary>
public sealed class KeywordStatementSyntax : StatementSyntax
{
    public KeywordStatementSyntax(SyntaxKind kind, SyntaxToken keyword)
        : base(kind)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
    }

    public SyntaxToken Keyword { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword);
}

/// <summary>On Error GoTo label | 0 | -1, or On Error Resume Next (MS-VBAL 5.4.4.1).</summary>
public sealed class OnErrorStatementSyntax : StatementSyntax
{
    public OnErrorStatementSyntax(SyntaxToken onKeyword, SyntaxToken errorToken, SyntaxToken actionKeyword, SyntaxToken? nextKeyword, ExpressionSyntax? target)
        : base(SyntaxKind.OnErrorStatement)
    {
        OnKeyword = onKeyword ?? throw new ArgumentNullException(nameof(onKeyword));
        ErrorToken = errorToken ?? throw new ArgumentNullException(nameof(errorToken));
        ActionKeyword = actionKeyword ?? throw new ArgumentNullException(nameof(actionKeyword));
        NextKeyword = nextKeyword;
        Target = target;
    }

    public SyntaxToken OnKeyword { get; }

    public SyntaxToken ErrorToken { get; }

    /// <summary>GoTo or Resume.</summary>
    public SyntaxToken ActionKeyword { get; }

    public SyntaxToken? NextKeyword { get; }

    /// <summary>The label name, 0, or -1 after GoTo.</summary>
    public ExpressionSyntax? Target { get; }

    public bool IsResumeNext => NextKeyword is not null;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(OnKeyword, ErrorToken, ActionKeyword, NextKeyword, Target);
}

/// <summary>On expression GoTo | GoSub labels (MS-VBAL 5.4.2.13 computed jump).</summary>
public sealed class OnGoToStatementSyntax : StatementSyntax
{
    public OnGoToStatementSyntax(SyntaxToken onKeyword, ExpressionSyntax expression, SyntaxToken jumpKeyword, SeparatedSyntaxList<ExpressionSyntax> labels)
        : base(SyntaxKind.OnGoToStatement)
    {
        OnKeyword = onKeyword ?? throw new ArgumentNullException(nameof(onKeyword));
        Expression = expression ?? throw new ArgumentNullException(nameof(expression));
        JumpKeyword = jumpKeyword ?? throw new ArgumentNullException(nameof(jumpKeyword));
        Labels = labels ?? throw new ArgumentNullException(nameof(labels));
    }

    public SyntaxToken OnKeyword { get; }

    public ExpressionSyntax Expression { get; }

    public SyntaxToken JumpKeyword { get; }

    public SeparatedSyntaxList<ExpressionSyntax> Labels { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(OnKeyword), One(Expression), One(JumpKeyword), Labels.AsChildren());
}

/// <summary>Resume, Resume Next, or Resume label (MS-VBAL 5.4.4.2).</summary>
public sealed class ResumeStatementSyntax : StatementSyntax
{
    public ResumeStatementSyntax(SyntaxToken resumeKeyword, SyntaxToken? nextKeyword, ExpressionSyntax? label)
        : base(SyntaxKind.ResumeStatement)
    {
        ResumeKeyword = resumeKeyword ?? throw new ArgumentNullException(nameof(resumeKeyword));
        NextKeyword = nextKeyword;
        Label = label;
    }

    public SyntaxToken ResumeKeyword { get; }

    public SyntaxToken? NextKeyword { get; }

    public ExpressionSyntax? Label { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(ResumeKeyword, NextKeyword, Label);
}

/// <summary>Error number (MS-VBAL 5.4.4.3).</summary>
public sealed class ErrorStatementSyntax : StatementSyntax
{
    public ErrorStatementSyntax(SyntaxToken errorToken, ExpressionSyntax number)
        : base(SyntaxKind.ErrorStatement)
    {
        ErrorToken = errorToken ?? throw new ArgumentNullException(nameof(errorToken));
        Number = number ?? throw new ArgumentNullException(nameof(number));
    }

    public SyntaxToken ErrorToken { get; }

    public ExpressionSyntax Number { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(ErrorToken, Number);
}

/// <summary>Exit Sub | Function | Property | Do | For (MS-VBAL 5.4.2.14 through 5.4.2.18).</summary>
public sealed class ExitStatementSyntax : StatementSyntax
{
    public ExitStatementSyntax(SyntaxToken exitKeyword, SyntaxToken blockKeyword)
        : base(SyntaxKind.ExitStatement)
    {
        ExitKeyword = exitKeyword ?? throw new ArgumentNullException(nameof(exitKeyword));
        BlockKeyword = blockKeyword ?? throw new ArgumentNullException(nameof(blockKeyword));
    }

    public SyntaxToken ExitKeyword { get; }

    public SyntaxToken BlockKeyword { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(ExitKeyword, BlockKeyword);
}

/// <summary>RaiseEvent name [(arguments)] (MS-VBAL 5.4.2.20).</summary>
public sealed class RaiseEventStatementSyntax : StatementSyntax
{
    public RaiseEventStatementSyntax(SyntaxToken raiseEventKeyword, SyntaxToken name, ArgumentListSyntax? arguments)
        : base(SyntaxKind.RaiseEventStatement)
    {
        RaiseEventKeyword = raiseEventKeyword ?? throw new ArgumentNullException(nameof(raiseEventKeyword));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Arguments = arguments;
    }

    public SyntaxToken RaiseEventKeyword { get; }

    public SyntaxToken Name { get; }

    public ArgumentListSyntax? Arguments { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(RaiseEventKeyword, Name, Arguments);
}

/// <summary>Name oldpath As newpath (MS-VBAL 5.4.5.10).</summary>
public sealed class NameStatementSyntax : StatementSyntax
{
    public NameStatementSyntax(SyntaxToken nameToken, ExpressionSyntax oldName, SyntaxToken asKeyword, ExpressionSyntax newName)
        : base(SyntaxKind.NameStatement)
    {
        NameToken = nameToken ?? throw new ArgumentNullException(nameof(nameToken));
        OldName = oldName ?? throw new ArgumentNullException(nameof(oldName));
        AsKeyword = asKeyword ?? throw new ArgumentNullException(nameof(asKeyword));
        NewName = newName ?? throw new ArgumentNullException(nameof(newName));
    }

    public SyntaxToken NameToken { get; }

    public ExpressionSyntax OldName { get; }

    public SyntaxToken AsKeyword { get; }

    public ExpressionSyntax NewName { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(NameToken, OldName, AsKeyword, NewName);
}

/// <summary>Open path [For mode] [Access access] [lock] As [#]number [Len = length] (MS-VBAL 5.4.5.1).</summary>
public sealed class OpenStatementSyntax : StatementSyntax
{
    public OpenStatementSyntax(
        SyntaxToken openKeyword,
        ExpressionSyntax pathName,
        SyntaxToken? forKeyword,
        SyntaxToken? mode,
        SyntaxToken? accessKeyword,
        SyntaxTokenList accessModes,
        SyntaxTokenList lockModes,
        SyntaxToken asKeyword,
        FileNumberSyntax fileNumber,
        SyntaxToken? lenKeyword,
        SyntaxToken? equalsToken,
        ExpressionSyntax? recordLength)
        : base(SyntaxKind.OpenStatement)
    {
        OpenKeyword = openKeyword ?? throw new ArgumentNullException(nameof(openKeyword));
        PathName = pathName ?? throw new ArgumentNullException(nameof(pathName));
        ForKeyword = forKeyword;
        Mode = mode;
        AccessKeyword = accessKeyword;
        AccessModes = accessModes ?? throw new ArgumentNullException(nameof(accessModes));
        LockModes = lockModes ?? throw new ArgumentNullException(nameof(lockModes));
        AsKeyword = asKeyword ?? throw new ArgumentNullException(nameof(asKeyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        LenKeyword = lenKeyword;
        EqualsToken = equalsToken;
        RecordLength = recordLength;
    }

    public SyntaxToken OpenKeyword { get; }

    public ExpressionSyntax PathName { get; }

    public SyntaxToken? ForKeyword { get; }

    /// <summary>Append, Binary, Input, Output, or Random.</summary>
    public SyntaxToken? Mode { get; }

    public SyntaxToken? AccessKeyword { get; }

    /// <summary>Read, Write, or Read Write.</summary>
    public SyntaxTokenList AccessModes { get; }

    /// <summary>Shared, Lock Read, Lock Write, or Lock Read Write.</summary>
    public SyntaxTokenList LockModes { get; }

    public SyntaxToken AsKeyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken? LenKeyword { get; }

    public SyntaxToken? EqualsToken { get; }

    public ExpressionSyntax? RecordLength { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(
            One(OpenKeyword),
            One(PathName),
            One(ForKeyword),
            One(Mode),
            One(AccessKeyword),
            AccessModes.AsChildren(),
            LockModes.AsChildren(),
            One(AsKeyword),
            One(FileNumber),
            One(LenKeyword),
            One(EqualsToken),
            One(RecordLength));
}

/// <summary>[#] expression naming an open file (MS-VBAL 5.4.5 file-number).</summary>
public sealed class FileNumberSyntax : SyntaxNode
{
    public FileNumberSyntax(SyntaxToken? hashToken, ExpressionSyntax number)
        : base(SyntaxKind.FileNumber)
    {
        HashToken = hashToken;
        Number = number ?? throw new ArgumentNullException(nameof(number));
    }

    public SyntaxToken? HashToken { get; }

    public ExpressionSyntax Number { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(HashToken, Number);
}

/// <summary>Close [file numbers] (MS-VBAL 5.4.5.1.1).</summary>
public sealed class CloseStatementSyntax : StatementSyntax
{
    public CloseStatementSyntax(SyntaxToken closeKeyword, SeparatedSyntaxList<FileNumberSyntax> fileNumbers)
        : base(SyntaxKind.CloseStatement)
    {
        CloseKeyword = closeKeyword ?? throw new ArgumentNullException(nameof(closeKeyword));
        FileNumbers = fileNumbers ?? throw new ArgumentNullException(nameof(fileNumbers));
    }

    public SyntaxToken CloseKeyword { get; }

    public SeparatedSyntaxList<FileNumberSyntax> FileNumbers { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(One(CloseKeyword), FileNumbers.AsChildren());
}

/// <summary>Seek file, position (MS-VBAL 5.4.5.2).</summary>
public sealed class SeekStatementSyntax : StatementSyntax
{
    public SeekStatementSyntax(SyntaxToken seekKeyword, FileNumberSyntax fileNumber, SyntaxToken commaToken, ExpressionSyntax position)
        : base(SyntaxKind.SeekStatement)
    {
        SeekKeyword = seekKeyword ?? throw new ArgumentNullException(nameof(seekKeyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken ?? throw new ArgumentNullException(nameof(commaToken));
        Position = position ?? throw new ArgumentNullException(nameof(position));
    }

    public SyntaxToken SeekKeyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken CommaToken { get; }

    public ExpressionSyntax Position { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(SeekKeyword, FileNumber, CommaToken, Position);
}

/// <summary>Lock | Unlock file [, [start] [To end]] (MS-VBAL 5.4.5.3).</summary>
public sealed class LockStatementSyntax : StatementSyntax
{
    public LockStatementSyntax(SyntaxToken keyword, FileNumberSyntax fileNumber, SyntaxToken? commaToken, RecordRangeSyntax? range)
        : base(SyntaxKind.LockStatement)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken;
        Range = range;
    }

    public SyntaxToken Keyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken? CommaToken { get; }

    public RecordRangeSyntax? Range { get; }

    public bool IsUnlock => Keyword.Kind == SyntaxKind.UnlockKeyword;

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Keyword, FileNumber, CommaToken, Range);
}

/// <summary>[start] [To end] record range of a Lock statement.</summary>
public sealed class RecordRangeSyntax : SyntaxNode
{
    public RecordRangeSyntax(ExpressionSyntax? fromRecord, SyntaxToken? toKeyword, ExpressionSyntax? toRecord)
        : base(SyntaxKind.RecordRange)
    {
        FromRecord = fromRecord;
        ToKeyword = toKeyword;
        ToRecord = toRecord;
    }

    public ExpressionSyntax? FromRecord { get; }

    public SyntaxToken? ToKeyword { get; }

    public ExpressionSyntax? ToRecord { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(FromRecord, ToKeyword, ToRecord);
}

/// <summary>Line Input #file, variable (MS-VBAL 5.4.5.4).</summary>
public sealed class LineInputStatementSyntax : StatementSyntax
{
    public LineInputStatementSyntax(SyntaxToken lineToken, SyntaxToken inputKeyword, FileNumberSyntax fileNumber, SyntaxToken commaToken, ExpressionSyntax variable)
        : base(SyntaxKind.LineInputStatement)
    {
        LineToken = lineToken ?? throw new ArgumentNullException(nameof(lineToken));
        InputKeyword = inputKeyword ?? throw new ArgumentNullException(nameof(inputKeyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken ?? throw new ArgumentNullException(nameof(commaToken));
        Variable = variable ?? throw new ArgumentNullException(nameof(variable));
    }

    public SyntaxToken LineToken { get; }

    public SyntaxToken InputKeyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken CommaToken { get; }

    public ExpressionSyntax Variable { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(LineToken, InputKeyword, FileNumber, CommaToken, Variable);
}

/// <summary>Width #file, width (MS-VBAL 5.4.5.5).</summary>
public sealed class WidthStatementSyntax : StatementSyntax
{
    public WidthStatementSyntax(SyntaxToken widthToken, FileNumberSyntax fileNumber, SyntaxToken commaToken, ExpressionSyntax width)
        : base(SyntaxKind.WidthStatement)
    {
        WidthToken = widthToken ?? throw new ArgumentNullException(nameof(widthToken));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken ?? throw new ArgumentNullException(nameof(commaToken));
        Width = width ?? throw new ArgumentNullException(nameof(width));
    }

    public SyntaxToken WidthToken { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken CommaToken { get; }

    public ExpressionSyntax Width { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(WidthToken, FileNumber, CommaToken, Width);
}

/// <summary>Print #file, items or Write #file, items (MS-VBAL 5.4.5.6, 5.4.5.7), told apart by <see cref="SyntaxNode.Kind"/>.</summary>
public sealed class FileOutputStatementSyntax : StatementSyntax
{
    public FileOutputStatementSyntax(SyntaxKind kind, SyntaxToken keyword, FileNumberSyntax fileNumber, SyntaxToken commaToken, SyntaxList<OutputItemSyntax> items)
        : base(kind)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken ?? throw new ArgumentNullException(nameof(commaToken));
        Items = items ?? throw new ArgumentNullException(nameof(items));
    }

    public SyntaxToken Keyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken CommaToken { get; }

    public SyntaxList<OutputItemSyntax> Items { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(Keyword), One(FileNumber), One(CommaToken), Items.AsChildren());
}

/// <summary>Debug.Print items, using the Print statement's output list (MS-VBAL 5.4.5.6 output-list).</summary>
public sealed class DebugPrintStatementSyntax : StatementSyntax
{
    public DebugPrintStatementSyntax(SyntaxToken debugKeyword, SyntaxToken dotToken, SyntaxToken printKeyword, SyntaxList<OutputItemSyntax> items)
        : base(SyntaxKind.DebugPrintStatement)
    {
        DebugKeyword = debugKeyword ?? throw new ArgumentNullException(nameof(debugKeyword));
        DotToken = dotToken ?? throw new ArgumentNullException(nameof(dotToken));
        PrintKeyword = printKeyword ?? throw new ArgumentNullException(nameof(printKeyword));
        Items = items ?? throw new ArgumentNullException(nameof(items));
    }

    public SyntaxToken DebugKeyword { get; }

    public SyntaxToken DotToken { get; }

    public SyntaxToken PrintKeyword { get; }

    public SyntaxList<OutputItemSyntax> Items { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(DebugKeyword), One(DotToken), One(PrintKeyword), Items.AsChildren());
}

/// <summary>One output item: an expression, Spc(n), or Tab[(n)], each optionally followed by ";" or "," (MS-VBAL 5.4.5.6 output-item).</summary>
public sealed class OutputItemSyntax : SyntaxNode
{
    public OutputItemSyntax(SyntaxNode? clause, SyntaxToken? separator)
        : base(SyntaxKind.OutputItem)
    {
        Clause = clause;
        Separator = separator;
    }

    /// <summary>An <see cref="ExpressionSyntax"/>, <see cref="SpcClauseSyntax"/>, or <see cref="TabClauseSyntax"/>; null for a bare separator.</summary>
    public SyntaxNode? Clause { get; }

    public SyntaxToken? Separator { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(Clause, Separator);
}

/// <summary>Spc(count).</summary>
public sealed class SpcClauseSyntax : SyntaxNode
{
    public SpcClauseSyntax(SyntaxToken spcKeyword, SyntaxToken openParenToken, ExpressionSyntax count, SyntaxToken closeParenToken)
        : base(SyntaxKind.SpcClause)
    {
        SpcKeyword = spcKeyword ?? throw new ArgumentNullException(nameof(spcKeyword));
        OpenParenToken = openParenToken ?? throw new ArgumentNullException(nameof(openParenToken));
        Count = count ?? throw new ArgumentNullException(nameof(count));
        CloseParenToken = closeParenToken ?? throw new ArgumentNullException(nameof(closeParenToken));
    }

    public SyntaxToken SpcKeyword { get; }

    public SyntaxToken OpenParenToken { get; }

    public ExpressionSyntax Count { get; }

    public SyntaxToken CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(SpcKeyword, OpenParenToken, Count, CloseParenToken);
}

/// <summary>Tab or Tab(column).</summary>
public sealed class TabClauseSyntax : SyntaxNode
{
    public TabClauseSyntax(SyntaxToken tabKeyword, SyntaxToken? openParenToken, ExpressionSyntax? column, SyntaxToken? closeParenToken)
        : base(SyntaxKind.TabClause)
    {
        TabKeyword = tabKeyword ?? throw new ArgumentNullException(nameof(tabKeyword));
        OpenParenToken = openParenToken;
        Column = column;
        CloseParenToken = closeParenToken;
    }

    public SyntaxToken TabKeyword { get; }

    public SyntaxToken? OpenParenToken { get; }

    public ExpressionSyntax? Column { get; }

    public SyntaxToken? CloseParenToken { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() => Children(TabKeyword, OpenParenToken, Column, CloseParenToken);
}

/// <summary>Input #file, variables (MS-VBAL 5.4.5.8).</summary>
public sealed class InputStatementSyntax : StatementSyntax
{
    public InputStatementSyntax(SyntaxToken inputKeyword, FileNumberSyntax fileNumber, SyntaxToken commaToken, SeparatedSyntaxList<ExpressionSyntax> variables)
        : base(SyntaxKind.InputStatement)
    {
        InputKeyword = inputKeyword ?? throw new ArgumentNullException(nameof(inputKeyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        CommaToken = commaToken ?? throw new ArgumentNullException(nameof(commaToken));
        Variables = variables ?? throw new ArgumentNullException(nameof(variables));
    }

    public SyntaxToken InputKeyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken CommaToken { get; }

    public SeparatedSyntaxList<ExpressionSyntax> Variables { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(One(InputKeyword), One(FileNumber), One(CommaToken), Variables.AsChildren());
}

/// <summary>Put #file, [record], variable or Get #file, [record], variable (MS-VBAL 5.4.5.9), told apart by <see cref="SyntaxNode.Kind"/>.</summary>
public sealed class PutGetStatementSyntax : StatementSyntax
{
    public PutGetStatementSyntax(SyntaxKind kind, SyntaxToken keyword, FileNumberSyntax fileNumber, SyntaxToken firstCommaToken, ExpressionSyntax? recordNumber, SyntaxToken secondCommaToken, ExpressionSyntax variable)
        : base(kind)
    {
        Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        FileNumber = fileNumber ?? throw new ArgumentNullException(nameof(fileNumber));
        FirstCommaToken = firstCommaToken ?? throw new ArgumentNullException(nameof(firstCommaToken));
        RecordNumber = recordNumber;
        SecondCommaToken = secondCommaToken ?? throw new ArgumentNullException(nameof(secondCommaToken));
        Variable = variable ?? throw new ArgumentNullException(nameof(variable));
    }

    public SyntaxToken Keyword { get; }

    public FileNumberSyntax FileNumber { get; }

    public SyntaxToken FirstCommaToken { get; }

    public ExpressionSyntax? RecordNumber { get; }

    public SyntaxToken SecondCommaToken { get; }

    public ExpressionSyntax Variable { get; }

    public override IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens() =>
        Children(Keyword, FileNumber, FirstCommaToken, RecordNumber, SecondCommaToken, Variable);
}
