using System.Globalization;

namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Recursive-descent parser for MS-VBAL section 5 over the preprocessed token stream. Produces a
/// full-fidelity tree and never throws: unexpected input yields missing tokens, skipped tokens in
/// <see cref="BadStatementSyntax"/> nodes, and VBA0001 diagnostics, and parsing continues at the
/// next statement (ARCHITECTURE.md section 4, step 2: error recovery from the start).
/// </summary>
internal sealed class Parser
{
    private readonly IReadOnlyList<SyntaxToken> tokens;
    private readonly SourceText source;
    private readonly string filePath;
    private readonly List<Diagnostic> diagnostics;
    private readonly Stack<SyntaxKind> openBlocks = new();
    private int index;
    private int pendingNextCloses;
    private int lastErrorPosition = -1;

    public Parser(IReadOnlyList<SyntaxToken> tokens, SourceText source, string filePath, List<Diagnostic> diagnostics)
    {
        this.tokens = tokens;
        this.source = source;
        this.filePath = filePath;
        this.diagnostics = diagnostics;
    }

    private SyntaxToken Current => tokens[index];

    private SyntaxToken Peek(int offset) => tokens[Math.Min(index + offset, tokens.Count - 1)];

    private SyntaxToken? Previous => index > 0 ? tokens[index - 1] : null;

    private bool AtEnd => Current.Kind == SyntaxKind.EndOfFileToken;

    /// <summary>True when the current token starts a statement (equivalently, the previous statement has ended).</summary>
    private bool AtStatementStart =>
        index == 0 || AtEnd || Previous!.EndsStatement || Previous.Kind == SyntaxKind.HeaderLineToken
        || Current.LeadingTrivia.Any(t => t.IsStatementTerminator);

    private bool AtLineStart =>
        index == 0 || AtEnd || Previous!.EndsLine || Previous.Kind == SyntaxKind.HeaderLineToken
        || Current.LeadingTrivia.Any(t => t.IsEndOfLine);

    // Module structure (MS-VBAL 4.2, 5.1).

    public ModuleSyntax ParseModule()
    {
        ModuleHeaderSyntax? header = null;
        if (Current.Kind == SyntaxKind.HeaderLineToken)
        {
            var lines = new List<SyntaxToken>();
            while (Current.Kind == SyntaxKind.HeaderLineToken)
            {
                lines.Add(Consume());
            }

            header = new ModuleHeaderSyntax(new SyntaxTokenList(lines));
        }

        var members = new List<StatementSyntax>();
        var sawProcedure = false;
        while (!AtEnd)
        {
            if (IsBlockTerminator(out _))
            {
                members.Add(SkipStatement(StrayTerminatorMessage()));
                continue;
            }

            var count = members.Count;
            ParseStatementInto(members);
            for (var i = count; i < members.Count; i++)
            {
                if (members[i] is ProcedureDeclarationSyntax)
                {
                    sawProcedure = true;
                }
                else if (sawProcedure && members[i] is not (AttributeStatementSyntax or BadStatementSyntax))
                {
                    ReportAt(members[i].FirstToken() ?? Current, "Only comments may appear after End Sub, End Function, or End Property.");
                }
            }
        }

        return new ModuleSyntax(header, new SyntaxList<StatementSyntax>(members), Consume());
    }

    /// <summary>Parses one expression that must use up all tokens; used for conditional compilation expressions.</summary>
    public ExpressionSyntax ParseStandaloneExpression()
    {
        var expression = ParseExpression();
        if (!AtEnd)
        {
            Report("Expected: end of statement");
        }

        return expression;
    }

    // Statement lists and error recovery.

    private SyntaxList<StatementSyntax> ParseStatementList(SyntaxKind blockKind)
    {
        openBlocks.Push(blockKind);
        var list = new List<StatementSyntax>();
        try
        {
            while (!AtEnd)
            {
                if (IsBlockTerminator(out var matchesOpenBlock))
                {
                    if (matchesOpenBlock)
                    {
                        break;
                    }

                    pendingNextCloses = 0;
                    list.Add(SkipStatement(StrayTerminatorMessage()));
                    continue;
                }

                ParseStatementInto(list);
            }
        }
        finally
        {
            openBlocks.Pop();
        }

        return new SyntaxList<StatementSyntax>(list);
    }

    /// <summary>The colon-separated statements after Then or Else on a single-line If (MS-VBAL 5.4.2.8).</summary>
    private SyntaxList<StatementSyntax> ParseSingleLineStatements()
    {
        var list = new List<StatementSyntax>();
        while (!AtEnd && Current.Kind != SyntaxKind.ElseKeyword && !(Previous?.EndsLine ?? false))
        {
            if (IsBlockTerminator(out var matchesOpenBlock) && matchesOpenBlock)
            {
                break;
            }

            if (Current.Kind == SyntaxKind.IntegerLiteralToken && list.Count == 0)
            {
                // "If x Then 100" jumps to line 100: an implicit GoTo (MS-VBAL 5.4.2.8 list-or-label).
                var label = Consume();
                list.Add(new GoToStatementSyntax(SyntaxKind.GoToStatement, new SyntaxToken(SyntaxKind.GoToKeyword, string.Empty, label.Start), label));
                continue;
            }

            ParseStatementInto(list);
        }

        return new SyntaxList<StatementSyntax>(list);
    }

    private void ParseStatementInto(List<StatementSyntax> list)
    {
        var start = index;
        var statement = ParseStatement();
        list.Add(statement);
        if (index == start)
        {
            // Nothing consumed: force progress.
            list.Add(SkipStatement("Syntax error"));
            return;
        }

        if (statement is LabelStatementSyntax || AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword)
        {
            return;
        }

        if (IsBlockTerminator(out var matchesOpenBlock) && matchesOpenBlock)
        {
            return;
        }

        Report("Expected: end of statement");
        list.Add(SkipStatement(null));
    }

    /// <summary>Consumes the rest of the current statement into a bad statement, reporting <paramref name="message"/> first if given.</summary>
    private BadStatementSyntax SkipStatement(string? message)
    {
        if (message is not null)
        {
            Report(message);
        }

        var skipped = new List<SyntaxToken>();
        do
        {
            skipped.Add(Consume());
        }
        while (!AtEnd && !AtStatementStart);

        return new BadStatementSyntax(new SyntaxTokenList(skipped));
    }

    /// <summary>
    /// True when the current token closes or continues a block (Else, ElseIf, End If, Case, Next,
    /// Loop, Wend, End Select, End With, End Type, End Enum, End Sub, End Function, End Property,
    /// or a new procedure header inside a procedure). <paramref name="matchesOpenBlock"/> says
    /// whether some enclosing block accepts it; otherwise it is stray.
    /// </summary>
    private bool IsBlockTerminator(out bool matchesOpenBlock)
    {
        var closes = Current.Kind switch
        {
            SyntaxKind.ElseKeyword or SyntaxKind.ElseIfKeyword or SyntaxKind.EndIfKeyword => SyntaxKind.IfBlock,
            SyntaxKind.CaseKeyword => SyntaxKind.SelectCaseBlock,
            SyntaxKind.NextKeyword => SyntaxKind.ForBlock,
            SyntaxKind.LoopKeyword => SyntaxKind.DoLoopBlock,
            SyntaxKind.WendKeyword => SyntaxKind.WhileBlock,
            SyntaxKind.EndKeyword => Peek(1).Kind switch
            {
                SyntaxKind.IfKeyword => SyntaxKind.IfBlock,
                SyntaxKind.SelectKeyword => SyntaxKind.SelectCaseBlock,
                SyntaxKind.WithKeyword => SyntaxKind.WithBlock,
                SyntaxKind.TypeKeyword => SyntaxKind.TypeDefinition,
                SyntaxKind.EnumKeyword => SyntaxKind.EnumDefinition,
                SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword => SyntaxKind.ProcedureDeclaration,
                SyntaxKind.IdentifierToken when IsWord(Peek(1), "Property") => SyntaxKind.ProcedureDeclaration,
                _ => SyntaxKind.None,
            },
            _ => SyntaxKind.None,
        };

        if (closes == SyntaxKind.None)
        {
            if (openBlocks.Contains(SyntaxKind.ProcedureDeclaration) && IsProcedureStart())
            {
                matchesOpenBlock = true;
                return true;
            }

            matchesOpenBlock = false;
            return false;
        }

        matchesOpenBlock = closes == SyntaxKind.ForBlock
            ? openBlocks.Contains(SyntaxKind.ForBlock) || openBlocks.Contains(SyntaxKind.ForEachBlock)
            : openBlocks.Contains(closes);
        return true;
    }

    private string StrayTerminatorMessage() => Current.Kind switch
    {
        SyntaxKind.ElseKeyword or SyntaxKind.ElseIfKeyword => "Else without If",
        SyntaxKind.EndIfKeyword => "End If without block If",
        SyntaxKind.CaseKeyword => "Case without Select Case",
        SyntaxKind.NextKeyword => "Next without For",
        SyntaxKind.LoopKeyword => "Loop without Do",
        SyntaxKind.WendKeyword => "Wend without While",
        SyntaxKind.EndKeyword => Peek(1).Kind switch
        {
            SyntaxKind.IfKeyword => "End If without block If",
            SyntaxKind.SelectKeyword => "End Select without Select Case",
            SyntaxKind.WithKeyword => "End With without With",
            SyntaxKind.TypeKeyword => "End Type without Type",
            SyntaxKind.EnumKeyword => "End Enum without Enum",
            _ => "End " + Peek(1).Text + " without " + Peek(1).Text,
        },
        _ => "Syntax error",
    };

    /// <summary>[Public | Private | Friend | Static]* (Sub | Function | Property Get/Let/Set).</summary>
    private bool IsProcedureStart()
    {
        var i = 0;
        while (Peek(i).Kind is SyntaxKind.PublicKeyword or SyntaxKind.PrivateKeyword or SyntaxKind.FriendKeyword or SyntaxKind.StaticKeyword)
        {
            i++;
        }

        var token = Peek(i);
        return token.Kind is SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword
            || (IsWord(token, "Property") && Peek(i + 1).Kind is SyntaxKind.GetKeyword or SyntaxKind.LetKeyword or SyntaxKind.SetKeyword);
    }

    // Statements (MS-VBAL 5.2, 5.3, 5.4).

    private StatementSyntax ParseStatement()
    {
        if (AtLineStart)
        {
            if (Current.Kind == SyntaxKind.IntegerLiteralToken)
            {
                return new LabelStatementSyntax(Consume(), Current.Kind == SyntaxKind.ColonToken ? Consume() : null);
            }

            if (Current.Kind == SyntaxKind.IdentifierToken && Peek(1).Kind == SyntaxKind.ColonToken)
            {
                return new LabelStatementSyntax(Consume(), Consume());
            }
        }

        switch (Current.Kind)
        {
            case SyntaxKind.AttributeKeyword:
                return ParseAttribute();
            case SyntaxKind.OptionKeyword:
                return ParseOption();
            case SyntaxKind.DefBoolKeyword or SyntaxKind.DefByteKeyword or SyntaxKind.DefCurKeyword or SyntaxKind.DefDateKeyword
                or SyntaxKind.DefDblKeyword or SyntaxKind.DefIntKeyword or SyntaxKind.DefLngKeyword or SyntaxKind.DefLngLngKeyword
                or SyntaxKind.DefLngPtrKeyword or SyntaxKind.DefObjKeyword or SyntaxKind.DefSngKeyword or SyntaxKind.DefStrKeyword
                or SyntaxKind.DefVarKeyword or SyntaxKind.DefDecKeyword:
                return ParseDefType();
            case SyntaxKind.DimKeyword or SyntaxKind.StaticKeyword or SyntaxKind.PrivateKeyword or SyntaxKind.PublicKeyword
                or SyntaxKind.GlobalKeyword or SyntaxKind.FriendKeyword or SyntaxKind.ConstKeyword or SyntaxKind.TypeKeyword
                or SyntaxKind.EnumKeyword or SyntaxKind.DeclareKeyword or SyntaxKind.EventKeyword or SyntaxKind.SubKeyword
                or SyntaxKind.FunctionKeyword:
                return ParseDeclaration();
            case SyntaxKind.IdentifierToken when IsWord(Current, "Property") && Peek(1).Kind is SyntaxKind.GetKeyword or SyntaxKind.LetKeyword or SyntaxKind.SetKeyword:
                return ParseDeclaration();
            case SyntaxKind.ImplementsKeyword:
                return new ImplementsStatementSyntax(Consume(), ParseTypeName());
            case SyntaxKind.IfKeyword:
                return ParseIf();
            case SyntaxKind.SelectKeyword:
                return ParseSelect();
            case SyntaxKind.ForKeyword:
                return ParseFor();
            case SyntaxKind.DoKeyword:
                return ParseDo();
            case SyntaxKind.WhileKeyword:
                return ParseWhile();
            case SyntaxKind.WithKeyword:
                return ParseWith();
            case SyntaxKind.GoToKeyword:
                return new GoToStatementSyntax(SyntaxKind.GoToStatement, Consume(), ExpectLabel());
            case SyntaxKind.GoSubKeyword:
                return new GoToStatementSyntax(SyntaxKind.GoSubStatement, Consume(), ExpectLabel());
            case SyntaxKind.ReturnKeyword:
                return new KeywordStatementSyntax(SyntaxKind.ReturnStatement, Consume());
            case SyntaxKind.EndKeyword:
                return new KeywordStatementSyntax(SyntaxKind.EndStatement, Consume());
            case SyntaxKind.StopKeyword:
                return new KeywordStatementSyntax(SyntaxKind.StopStatement, Consume());
            case SyntaxKind.OnKeyword:
                return ParseOn();
            case SyntaxKind.ResumeKeyword:
                return ParseResume();
            case SyntaxKind.ExitKeyword:
                return ParseExit();
            case SyntaxKind.RaiseEventKeyword:
                return ParseRaiseEvent();
            case SyntaxKind.ReDimKeyword:
                return ParseReDim();
            case SyntaxKind.EraseKeyword:
                return new EraseStatementSyntax(Consume(), ParseSeparatedList(ParseExpression));
            case SyntaxKind.LSetKeyword:
                return ParseLSetRSet(SyntaxKind.LSetStatement);
            case SyntaxKind.RSetKeyword:
                return ParseLSetRSet(SyntaxKind.RSetStatement);
            case SyntaxKind.LetKeyword:
                return ParseAssignment(SyntaxKind.LetAssignmentStatement, Consume());
            case SyntaxKind.SetKeyword:
                return ParseAssignment(SyntaxKind.SetAssignmentStatement, Consume());
            case SyntaxKind.CallKeyword:
                return new CallStatementSyntax(Consume(), ParsePostfixExpression(), null);
            case SyntaxKind.OpenKeyword:
                return ParseOpen();
            case SyntaxKind.CloseKeyword:
                return ParseClose();
            case SyntaxKind.SeekKeyword:
                return ParseSeek();
            case SyntaxKind.LockKeyword or SyntaxKind.UnlockKeyword:
                return ParseLock();
            case SyntaxKind.PrintKeyword:
                return ParseFileOutput(SyntaxKind.PrintStatement);
            case SyntaxKind.WriteKeyword:
                return ParseFileOutput(SyntaxKind.WriteStatement);
            case SyntaxKind.InputKeyword:
                return ParseInput();
            case SyntaxKind.PutKeyword:
                return ParsePutGet(SyntaxKind.PutStatement);
            case SyntaxKind.GetKeyword:
                return ParsePutGet(SyntaxKind.GetStatement);
            case SyntaxKind.DebugKeyword when Peek(1).Kind == SyntaxKind.DotToken && Peek(2).Kind == SyntaxKind.PrintKeyword:
                return new DebugPrintStatementSyntax(Consume(), Consume(), Consume(), ParseOutputList());
            case SyntaxKind.IdentifierToken:
                return ParseIdentifierStatement();
            default:
                if (CanStartExpression(Current))
                {
                    return ParseExpressionStatement();
                }

                return SkipStatement("Syntax error");
        }
    }

    /// <summary>Statements introduced by a non-reserved word: Name, Error, Mid, Line Input, Width, Reset; else a call or assignment.</summary>
    private StatementSyntax ParseIdentifierStatement()
    {
        var next = Peek(1);
        var nextContinuesExpression = next.Kind is SyntaxKind.EqualsToken or SyntaxKind.OpenParenToken or SyntaxKind.DotToken
            or SyntaxKind.BangToken or SyntaxKind.ColonEqualsToken;
        var nextEndsStatement = Current.EndsStatement || next.Kind == SyntaxKind.EndOfFileToken;

        if (IsWord(Current, "Name") && !nextContinuesExpression && !nextEndsStatement)
        {
            var nameToken = Consume();
            var oldName = ParseExpression();
            var asKeyword = Expect(SyntaxKind.AsKeyword);
            return new NameStatementSyntax(nameToken, oldName, asKeyword, ParseExpression());
        }

        if (IsWord(Current, "Error") && !nextEndsStatement && next.Kind is not (SyntaxKind.EqualsToken or SyntaxKind.DotToken or SyntaxKind.BangToken or SyntaxKind.ColonEqualsToken))
        {
            return new ErrorStatementSyntax(Consume(), ParseExpression());
        }

        // Mid, Mid$, MidB, and MidB$ all start the Mid statement (MS-VBAL 5.4.3.3); the $ is the type suffix.
        if (IsMidWord(Current) && next.Kind == SyntaxKind.OpenParenToken && IsMidStatement())
        {
            return ParseMid();
        }

        if (IsWord(Current, "Line") && next.Kind == SyntaxKind.InputKeyword)
        {
            var lineToken = Consume();
            var inputKeyword = Consume();
            var fileNumber = ParseFileNumber();
            var comma = Expect(SyntaxKind.CommaToken);
            return new LineInputStatementSyntax(lineToken, inputKeyword, fileNumber, comma, ParsePostfixExpression());
        }

        if (IsWord(Current, "Width") && next.Kind == SyntaxKind.HashToken)
        {
            var widthToken = Consume();
            var fileNumber = ParseFileNumber();
            var comma = Expect(SyntaxKind.CommaToken);
            return new WidthStatementSyntax(widthToken, fileNumber, comma, ParseExpression());
        }

        if (IsWord(Current, "Reset") && nextEndsStatement)
        {
            return new KeywordStatementSyntax(SyntaxKind.ResetStatement, Consume());
        }

        return ParseExpressionStatement();
    }

    /// <summary>True when the "(" at the current position closes right before As, a comma, or the end of the statement.</summary>
    private bool IsFinalParenthesizedGroup()
    {
        var depth = 0;
        for (var i = 0; index + i < tokens.Count; i++)
        {
            var token = Peek(i);
            if (token.Kind == SyntaxKind.OpenParenToken)
            {
                depth++;
            }
            else if (token.Kind == SyntaxKind.CloseParenToken)
            {
                depth--;
                if (depth == 0)
                {
                    var next = Peek(i + 1);
                    return token.EndsStatement || next.Kind is SyntaxKind.AsKeyword or SyntaxKind.CommaToken or SyntaxKind.EndOfFileToken or SyntaxKind.ElseKeyword;
                }
            }

            if (token.Kind == SyntaxKind.EndOfFileToken || (token.EndsStatement && depth > 0))
            {
                return true;
            }
        }

        return true;
    }

    /// <summary>Looks past "Mid(...)" for the "=" that makes it a Mid statement rather than a call.</summary>
    private bool IsMidStatement()
    {
        var depth = 0;
        for (var i = 1; index + i < tokens.Count; i++)
        {
            var token = Peek(i);
            if (token.Kind == SyntaxKind.OpenParenToken)
            {
                depth++;
            }
            else if (token.Kind == SyntaxKind.CloseParenToken)
            {
                depth--;
                if (depth == 0)
                {
                    return Peek(i + 1).Kind == SyntaxKind.EqualsToken;
                }
            }
            else if (token.Kind == SyntaxKind.EndOfFileToken)
            {
                return false;
            }

            if (token.EndsStatement)
            {
                return false;
            }
        }

        return false;
    }

    private MidStatementSyntax ParseMid()
    {
        var midKeyword = Consume();
        var open = Expect(SyntaxKind.OpenParenToken);
        var target = ParseExpression();
        var firstComma = Expect(SyntaxKind.CommaToken);
        var start = ParseExpression();
        SyntaxToken? secondComma = null;
        ExpressionSyntax? length = null;
        if (Current.Kind == SyntaxKind.CommaToken)
        {
            secondComma = Consume();
            length = ParseExpression();
        }

        var close = Expect(SyntaxKind.CloseParenToken);
        var equals = Expect(SyntaxKind.EqualsToken);
        return new MidStatementSyntax(midKeyword, open, target, firstComma, start, secondComma, length, close, equals, ParseExpression());
    }

    /// <summary>An assignment or a call written without Call (MS-VBAL 5.4.2.1, 5.4.3.8).</summary>
    private StatementSyntax ParseExpressionStatement()
    {
        var expression = ParsePostfixExpression();
        if (Current.Kind == SyntaxKind.EqualsToken)
        {
            var equals = Consume();
            return new AssignmentStatementSyntax(SyntaxKind.LetAssignmentStatement, null, expression, equals, ParseExpression());
        }

        // "Foo (1), 2" is a call whose first argument is the parenthesized "(1)" (MS-VBAL 5.4.2.1): the
        // postfix parser saw an index; reinterpret it when more arguments follow, or when an operator
        // continues the first argument, as in VBA-TDD's "Check (a = 0 Or b) And (c), message".
        ArgumentSyntax? first = null;
        var operatorFollows = !AtStatementStart && SyntaxFacts.GetBinaryOperatorPrecedence(Current.Kind) > 0;
        if ((Current.Kind == SyntaxKind.CommaToken || operatorFollows)
            && expression is IndexExpressionSyntax { Arguments: { OpenParenToken: { } open, CloseParenToken: { } close } args } indexed
            && args.Arguments.Count == 1
            && args.Arguments[0] is { Name: null, ByValKeyword: null, Expression: { } inner })
        {
            expression = indexed.Expression;
            ExpressionSyntax argument = new ParenthesizedExpressionSyntax(open, inner, close);
            if (operatorFollows)
            {
                argument = ParseBinaryRest(argument, 1);
            }

            first = new ArgumentSyntax(null, null, null, argument);
            return new CallStatementSyntax(null, expression, ParseBareArgumentList(first));
        }

        // A keyword followed by := is a named argument (Sheets.Add Type:="Worksheet"), not a new statement.
        var namedByKeyword = Current.IsKeyword && Peek(1).Kind == SyntaxKind.ColonEqualsToken;
        if (AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword || (!CanStartArgument(Current) && !namedByKeyword))
        {
            return new CallStatementSyntax(null, expression, null);
        }

        return new CallStatementSyntax(null, expression, ParseBareArgumentList(first));
    }

    private AssignmentStatementSyntax ParseAssignment(SyntaxKind kind, SyntaxToken keyword)
    {
        var target = ParsePostfixExpression();
        var equals = Expect(SyntaxKind.EqualsToken);
        return new AssignmentStatementSyntax(kind, keyword, target, equals, ParseExpression());
    }

    private LSetRSetStatementSyntax ParseLSetRSet(SyntaxKind kind)
    {
        var keyword = Consume();
        var target = ParsePostfixExpression();
        var equals = Expect(SyntaxKind.EqualsToken);
        return new LSetRSetStatementSyntax(kind, keyword, target, equals, ParseExpression());
    }

    private AttributeStatementSyntax ParseAttribute()
    {
        var keyword = Consume();
        ExpressionSyntax name = new IdentifierNameSyntax(ExpectName());
        while (Current.Kind == SyntaxKind.DotToken)
        {
            var dot = Consume();
            name = new MemberAccessExpressionSyntax(name, dot, new IdentifierNameSyntax(ExpectName()));
        }

        var equals = Expect(SyntaxKind.EqualsToken);
        return new AttributeStatementSyntax(keyword, name, equals, ParseSeparatedList(ParseExpression));
    }

    // Option Explicit / Base / Compare / Private Module (MS-VBAL 5.2.1).
    private OptionStatementSyntax ParseOption()
    {
        var keyword = Consume();
        if (Current.Kind == SyntaxKind.PrivateKeyword)
        {
            var privateKeyword = Consume();
            return new OptionStatementSyntax(keyword, privateKeyword, ExpectWord("Module"));
        }

        if (IsWord(Current, "Explicit"))
        {
            return new OptionStatementSyntax(keyword, Consume(), null);
        }

        if (IsWord(Current, "Base"))
        {
            var name = Consume();
            return new OptionStatementSyntax(keyword, name, Expect(SyntaxKind.IntegerLiteralToken));
        }

        if (IsWord(Current, "Compare"))
        {
            var name = Consume();
            if (IsWord(Current, "Binary") || IsWord(Current, "Text") || IsWord(Current, "Database"))
            {
                return new OptionStatementSyntax(keyword, name, Consume());
            }

            Report("Expected: Binary, Text, or Database");
            return new OptionStatementSyntax(keyword, name, SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start));
        }

        Report("Expected: Explicit, Base, Compare, or Private");
        return new OptionStatementSyntax(keyword, SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start), null);
    }

    // DefInt A-Z (MS-VBAL 5.2.2).
    private DefTypeStatementSyntax ParseDefType()
    {
        var keyword = Consume();
        return new DefTypeStatementSyntax(keyword, ParseSeparatedList(() =>
        {
            var first = ExpectIdentifier();
            if (Current.Kind == SyntaxKind.MinusToken)
            {
                var minus = Consume();
                return new LetterRangeSyntax(first, minus, ExpectIdentifier());
            }

            return new LetterRangeSyntax(first, null, null);
        }));
    }

    // Declarations with modifiers (MS-VBAL 5.2.3, 5.3.1).
    private StatementSyntax ParseDeclaration()
    {
        var modifiers = new List<SyntaxToken>();
        while (Current.Kind is SyntaxKind.DimKeyword or SyntaxKind.StaticKeyword or SyntaxKind.PrivateKeyword or SyntaxKind.PublicKeyword
            or SyntaxKind.GlobalKeyword or SyntaxKind.FriendKeyword or SyntaxKind.SharedKeyword)
        {
            modifiers.Add(Consume());
        }

        var modifierList = new SyntaxTokenList(modifiers);
        switch (Current.Kind)
        {
            case SyntaxKind.ConstKeyword:
                return ParseConst(modifierList);
            case SyntaxKind.TypeKeyword:
                return ParseTypeDefinition(modifierList);
            case SyntaxKind.EnumKeyword:
                return ParseEnum(modifierList);
            case SyntaxKind.DeclareKeyword:
                return ParseDeclare(modifierList);
            case SyntaxKind.EventKeyword:
                return ParseEvent(modifierList);
            case SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword:
                return ParseProcedure(modifierList);
            case SyntaxKind.IdentifierToken when IsWord(Current, "Property") && Peek(1).Kind is SyntaxKind.GetKeyword or SyntaxKind.LetKeyword or SyntaxKind.SetKeyword:
                return ParseProcedure(modifierList);
            default:
                return new VariableDeclarationSyntax(modifierList, ParseSeparatedList(ParseVariableDeclarator));
        }
    }

    private VariableDeclaratorSyntax ParseVariableDeclarator()
    {
        var withEvents = Current.Kind == SyntaxKind.WithEventsKeyword ? Consume() : null;
        var name = ExpectIdentifier();
        var bounds = Current.Kind == SyntaxKind.OpenParenToken ? ParseArrayBounds() : null;
        var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: false) : null;
        return new VariableDeclaratorSyntax(withEvents, name, bounds, asClause);
    }

    private ConstDeclarationSyntax ParseConst(SyntaxTokenList modifiers)
    {
        var keyword = Consume();
        return new ConstDeclarationSyntax(modifiers, keyword, ParseSeparatedList(() =>
        {
            var name = ExpectIdentifier();
            var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: false) : null;
            var equals = Expect(SyntaxKind.EqualsToken);
            return new ConstDeclaratorSyntax(name, asClause, equals, ParseExpression());
        }));
    }

    private ArrayBoundsSyntax ParseArrayBounds()
    {
        var open = Expect(SyntaxKind.OpenParenToken);
        var bounds = Current.Kind == SyntaxKind.CloseParenToken
            ? new SeparatedSyntaxList<BoundSyntax>([])
            : ParseSeparatedList(() =>
            {
                var first = ParseExpression();
                if (Current.Kind == SyntaxKind.ToKeyword)
                {
                    var to = Consume();
                    return new BoundSyntax(first, to, ParseExpression());
                }

                return new BoundSyntax(null, null, first);
            });
        return new ArrayBoundsSyntax(open, bounds, Expect(SyntaxKind.CloseParenToken));
    }

    private AsClauseSyntax ParseAsClause(bool allowArrayDesignator)
    {
        var asKeyword = Expect(SyntaxKind.AsKeyword);
        var newKeyword = Current.Kind == SyntaxKind.NewKeyword ? Consume() : null;
        TypeSyntax type;
        if (SyntaxFacts.IsBuiltinTypeKeyword(Current.Kind))
        {
            var keyword = Consume();
            if (keyword.Kind == SyntaxKind.StringKeyword && Current.Kind == SyntaxKind.AsteriskToken)
            {
                var asterisk = Consume();
                type = new BuiltinTypeSyntax(keyword, asterisk, ParseExpression());
            }
            else
            {
                type = new BuiltinTypeSyntax(keyword, null, null);
            }
        }
        else
        {
            type = new NamedTypeSyntax(ParseTypeName());
        }

        ArrayDesignatorSyntax? designator = null;
        if (allowArrayDesignator && Current.Kind == SyntaxKind.OpenParenToken && Peek(1).Kind == SyntaxKind.CloseParenToken)
        {
            designator = new ArrayDesignatorSyntax(Consume(), Consume());
        }

        return new AsClauseSyntax(asKeyword, newKeyword, type, designator);
    }

    /// <summary>A possibly qualified type or class name: Name, Lib.Name, Lib.Sub.Name.</summary>
    private ExpressionSyntax ParseTypeName()
    {
        ExpressionSyntax name = new IdentifierNameSyntax(ExpectName());
        while (Current.Kind == SyntaxKind.DotToken && !AtStatementStart)
        {
            var dot = Consume();
            name = new MemberAccessExpressionSyntax(name, dot, new IdentifierNameSyntax(ExpectName()));
        }

        return name;
    }

    private TypeDefinitionSyntax ParseTypeDefinition(SyntaxTokenList modifiers)
    {
        var keyword = Consume();
        var name = ExpectIdentifier();
        // MS-VBAL 5.2.3.3: udt-member = reserved-name-member-dcl / untyped-name-member-dcl, so
        // "Type As Long" (the Windows PICTDESC struct) or "Next As Long" are valid members.
        var members = ParseBlockMembers(SyntaxKind.TypeDefinition, IsReservedMemberName, () =>
        {
            var memberName = Current.IsKeyword ? Consume() : ExpectIdentifier();
            var bounds = Current.Kind == SyntaxKind.OpenParenToken ? ParseArrayBounds() : null;
            return new TypeMemberSyntax(memberName, bounds, ParseAsClause(allowArrayDesignator: false));
        });
        return new TypeDefinitionSyntax(modifiers, keyword, name, members, ExpectEndBlock(SyntaxKind.TypeKeyword));
    }

    private EnumDefinitionSyntax ParseEnum(SyntaxTokenList modifiers)
    {
        var keyword = Consume();
        var name = ExpectIdentifier();
        var members = ParseBlockMembers(SyntaxKind.EnumDefinition, static _ => false, () =>
        {
            var memberName = ExpectIdentifier();
            if (Current.Kind == SyntaxKind.EqualsToken)
            {
                var equals = Consume();
                return new EnumMemberSyntax(memberName, equals, ParseExpression());
            }

            return new EnumMemberSyntax(memberName, null, null);
        });
        return new EnumDefinitionSyntax(modifiers, keyword, name, members, ExpectEndBlock(SyntaxKind.EnumKeyword));
    }

    /// <summary>MS-VBAL 5.2.3.3 reserved-member-name: every reserved identifier except the reserved type identifiers.</summary>
    private static bool IsReservedMemberName(SyntaxToken token) =>
        token.IsKeyword && token.Kind is not (
            SyntaxKind.BooleanKeyword or SyntaxKind.ByteKeyword or SyntaxKind.CurrencyKeyword or SyntaxKind.DoubleKeyword
            or SyntaxKind.IntegerKeyword or SyntaxKind.LongKeyword or SyntaxKind.LongLongKeyword or SyntaxKind.LongPtrKeyword
            or SyntaxKind.SingleKeyword or SyntaxKind.VariantKeyword);

    /// <summary>Members of a Type or Enum block: one per statement, until End Type / End Enum.</summary>
    private SyntaxList<StatementSyntax> ParseBlockMembers(SyntaxKind blockKind, Func<SyntaxToken, bool> allowsKeywordName, Func<StatementSyntax> parseMember)
    {
        openBlocks.Push(blockKind);
        var list = new List<StatementSyntax>();
        try
        {
            while (!AtEnd)
            {
                // A reserved member name such as "Next As Long" is a member, not a block terminator.
                var isKeywordMember = allowsKeywordName(Current) && Peek(1).Kind is SyntaxKind.AsKeyword or SyntaxKind.OpenParenToken;
                if (!isKeywordMember && IsBlockTerminator(out var matchesOpenBlock))
                {
                    if (matchesOpenBlock)
                    {
                        break;
                    }

                    list.Add(SkipStatement(StrayTerminatorMessage()));
                    continue;
                }

                if (Current.Kind is not (SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken) && !isKeywordMember)
                {
                    list.Add(SkipStatement("Expected: identifier"));
                    continue;
                }

                var start = index;
                list.Add(parseMember());
                if (index == start)
                {
                    list.Add(SkipStatement("Syntax error"));
                }
                else if (!AtStatementStart)
                {
                    Report("Expected: end of statement");
                    list.Add(SkipStatement(null));
                }
            }
        }
        finally
        {
            openBlocks.Pop();
        }

        return new SyntaxList<StatementSyntax>(list);
    }

    private DeclareStatementSyntax ParseDeclare(SyntaxTokenList modifiers)
    {
        var keyword = Consume();
        var ptrSafe = IsWord(Current, "PtrSafe") ? Consume() : null;
        SyntaxToken procedureKeyword;
        if (Current.Kind is SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword)
        {
            procedureKeyword = Consume();
        }
        else
        {
            Report("Expected: Sub or Function");
            procedureKeyword = SyntaxToken.Missing(SyntaxKind.SubKeyword, Current.Start);
        }

        var name = ExpectIdentifier();
        var cdecl = Current.Kind == SyntaxKind.CDeclKeyword ? Consume() : null;
        var lib = ExpectWord("Lib");
        var libraryName = Expect(SyntaxKind.StringLiteralToken);
        SyntaxToken? alias = null;
        SyntaxToken? aliasName = null;
        if (IsWord(Current, "Alias"))
        {
            alias = Consume();
            aliasName = Expect(SyntaxKind.StringLiteralToken);
        }

        var parameters = Current.Kind == SyntaxKind.OpenParenToken ? ParseParameterList() : null;
        var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: true) : null;
        return new DeclareStatementSyntax(modifiers, keyword, ptrSafe, procedureKeyword, name, cdecl, lib, libraryName, alias, aliasName, parameters, asClause);
    }

    private EventDeclarationSyntax ParseEvent(SyntaxTokenList modifiers)
    {
        var keyword = Consume();
        var name = ExpectIdentifier();
        var parameters = Current.Kind == SyntaxKind.OpenParenToken ? ParseParameterList() : null;
        return new EventDeclarationSyntax(modifiers, keyword, name, parameters);
    }

    // Sub, Function, Property (MS-VBAL 5.3.1).
    private ProcedureDeclarationSyntax ParseProcedure(SyntaxTokenList modifiers)
    {
        var procedureKeyword = Consume();
        SyntaxToken? accessor = null;
        if (IsWord(procedureKeyword, "Property"))
        {
            if (Current.Kind is SyntaxKind.GetKeyword or SyntaxKind.LetKeyword or SyntaxKind.SetKeyword)
            {
                accessor = Consume();
            }
            else
            {
                Report("Expected: Get, Let, or Set");
                accessor = SyntaxToken.Missing(SyntaxKind.GetKeyword, Current.Start);
            }
        }

        var name = ExpectIdentifier();
        var parameters = Current.Kind == SyntaxKind.OpenParenToken ? ParseParameterList() : null;
        var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: true) : null;
        var body = ParseStatementList(SyntaxKind.ProcedureDeclaration);

        EndBlockStatementSyntax end;
        var expectedEnd = accessor is not null ? "Property" : procedureKeyword.Text;
        if (Current.Kind == SyntaxKind.EndKeyword
            && (Peek(1).Kind is SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword || IsWord(Peek(1), "Property")))
        {
            var endKeyword = Consume();
            var blockKeyword = Consume();

            // VBA does not check which of the three closes a procedure: End Sub, End Function, and
            // End Property are interchangeable there, which the VBE accepts in all six pairings and
            // stdVBA's sources rely on (docs/vba-quirks.md; R3). The token is kept as it was
            // written, so printing gives the file back byte for byte (R14).
            end = new EndBlockStatementSyntax(endKeyword, blockKeyword);
        }
        else
        {
            ReportMissingBlockEnd("Expected: End " + expectedEnd);
            end = new EndBlockStatementSyntax(SyntaxToken.Missing(SyntaxKind.EndKeyword, Current.Start), SyntaxToken.Missing(procedureKeyword.Kind, Current.Start));
        }

        return new ProcedureDeclarationSyntax(modifiers, procedureKeyword, accessor, name, parameters, asClause, body, end);
    }

    private ParameterListSyntax ParseParameterList()
    {
        var open = Expect(SyntaxKind.OpenParenToken);
        var parameters = Current.Kind == SyntaxKind.CloseParenToken
            ? new SeparatedSyntaxList<ParameterSyntax>([])
            : ParseSeparatedList(() =>
            {
                var modifiers = new List<SyntaxToken>();
                while (Current.Kind is SyntaxKind.OptionalKeyword or SyntaxKind.ByValKeyword or SyntaxKind.ByRefKeyword or SyntaxKind.ParamArrayKeyword)
                {
                    modifiers.Add(Consume());
                }

                var name = ExpectIdentifier();
                var bounds = Current.Kind == SyntaxKind.OpenParenToken ? ParseArrayBounds() : null;
                var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: false) : null;
                SyntaxToken? equals = null;
                ExpressionSyntax? defaultValue = null;
                if (Current.Kind == SyntaxKind.EqualsToken)
                {
                    equals = Consume();
                    defaultValue = ParseExpression();
                }

                return new ParameterSyntax(new SyntaxTokenList(modifiers), name, bounds, asClause, equals, defaultValue);
            });
        return new ParameterListSyntax(open, parameters, Expect(SyntaxKind.CloseParenToken));
    }

    // Control flow (MS-VBAL 5.4.2).

    private StatementSyntax ParseIf()
    {
        var ifKeyword = Consume();
        var condition = ParseExpression();
        var thenKeyword = Expect(SyntaxKind.ThenKeyword);
        var singleLine = !AtEnd && !(Previous?.EndsLine ?? false) && !Current.LeadingTrivia.Any(t => t.IsEndOfLine);
        if (singleLine)
        {
            var statements = ParseSingleLineStatements();
            SingleLineElseClauseSyntax? elseClause = null;
            if (Current.Kind == SyntaxKind.ElseKeyword && !AtStatementStart)
            {
                var elseKeyword = Consume();
                elseClause = new SingleLineElseClauseSyntax(elseKeyword, ParseSingleLineStatements());
            }

            return new SingleLineIfStatementSyntax(ifKeyword, condition, thenKeyword, statements, elseClause);
        }

        var body = ParseStatementList(SyntaxKind.IfBlock);
        var elseIfBlocks = new List<ElseIfBlockSyntax>();
        while (Current.Kind == SyntaxKind.ElseIfKeyword)
        {
            var elseIfKeyword = Consume();
            var elseIfCondition = ParseExpression();
            var elseIfThen = Expect(SyntaxKind.ThenKeyword);
            elseIfBlocks.Add(new ElseIfBlockSyntax(elseIfKeyword, elseIfCondition, elseIfThen, ParseStatementList(SyntaxKind.IfBlock)));
        }

        ElseBlockSyntax? elseBlock = null;
        if (Current.Kind == SyntaxKind.ElseKeyword)
        {
            var elseKeyword = Consume();
            elseBlock = new ElseBlockSyntax(elseKeyword, ParseStatementList(SyntaxKind.IfBlock));
        }

        EndBlockStatementSyntax endIf;
        if (Current.Kind == SyntaxKind.EndIfKeyword)
        {
            endIf = new EndBlockStatementSyntax(Consume(), null);
        }
        else
        {
            endIf = ExpectEndBlock(SyntaxKind.IfKeyword);
        }

        return new IfBlockSyntax(ifKeyword, condition, thenKeyword, body, new SyntaxList<ElseIfBlockSyntax>(elseIfBlocks), elseBlock, endIf);
    }

    private SelectCaseBlockSyntax ParseSelect()
    {
        var selectKeyword = Consume();
        var caseKeyword = Expect(SyntaxKind.CaseKeyword);
        var expression = ParseExpression();
        var caseBlocks = new List<CaseBlockSyntax>();

        // Only trivia may sit between Select Case and the first Case.
        var stray = ParseStatementList(SyntaxKind.SelectCaseBlock);
        if (stray.Count > 0)
        {
            ReportAt(stray[0].FirstToken() ?? Current, "Expected: Case");
            caseBlocks.Add(new CaseBlockSyntax(SyntaxToken.Missing(SyntaxKind.CaseKeyword, stray[0].Start), null, new SeparatedSyntaxList<CaseClauseSyntax>([]), stray));
        }

        while (Current.Kind == SyntaxKind.CaseKeyword)
        {
            var keyword = Consume();
            SyntaxToken? elseKeyword = null;
            var clauses = new SeparatedSyntaxList<CaseClauseSyntax>([]);
            if (Current.Kind == SyntaxKind.ElseKeyword)
            {
                elseKeyword = Consume();
            }
            else
            {
                clauses = ParseSeparatedList(ParseCaseClause);
            }

            caseBlocks.Add(new CaseBlockSyntax(keyword, elseKeyword, clauses, ParseStatementList(SyntaxKind.SelectCaseBlock)));
        }

        return new SelectCaseBlockSyntax(selectKeyword, caseKeyword, expression, new SyntaxList<CaseBlockSyntax>(caseBlocks), ExpectEndBlock(SyntaxKind.SelectKeyword));
    }

    private CaseClauseSyntax ParseCaseClause()
    {
        if (Current.Kind == SyntaxKind.IsKeyword)
        {
            var isKeyword = Consume();
            SyntaxToken operatorToken;
            if (Current.Kind is SyntaxKind.EqualsToken or SyntaxKind.LessThanGreaterThanToken or SyntaxKind.LessThanToken
                or SyntaxKind.GreaterThanToken or SyntaxKind.LessThanEqualsToken or SyntaxKind.GreaterThanEqualsToken)
            {
                operatorToken = Consume();
            }
            else
            {
                Report("Expected: comparison operator");
                operatorToken = SyntaxToken.Missing(SyntaxKind.EqualsToken, Current.Start);
            }

            return new IsCaseClauseSyntax(isKeyword, operatorToken, ParseExpression());
        }

        var value = ParseExpression();
        if (Current.Kind == SyntaxKind.ToKeyword)
        {
            var toKeyword = Consume();
            return new RangeCaseClauseSyntax(value, toKeyword, ParseExpression());
        }

        return new ValueCaseClauseSyntax(value);
    }

    private StatementSyntax ParseFor()
    {
        var forKeyword = Consume();
        if (Current.Kind == SyntaxKind.EachKeyword)
        {
            var eachKeyword = Consume();
            var element = ParsePostfixExpression();
            var inKeyword = Expect(SyntaxKind.InKeyword);
            var collection = ParseExpression();
            var body = ParseStatementList(SyntaxKind.ForEachBlock);
            return new ForEachBlockSyntax(forKeyword, eachKeyword, element, inKeyword, collection, body, ParseNextFor());
        }

        var variable = ParsePostfixExpression();
        var equals = Expect(SyntaxKind.EqualsToken);
        var from = ParseExpression();
        var toKeyword = Expect(SyntaxKind.ToKeyword);
        var to = ParseExpression();
        StepClauseSyntax? step = null;
        if (IsWord(Current, "Step") && !AtStatementStart)
        {
            var stepKeyword = Consume();
            step = new StepClauseSyntax(stepKeyword, ParseExpression());
        }

        var statements = ParseStatementList(SyntaxKind.ForBlock);
        return new ForBlockSyntax(forKeyword, variable, equals, from, toKeyword, to, step, statements, ParseNextFor());
    }

    /// <summary>
    /// The Next of a For loop. "Next a, b" closes two loops: the innermost leaves it for the
    /// enclosing loop, and the outermost loop it names takes the statement.
    /// </summary>
    private NextStatementSyntax? ParseNextFor()
    {
        if (Current.Kind != SyntaxKind.NextKeyword)
        {
            ReportMissingBlockEnd("Expected: Next");
            return null;
        }

        if (pendingNextCloses > 0)
        {
            pendingNextCloses--;
            return pendingNextCloses == 0 ? ParseNextStatement() : null;
        }

        var variables = CountNextVariables();
        if (variables <= 1)
        {
            return ParseNextStatement();
        }

        pendingNextCloses = variables - 1;
        return null;
    }

    private int CountNextVariables()
    {
        if (Current.EndsStatement)
        {
            return 0;
        }

        var count = 1;
        var depth = 0;
        for (var i = 1; index + i < tokens.Count; i++)
        {
            var token = Peek(i);
            if (token.Kind == SyntaxKind.OpenParenToken)
            {
                depth++;
            }
            else if (token.Kind == SyntaxKind.CloseParenToken)
            {
                depth--;
            }
            else if (token.Kind == SyntaxKind.CommaToken && depth == 0)
            {
                count++;
            }

            if (token.EndsStatement || token.Kind == SyntaxKind.EndOfFileToken)
            {
                break;
            }
        }

        return count;
    }

    private NextStatementSyntax ParseNextStatement()
    {
        var nextKeyword = Consume();
        var variables = AtStatementStart
            ? new SeparatedSyntaxList<ExpressionSyntax>([])
            : ParseSeparatedList(ParsePostfixExpression);
        return new NextStatementSyntax(nextKeyword, variables);
    }

    private DoLoopBlockSyntax ParseDo()
    {
        var doKeyword = Consume();
        WhileOrUntilClauseSyntax? top = null;
        if (Current.Kind is SyntaxKind.WhileKeyword or SyntaxKind.UntilKeyword && !AtStatementStart)
        {
            var keyword = Consume();
            top = new WhileOrUntilClauseSyntax(keyword, ParseExpression());
        }

        var statements = ParseStatementList(SyntaxKind.DoLoopBlock);
        SyntaxToken loopKeyword;
        if (Current.Kind == SyntaxKind.LoopKeyword)
        {
            loopKeyword = Consume();
        }
        else
        {
            ReportMissingBlockEnd("Expected: Loop");
            loopKeyword = SyntaxToken.Missing(SyntaxKind.LoopKeyword, Current.Start);
        }
        WhileOrUntilClauseSyntax? bottom = null;
        if (!loopKeyword.IsMissing && Current.Kind is SyntaxKind.WhileKeyword or SyntaxKind.UntilKeyword && !AtStatementStart)
        {
            var keyword = Consume();
            bottom = new WhileOrUntilClauseSyntax(keyword, ParseExpression());
        }

        return new DoLoopBlockSyntax(doKeyword, top, statements, loopKeyword, bottom);
    }

    private WhileBlockSyntax ParseWhile()
    {
        var whileKeyword = Consume();
        var condition = ParseExpression();
        var statements = ParseStatementList(SyntaxKind.WhileBlock);
        SyntaxToken wendKeyword;
        if (Current.Kind == SyntaxKind.WendKeyword)
        {
            wendKeyword = Consume();
        }
        else
        {
            ReportMissingBlockEnd("Expected: Wend");
            wendKeyword = SyntaxToken.Missing(SyntaxKind.WendKeyword, Current.Start);
        }

        return new WhileBlockSyntax(whileKeyword, condition, statements, wendKeyword);
    }

    private WithBlockSyntax ParseWith()
    {
        var withKeyword = Consume();
        var expression = ParseExpression();
        var statements = ParseStatementList(SyntaxKind.WithBlock);
        return new WithBlockSyntax(withKeyword, expression, statements, ExpectEndBlock(SyntaxKind.WithKeyword));
    }

    private StatementSyntax ParseOn()
    {
        var onKeyword = Consume();
        if (IsWord(Current, "Error"))
        {
            var errorToken = Consume();
            if (Current.Kind == SyntaxKind.GoToKeyword)
            {
                var goToKeyword = Consume();
                return new OnErrorStatementSyntax(onKeyword, errorToken, goToKeyword, null, ParseLabelExpression());
            }

            if (Current.Kind == SyntaxKind.ResumeKeyword)
            {
                var resumeKeyword = Consume();
                return new OnErrorStatementSyntax(onKeyword, errorToken, resumeKeyword, Expect(SyntaxKind.NextKeyword), null);
            }

            Report("Expected: GoTo or Resume");
            return new OnErrorStatementSyntax(onKeyword, errorToken, SyntaxToken.Missing(SyntaxKind.GoToKeyword, Current.Start), null, null);
        }

        var expression = ParseExpression();
        SyntaxToken jumpKeyword;
        if (Current.Kind is SyntaxKind.GoToKeyword or SyntaxKind.GoSubKeyword)
        {
            jumpKeyword = Consume();
        }
        else
        {
            Report("Expected: GoTo or GoSub");
            jumpKeyword = SyntaxToken.Missing(SyntaxKind.GoToKeyword, Current.Start);
        }

        return new OnGoToStatementSyntax(onKeyword, expression, jumpKeyword, ParseSeparatedList(ParseLabelExpression));
    }

    private ResumeStatementSyntax ParseResume()
    {
        var resumeKeyword = Consume();
        if (AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword)
        {
            return new ResumeStatementSyntax(resumeKeyword, null, null);
        }

        if (Current.Kind == SyntaxKind.NextKeyword)
        {
            return new ResumeStatementSyntax(resumeKeyword, Consume(), null);
        }

        return new ResumeStatementSyntax(resumeKeyword, null, ParseLabelExpression());
    }

    /// <summary>A label reference: a name, a line number, or 0 / -1 after On Error GoTo.</summary>
    private ExpressionSyntax ParseLabelExpression()
    {
        if (Current.Kind == SyntaxKind.MinusToken)
        {
            var minus = Consume();
            return new UnaryExpressionSyntax(minus, new LiteralExpressionSyntax(Expect(SyntaxKind.IntegerLiteralToken)));
        }

        if (Current.Kind == SyntaxKind.IntegerLiteralToken)
        {
            return new LiteralExpressionSyntax(Consume());
        }

        return new IdentifierNameSyntax(ExpectLabel());
    }

    private ExitStatementSyntax ParseExit()
    {
        var exitKeyword = Consume();
        if (Current.Kind is SyntaxKind.SubKeyword or SyntaxKind.FunctionKeyword or SyntaxKind.DoKeyword or SyntaxKind.ForKeyword || IsWord(Current, "Property"))
        {
            return new ExitStatementSyntax(exitKeyword, Consume());
        }

        Report("Expected: Sub, Function, Property, Do, or For");
        return new ExitStatementSyntax(exitKeyword, SyntaxToken.Missing(SyntaxKind.SubKeyword, Current.Start));
    }

    private RaiseEventStatementSyntax ParseRaiseEvent()
    {
        var keyword = Consume();
        var name = ExpectIdentifier();
        var arguments = Current.Kind == SyntaxKind.OpenParenToken ? ParseParenthesizedArgumentList() : null;
        return new RaiseEventStatementSyntax(keyword, name, arguments);
    }

    private ReDimStatementSyntax ParseReDim()
    {
        var keyword = Consume();
        var preserve = Current.Kind == SyntaxKind.PreserveKeyword ? Consume() : null;
        return new ReDimStatementSyntax(keyword, preserve, ParseSeparatedList(() =>
        {
            // MS-VBAL 5.4.3.3 allows only a name, but Excel accepts any l-expression as the
            // target (This.Items, List(i).Elements); the last parenthesized group is the bounds
            // (docs/vba-quirks.md).
            ExpressionSyntax target = ParsePrimary();
            while (!AtStatementStart)
            {
                if (Current.Kind is SyntaxKind.DotToken or SyntaxKind.BangToken)
                {
                    var separator = Consume();
                    var name = new IdentifierNameSyntax(ExpectName());
                    target = separator.Kind == SyntaxKind.DotToken
                        ? new MemberAccessExpressionSyntax(target, separator, name)
                        : new DictionaryAccessExpressionSyntax(target, separator, name);
                }
                else if (Current.Kind == SyntaxKind.OpenParenToken && !IsFinalParenthesizedGroup())
                {
                    target = new IndexExpressionSyntax(target, ParseParenthesizedArgumentList());
                }
                else
                {
                    break;
                }
            }

            var bounds = ParseArrayBounds();
            var asClause = Current.Kind == SyntaxKind.AsKeyword ? ParseAsClause(allowArrayDesignator: false) : null;
            return new ReDimDeclaratorSyntax(target, bounds, asClause);
        }));
    }

    // File statements (MS-VBAL 5.4.5).

    private OpenStatementSyntax ParseOpen()
    {
        var openKeyword = Consume();
        var pathName = ParseExpression();
        SyntaxToken? forKeyword = null;
        SyntaxToken? mode = null;
        if (Current.Kind == SyntaxKind.ForKeyword)
        {
            forKeyword = Consume();
            if (Current.Kind == SyntaxKind.InputKeyword || IsWord(Current, "Append") || IsWord(Current, "Binary") || IsWord(Current, "Output") || IsWord(Current, "Random"))
            {
                mode = Consume();
            }
            else
            {
                Report("Expected: Append, Binary, Input, Output, or Random");
                mode = SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start);
            }
        }

        SyntaxToken? accessKeyword = null;
        var accessModes = new List<SyntaxToken>();
        if (IsWord(Current, "Access"))
        {
            accessKeyword = Consume();
            if (IsWord(Current, "Read"))
            {
                accessModes.Add(Consume());
            }

            if (Current.Kind == SyntaxKind.WriteKeyword)
            {
                accessModes.Add(Consume());
            }

            if (accessModes.Count == 0)
            {
                Report("Expected: Read or Write");
            }
        }

        var lockModes = new List<SyntaxToken>();
        if (Current.Kind == SyntaxKind.SharedKeyword)
        {
            lockModes.Add(Consume());
        }
        else if (Current.Kind == SyntaxKind.LockKeyword)
        {
            lockModes.Add(Consume());
            if (IsWord(Current, "Read"))
            {
                lockModes.Add(Consume());
            }

            if (Current.Kind == SyntaxKind.WriteKeyword)
            {
                lockModes.Add(Consume());
            }
        }

        var asKeyword = Expect(SyntaxKind.AsKeyword);
        var fileNumber = ParseFileNumber();
        SyntaxToken? lenKeyword = null;
        SyntaxToken? equals = null;
        ExpressionSyntax? recordLength = null;
        if (Current.Kind == SyntaxKind.LenKeyword && !AtStatementStart)
        {
            lenKeyword = Consume();
            equals = Expect(SyntaxKind.EqualsToken);
            recordLength = ParseExpression();
        }

        return new OpenStatementSyntax(openKeyword, pathName, forKeyword, mode, accessKeyword, new SyntaxTokenList(accessModes), new SyntaxTokenList(lockModes), asKeyword, fileNumber, lenKeyword, equals, recordLength);
    }

    private FileNumberSyntax ParseFileNumber()
    {
        var hash = Current.Kind == SyntaxKind.HashToken ? Consume() : null;
        return new FileNumberSyntax(hash, ParseExpression());
    }

    private CloseStatementSyntax ParseClose()
    {
        var keyword = Consume();
        var fileNumbers = AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword
            ? new SeparatedSyntaxList<FileNumberSyntax>([])
            : ParseSeparatedList(ParseFileNumber);
        return new CloseStatementSyntax(keyword, fileNumbers);
    }

    private SeekStatementSyntax ParseSeek()
    {
        var keyword = Consume();
        var fileNumber = ParseFileNumber();
        var comma = Expect(SyntaxKind.CommaToken);
        return new SeekStatementSyntax(keyword, fileNumber, comma, ParseExpression());
    }

    private LockStatementSyntax ParseLock()
    {
        var keyword = Consume();
        var fileNumber = ParseFileNumber();
        if (Current.Kind != SyntaxKind.CommaToken)
        {
            return new LockStatementSyntax(keyword, fileNumber, null, null);
        }

        var comma = Consume();
        var start = Current.Kind == SyntaxKind.ToKeyword ? null : ParseExpression();
        SyntaxToken? toKeyword = null;
        ExpressionSyntax? end = null;
        if (Current.Kind == SyntaxKind.ToKeyword)
        {
            toKeyword = Consume();
            end = ParseExpression();
        }

        return new LockStatementSyntax(keyword, fileNumber, comma, new RecordRangeSyntax(start, toKeyword, end));
    }

    private FileOutputStatementSyntax ParseFileOutput(SyntaxKind kind)
    {
        var keyword = Consume();
        var fileNumber = ParseFileNumber();
        var comma = Expect(SyntaxKind.CommaToken);
        return new FileOutputStatementSyntax(kind, keyword, fileNumber, comma, ParseOutputList());
    }

    /// <summary>output-list = *output-item; output-item = [output-clause] [char-position] (MS-VBAL 5.4.5.6).</summary>
    private SyntaxList<OutputItemSyntax> ParseOutputList()
    {
        var items = new List<OutputItemSyntax>();
        while (!AtStatementStart && Current.Kind != SyntaxKind.ElseKeyword)
        {
            var start = index;
            SyntaxNode? clause = null;
            if (Current.Kind == SyntaxKind.SpcKeyword)
            {
                var spc = Consume();
                var open = Expect(SyntaxKind.OpenParenToken);
                var count = ParseExpression();
                clause = new SpcClauseSyntax(spc, open, count, Expect(SyntaxKind.CloseParenToken));
            }
            else if (Current.Kind == SyntaxKind.TabKeyword)
            {
                var tab = Consume();
                if (Current.Kind == SyntaxKind.OpenParenToken)
                {
                    var open = Consume();
                    var column = ParseExpression();
                    clause = new TabClauseSyntax(tab, open, column, Expect(SyntaxKind.CloseParenToken));
                }
                else
                {
                    clause = new TabClauseSyntax(tab, null, null, null);
                }
            }
            else if (Current.Kind is not (SyntaxKind.SemicolonToken or SyntaxKind.CommaToken))
            {
                clause = ParseExpression();
            }

            var separator = Current.Kind is SyntaxKind.SemicolonToken or SyntaxKind.CommaToken ? Consume() : null;
            items.Add(new OutputItemSyntax(clause, separator));
            if (index == start)
            {
                break;
            }
        }

        return new SyntaxList<OutputItemSyntax>(items);
    }

    private InputStatementSyntax ParseInput()
    {
        var keyword = Consume();
        var fileNumber = ParseFileNumber();
        var comma = Expect(SyntaxKind.CommaToken);
        return new InputStatementSyntax(keyword, fileNumber, comma, ParseSeparatedList(ParsePostfixExpression));
    }

    private PutGetStatementSyntax ParsePutGet(SyntaxKind kind)
    {
        var keyword = Consume();
        var fileNumber = ParseFileNumber();
        var firstComma = Expect(SyntaxKind.CommaToken);
        var record = Current.Kind == SyntaxKind.CommaToken ? null : ParseExpression();
        var secondComma = Expect(SyntaxKind.CommaToken);
        return new PutGetStatementSyntax(kind, keyword, fileNumber, firstComma, record, secondComma, ParseExpression());
    }

    // Expressions (MS-VBAL 5.6), by precedence climbing over the table in 5.6.9.

    private ExpressionSyntax ParseExpression() => ParseBinary(1);

    private ExpressionSyntax ParseBinary(int minPrecedence) => ParseBinaryRest(ParseUnary(), minPrecedence);

    /// <summary>The operators and right operands that follow an already parsed left operand.</summary>
    private ExpressionSyntax ParseBinaryRest(ExpressionSyntax left, int minPrecedence)
    {
        while (!AtStatementStart)
        {
            var precedence = SyntaxFacts.GetBinaryOperatorPrecedence(Current.Kind);
            if (precedence == 0 || precedence < minPrecedence)
            {
                break;
            }

            var operatorToken = Consume();
            var right = ParseBinary(precedence + 1);
            left = new BinaryExpressionSyntax(left, operatorToken, right);
        }

        return left;
    }

    private ExpressionSyntax ParseUnary()
    {
        switch (Current.Kind)
        {
            case SyntaxKind.MinusToken or SyntaxKind.PlusToken:
                {
                    var operatorToken = Consume();
                    return new UnaryExpressionSyntax(operatorToken, ParseBinary(SyntaxFacts.UnaryNegationPrecedence + 1));
                }

            case SyntaxKind.NotKeyword:
                {
                    var operatorToken = Consume();
                    return new UnaryExpressionSyntax(operatorToken, ParseBinary(SyntaxFacts.NotPrecedence + 1));
                }

            case SyntaxKind.AddressOfKeyword:
                return new AddressOfExpressionSyntax(Consume(), ParsePostfixExpression());

            case SyntaxKind.TypeOfKeyword:
                {
                    var typeOfKeyword = Consume();
                    var expression = ParseBinary(SyntaxFacts.GetBinaryOperatorPrecedence(SyntaxKind.IsKeyword) + 1);
                    var isKeyword = Expect(SyntaxKind.IsKeyword);
                    return new TypeOfExpressionSyntax(typeOfKeyword, expression, isKeyword, ParseTypeName());
                }

            case SyntaxKind.NewKeyword:
                return new NewExpressionSyntax(Consume(), ParseTypeName());

            default:
                return ParsePostfixExpression();
        }
    }

    /// <summary>l-expression (MS-VBAL 5.6): a primary followed by member access, dictionary access, and index or call suffixes.</summary>
    private ExpressionSyntax ParsePostfixExpression()
    {
        var expression = ParsePrimary();
        while (!AtStatementStart)
        {
            if (Current.Kind == SyntaxKind.DotToken)
            {
                var dot = Consume();
                expression = new MemberAccessExpressionSyntax(expression, dot, new IdentifierNameSyntax(ExpectName()));
            }
            else if (Current.Kind == SyntaxKind.BangToken)
            {
                var bang = Consume();
                expression = new DictionaryAccessExpressionSyntax(expression, bang, new IdentifierNameSyntax(ExpectName()));
            }
            else if (Current.Kind == SyntaxKind.OpenParenToken)
            {
                expression = new IndexExpressionSyntax(expression, ParseParenthesizedArgumentList());
            }
            else
            {
                break;
            }
        }

        return expression;
    }

    private ExpressionSyntax ParsePrimary()
    {
        var token = Current;
        if (SyntaxFacts.IsLiteralToken(token.Kind) || SyntaxFacts.IsLiteralKeyword(token.Kind))
        {
            return new LiteralExpressionSyntax(Consume());
        }

        if (token.Kind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken || SyntaxFacts.IsExpressionKeyword(token.Kind))
        {
            return new IdentifierNameSyntax(Consume());
        }

        switch (token.Kind)
        {
            case SyntaxKind.OpenParenToken:
                {
                    var open = Consume();
                    var inner = ParseExpression();
                    return new ParenthesizedExpressionSyntax(open, inner, Expect(SyntaxKind.CloseParenToken));
                }

            case SyntaxKind.DotToken:
                {
                    // With-block member access (MS-VBAL 5.6.14).
                    var dot = Consume();
                    return new MemberAccessExpressionSyntax(null, dot, new IdentifierNameSyntax(ExpectName()));
                }

            case SyntaxKind.BangToken:
                {
                    var bang = Consume();
                    return new DictionaryAccessExpressionSyntax(null, bang, new IdentifierNameSyntax(ExpectName()));
                }

            case SyntaxKind.HashToken when Peek(1).Kind is SyntaxKind.IntegerLiteralToken or SyntaxKind.IdentifierToken:
                {
                    // A file number in argument position, as in Input(5, #1).
                    var hash = Consume();
                    return new UnaryExpressionSyntax(hash, ParsePrimary());
                }

            default:
                ReportMissing("Expected: expression");
                return new IdentifierNameSyntax(SyntaxToken.Missing(SyntaxKind.IdentifierToken, token.Start));
        }
    }

    private ArgumentListSyntax ParseParenthesizedArgumentList()
    {
        var open = Consume();
        var items = new List<SyntaxNodeOrToken>();
        if (Current.Kind != SyntaxKind.CloseParenToken)
        {
            while (true)
            {
                items.Add(ParseArgument(SyntaxKind.CloseParenToken));
                if (Current.Kind == SyntaxKind.CommaToken)
                {
                    items.Add(Consume());
                    continue;
                }

                break;
            }
        }

        return new ArgumentListSyntax(open, new SeparatedSyntaxList<ArgumentSyntax>(items), Expect(SyntaxKind.CloseParenToken));
    }

    /// <summary>Arguments of a call statement, written without parentheses and ended by the statement.</summary>
    private ArgumentListSyntax ParseBareArgumentList(ArgumentSyntax? first)
    {
        var items = new List<SyntaxNodeOrToken>();
        if (first is not null)
        {
            items.Add(first);
        }
        else
        {
            items.Add(ParseArgument(SyntaxKind.None));
        }

        while (Current.Kind == SyntaxKind.CommaToken)
        {
            items.Add(Consume());
            if (AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword)
            {
                items.Add(new ArgumentSyntax(null, null, null, null));
                break;
            }

            items.Add(ParseArgument(SyntaxKind.None));
        }

        return new ArgumentListSyntax(null, new SeparatedSyntaxList<ArgumentSyntax>(items), null);
    }

    private ArgumentSyntax ParseArgument(SyntaxKind closer)
    {
        if (Current.Kind == SyntaxKind.CommaToken || Current.Kind == closer || AtStatementStart || Current.Kind == SyntaxKind.ElseKeyword)
        {
            return new ArgumentSyntax(null, null, null, null);
        }

        SyntaxToken? name = null;
        SyntaxToken? colonEquals = null;
        if (Peek(1).Kind == SyntaxKind.ColonEqualsToken && (Current.Kind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken || Current.IsKeyword))
        {
            name = Consume();
            colonEquals = Consume();
        }

        var byVal = Current.Kind == SyntaxKind.ByValKeyword ? Consume() : null;
        return new ArgumentSyntax(name, colonEquals, byVal, ParseExpression());
    }

    private static bool CanStartExpression(SyntaxToken token) =>
        SyntaxFacts.IsLiteralToken(token.Kind) || SyntaxFacts.IsLiteralKeyword(token.Kind) || SyntaxFacts.IsExpressionKeyword(token.Kind)
        || token.Kind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken or SyntaxKind.OpenParenToken or SyntaxKind.DotToken
            or SyntaxKind.BangToken or SyntaxKind.MinusToken or SyntaxKind.PlusToken or SyntaxKind.NotKeyword or SyntaxKind.NewKeyword
            or SyntaxKind.TypeOfKeyword or SyntaxKind.AddressOfKeyword;

    private static bool CanStartArgument(SyntaxToken token) =>
        CanStartExpression(token) || token.Kind is SyntaxKind.CommaToken or SyntaxKind.ByValKeyword or SyntaxKind.HashToken;

    // Token helpers.

    private SyntaxToken Consume()
    {
        var token = Current;
        if (!AtEnd)
        {
            index++;
        }

        return token;
    }

    private SyntaxToken Expect(SyntaxKind kind)
    {
        if (Current.Kind == kind)
        {
            return Consume();
        }

        ReportMissing("Expected: " + SyntaxFacts.GetText(kind));
        return SyntaxToken.Missing(kind, Current.Start);
    }

    /// <summary>An identifier, typed name, or foreign name in a declaring position.</summary>
    private SyntaxToken ExpectIdentifier()
    {
        if (Current.Kind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken)
        {
            return Consume();
        }

        ReportMissing("Expected: identifier");
        return SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start);
    }

    /// <summary>unrestricted-name (MS-VBAL 5.6.10): an identifier, a foreign name, or any keyword, as after "." or "!".</summary>
    private SyntaxToken ExpectName()
    {
        if (Current.Kind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken || Current.IsKeyword)
        {
            return Consume();
        }

        ReportMissing("Expected: identifier");
        return SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start);
    }

    private SyntaxToken ExpectLabel()
    {
        if (Current.Kind is SyntaxKind.IdentifierToken or SyntaxKind.IntegerLiteralToken)
        {
            return Consume();
        }

        ReportMissing("Expected: label");
        return SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start);
    }

    private SyntaxToken ExpectWord(string word)
    {
        if (IsWord(Current, word))
        {
            return Consume();
        }

        ReportMissing("Expected: " + word);
        return SyntaxToken.Missing(SyntaxKind.IdentifierToken, Current.Start);
    }

    private EndBlockStatementSyntax ExpectEndBlock(SyntaxKind blockKeyword)
    {
        if (Current.Kind == SyntaxKind.EndKeyword && Peek(1).Kind == blockKeyword)
        {
            return new EndBlockStatementSyntax(Consume(), Consume());
        }

        ReportMissingBlockEnd("Expected: End " + SyntaxFacts.GetText(blockKeyword));
        return new EndBlockStatementSyntax(SyntaxToken.Missing(SyntaxKind.EndKeyword, Current.Start), SyntaxToken.Missing(blockKeyword, Current.Start));
    }

    private static bool IsWord(SyntaxToken token, string word) =>
        token.Kind == SyntaxKind.IdentifierToken && token.TypeSuffix == '\0' && string.Equals(token.NameValue, word, StringComparison.OrdinalIgnoreCase);

    private static bool IsMidWord(SyntaxToken token) =>
        token.Kind == SyntaxKind.IdentifierToken && token.TypeSuffix is '\0' or '$'
        && (string.Equals(token.NameValue, "Mid", StringComparison.OrdinalIgnoreCase) || string.Equals(token.NameValue, "MidB", StringComparison.OrdinalIgnoreCase));

    private SeparatedSyntaxList<TNode> ParseSeparatedList<TNode>(Func<TNode> parseItem)
        where TNode : SyntaxNode
    {
        var items = new List<SyntaxNodeOrToken>();
        while (true)
        {
            var start = index;
            items.Add(parseItem());
            if (Current.Kind != SyntaxKind.CommaToken || index == start)
            {
                break;
            }

            items.Add(Consume());
        }

        return new SeparatedSyntaxList<TNode>(items);
    }

    private void Report(string message) => ReportAt(Current, message);

    private void ReportAt(SyntaxToken token, string message) => ReportAtOffset(token.Start, message);

    /// <summary>
    /// A token was expected where the statement already ended: report at the end of the previous
    /// token (the end of the line, as the VBE does), otherwise at the token found instead.
    /// </summary>
    private void ReportMissing(string message) =>
        ReportAtOffset(AtStatementStart && Previous is { } previous ? previous.End : Current.Start, message);

    /// <summary>A block end was expected: report at the token found instead, or at the end of the file.</summary>
    private void ReportMissingBlockEnd(string message) =>
        ReportAtOffset(AtEnd && Previous is { } previous ? previous.End : Current.Start, message);

    private void ReportAtOffset(int offset, string message)
    {
        if (offset == lastErrorPosition)
        {
            return;
        }

        lastErrorPosition = offset;
        var (line, column) = source.GetLinePosition(offset);
        diagnostics.Add(new Diagnostic(DiagnosticIds.SyntaxError, DiagnosticSeverity.Error, message, filePath, line, column));
    }
}
