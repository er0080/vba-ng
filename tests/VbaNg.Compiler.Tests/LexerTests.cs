using VbaNg.Compiler.Syntax;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>Lexer tests for MS-VBAL section 3: tokens, trivia, and byte-exact reconstruction.</summary>
public sealed class LexerTests
{
    private const string FilePath = @"C:\test\Module1.bas";

    private static LexResult Lex(string text) => Lexer.Tokenize(new SourceText(text), FilePath);

    /// <summary>Tokens without the end-of-file token.</summary>
    private static List<SyntaxToken> Tokens(string text) => Lex(text).Tokens.SkipLast(1).ToList();

    private static List<SyntaxKind> Kinds(string text) => Tokens(text).Select(t => t.Kind).ToList();

    private static string RoundTrip(string text) => string.Concat(Lex(text).Tokens.Select(t => t.ToFullString()));

    [Theory]
    [InlineData("x = 1\r\n")]
    [InlineData("  Dim a As Long ' comment\r\n\r\n    ' only a comment\r\nRem another\r\n")]
    [InlineData("Total = a + _\r\n        b + _ \r\n        c\r\n")]
    [InlineData("Foo: x = 1: y = 2\r\n10 Print #1, x; y, Spc(3)\r\n")]
    [InlineData("#If Win64 Then\r\n    Declare PtrSafe Sub S Lib \"k\" ()\r\n#Else\r\n    Declare Sub S Lib \"k\" ()\r\n#End If\r\n")]
    [InlineData("d = #1/1/2000# + #12:00:00 PM#: s = \"a\"\"b\" & [A1] & &HFF& & &O17 & 1.5E-3# & x$\r\n")]
    [InlineData("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\n")]
    [InlineData("no trailing newline")]
    [InlineData("mixed\nendings\r\nhere\rdone")]
    [InlineData("' comment continued _\r\nstill comment\r\nx = 1\r\n")]
    [InlineData("s = \"unterminated\r\nx = 2\r\n")]
    [InlineData("a = b ? c\r\n")]
    public void Tokens_ReconstructTheSourceExactly(string text)
    {
        Assert.Equal(text, RoundTrip(text));
    }

    [Fact]
    public void Keywords_AreCaseInsensitive()
    {
        var kinds = Kinds("dim DIM Dim End end If then");

        Assert.Equal(
            [SyntaxKind.DimKeyword, SyntaxKind.DimKeyword, SyntaxKind.DimKeyword, SyntaxKind.EndKeyword, SyntaxKind.EndKeyword, SyntaxKind.IfKeyword, SyntaxKind.ThenKeyword],
            kinds);
    }

    [Fact]
    public void ReservedIdentifiers_AreKeywordsAndOthersAreIdentifiers()
    {
        // Name, Error, Mid, Property, Object, Explicit are quoted in the grammar but not reserved (MS-VBAL 3.3.5.2).
        // Local is the reverse: absent from the grammar, reserved by the VBE (docs/vba-quirks.md).
        var tokens = Tokens("Name Error Mid Property Object Explicit Me Attribute Decimal CDecl Local");

        Assert.All(tokens.Take(6), t => Assert.Equal(SyntaxKind.IdentifierToken, t.Kind));
        Assert.Equal(SyntaxKind.MeKeyword, tokens[6].Kind);
        Assert.Equal(SyntaxKind.AttributeKeyword, tokens[7].Kind);
        Assert.Equal(SyntaxKind.DecimalKeyword, tokens[8].Kind);
        Assert.Equal(SyntaxKind.CDeclKeyword, tokens[9].Kind);
        Assert.Equal(SyntaxKind.LocalKeyword, tokens[10].Kind);
    }

    [Fact]
    public void Identifiers_KeepCasingAndExposeNameValue()
    {
        var token = Tokens("myVar")[0];

        Assert.Equal("myVar", token.Text);
        Assert.Equal("myVar", token.NameValue);
        Assert.Equal('\0', token.TypeSuffix);
    }

    [Theory]
    [InlineData("x$", '$')]
    [InlineData("count&", '&')]
    [InlineData("n%", '%')]
    [InlineData("f!", '!')]
    [InlineData("d#", '#')]
    [InlineData("c@", '@')]
    [InlineData("big^", '^')]
    [InlineData("Left$", '$')]
    [InlineData("String$", '$')]
    public void TypedNames_AreOneIdentifierTokenWithSuffix(string text, char suffix)
    {
        var token = Assert.Single(Tokens(text));

        Assert.Equal(SyntaxKind.IdentifierToken, token.Kind);
        Assert.Equal(text, token.Text);
        Assert.Equal(suffix, token.TypeSuffix);
        Assert.Equal(text[..^1], token.NameValue);
    }

    [Theory]
    [InlineData("a&b", SyntaxKind.IdentifierToken, SyntaxKind.IdentifierToken)]
    [InlineData("a&\"x\"", SyntaxKind.IdentifierToken, SyntaxKind.StringLiteralToken)]
    [InlineData("a^b", SyntaxKind.IdentifierToken, SyntaxKind.IdentifierToken)]
    [InlineData("a^(2)", SyntaxKind.IdentifierToken, SyntaxKind.OpenParenToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.CloseParenToken)]
    [InlineData("rs!Field", SyntaxKind.IdentifierToken, SyntaxKind.BangToken, SyntaxKind.IdentifierToken)]
    [InlineData("rs![Field Name]", SyntaxKind.IdentifierToken, SyntaxKind.BangToken, SyntaxKind.ForeignNameToken)]
    [InlineData("x&(5)", SyntaxKind.IdentifierToken, SyntaxKind.OpenParenToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.CloseParenToken)]
    public void TypeSuffixCharacters_AfterANameAreSuffixesExceptTheBang(string text, params SyntaxKind[] expected)
    {
        Assert.Equal(expected, Kinds(text));
    }

    [Theory]
    [InlineData("0", SyntaxKind.IntegerLiteralToken)]
    [InlineData("32768", SyntaxKind.IntegerLiteralToken)]
    [InlineData("10&", SyntaxKind.IntegerLiteralToken)]
    [InlineData("10%", SyntaxKind.IntegerLiteralToken)]
    [InlineData("10^", SyntaxKind.IntegerLiteralToken)]
    [InlineData("&HFF", SyntaxKind.IntegerLiteralToken)]
    [InlineData("&hff&", SyntaxKind.IntegerLiteralToken)]
    [InlineData("&O17", SyntaxKind.IntegerLiteralToken)]
    [InlineData("&17", SyntaxKind.IntegerLiteralToken)]
    [InlineData("1.5", SyntaxKind.FloatLiteralToken)]
    [InlineData(".5", SyntaxKind.FloatLiteralToken)]
    [InlineData("1.", SyntaxKind.FloatLiteralToken)]
    [InlineData("1E3", SyntaxKind.FloatLiteralToken)]
    [InlineData("1.5e-3", SyntaxKind.FloatLiteralToken)]
    [InlineData("2D+2", SyntaxKind.FloatLiteralToken)]
    [InlineData("1#", SyntaxKind.FloatLiteralToken)]
    [InlineData("1!", SyntaxKind.FloatLiteralToken)]
    [InlineData("1@", SyntaxKind.FloatLiteralToken)]
    [InlineData("1.5@", SyntaxKind.FloatLiteralToken)]
    public void Numbers_LexAsOneToken(string text, SyntaxKind kind)
    {
        var token = Assert.Single(Tokens(text));

        Assert.Equal(kind, token.Kind);
        Assert.Equal(text, token.Text);
    }

    [Fact]
    public void Ampersand_BeforeOctalDigitsAfterAnOperand_IsStillTheOctalLiteral()
    {
        Assert.Equal(
            [SyntaxKind.CloseParenToken, SyntaxKind.IntegerLiteralToken],
            Kinds(") &7"));
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IntegerLiteralToken],
            Kinds("x = &7"));
    }

    [Fact]
    public void HexAndOctalPrefixes_AlwaysStartALiteral()
    {
        // Err.Raise &H80040201 passes the literal as a bare argument.
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.DotToken, SyntaxKind.IdentifierToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.CommaToken, SyntaxKind.StringLiteralToken],
            Kinds("Err.Raise &H80040201, \"src\""));
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.IntegerLiteralToken],
            Kinds("x &H1"));
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.AmpersandToken, SyntaxKind.IdentifierToken],
            Kinds("x &Hz"));
    }

    [Fact]
    public void Strings_UnescapeDoubledQuotes()
    {
        var token = Assert.Single(Tokens("\"say \"\"hi\"\"\""));

        Assert.Equal(SyntaxKind.StringLiteralToken, token.Kind);
        Assert.Equal("say \"hi\"", token.Value);
    }

    [Fact]
    public void Strings_LeftOpenAtEndOfLine_CloseThere()
    {
        var result = Lex("s = \"open\r\nx = 1\r\n");

        Assert.Empty(result.Diagnostics);
        var literal = result.Tokens[2];
        Assert.Equal(SyntaxKind.StringLiteralToken, literal.Kind);
        Assert.Equal("\"open", literal.Text);
        Assert.Equal("open", literal.Value);
        Assert.True(literal.EndsLine);
    }

    [Theory]
    [InlineData("#1/1/2000#")]
    [InlineData("#12:00:00 PM#")]
    [InlineData("#1/1/2000 12:00 PM#")]
    [InlineData("#Jan 1, 2000#")]
    [InlineData("#1 january 2000#")]
    [InlineData("#2000-01-01#")]
    [InlineData("#3 PM#")]
    [InlineData("#10.30#")]
    public void Dates_LexAsOneToken(string text)
    {
        var token = Assert.Single(Tokens(text));

        Assert.Equal(SyntaxKind.DateLiteralToken, token.Kind);
    }

    [Fact]
    public void Hash_BeforeAFileNumber_IsAToken()
    {
        Assert.Equal(
            [SyntaxKind.PrintKeyword, SyntaxKind.HashToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.CommaToken, SyntaxKind.IdentifierToken],
            Kinds("Print #1, x"));
        Assert.Equal(
            [SyntaxKind.GetKeyword, SyntaxKind.HashToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.CommaToken, SyntaxKind.CommaToken, SyntaxKind.IdentifierToken],
            Kinds("Get #1, , x#"));
    }

    [Fact]
    public void Hash_AtLineStartBeforeADirectiveWord_IsAToken()
    {
        Assert.Equal(
            [SyntaxKind.HashToken, SyntaxKind.IfKeyword, SyntaxKind.IdentifierToken, SyntaxKind.ThenKeyword],
            Kinds("#If Win64 Then"));
        Assert.Equal(
            [SyntaxKind.HashToken, SyntaxKind.EndKeyword, SyntaxKind.IfKeyword],
            Kinds("#End If"));
        Assert.Equal(
            [SyntaxKind.HashToken, SyntaxKind.ConstKeyword, SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IntegerLiteralToken],
            Kinds("#Const DEBUGGING = 1"));
    }

    [Fact]
    public void ForeignNames_KeepTheBracketedText()
    {
        var token = Assert.Single(Tokens("[Sheet1!A1:B2]"));

        Assert.Equal(SyntaxKind.ForeignNameToken, token.Kind);
        Assert.Equal("Sheet1!A1:B2", token.NameValue);
    }

    [Fact]
    public void Continuation_AtTheStartOfALine_IsNotAContinuation()
    {
        // The VBE rejects a line that starts with the underscore (docs/vba-quirks.md, Lexer continuation).
        var result = Lex("x = 1 + _\r\n_\r\n2\r\n");

        Assert.NotEmpty(result.Diagnostics);
        Assert.Contains(result.Diagnostics, d => d.Line == 2);
    }

    [Fact]
    public void ForeignName_LeftOpen_ReportsVba0001()
    {
        var result = Lex("x = [oops\r\n");

        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(DiagnosticIds.SyntaxError, diagnostic.Id);
        Assert.Equal(1, diagnostic.Line);
        Assert.Equal(5, diagnostic.Column);
    }

    [Fact]
    public void Comments_AreTrailingTriviaOfTheLastTokenOnTheLine()
    {
        var tokens = Tokens("x = 1 ' note\r\ny = 2");

        var one = tokens[2];
        Assert.Equal(
            [SyntaxKind.WhitespaceTrivia, SyntaxKind.CommentTrivia, SyntaxKind.EndOfLineTrivia],
            one.TrailingTrivia.Select(t => t.Kind));
        Assert.Equal("' note", one.TrailingTrivia[1].Text);
        Assert.True(one.EndsStatement);
        Assert.True(one.EndsLine);
    }

    [Fact]
    public void RemComments_OnlyAtStatementStart()
    {
        var tokens = Tokens("Rem first\r\nx = 1: Rem second\r\nCall Rem\r\n");

        Assert.Equal(SyntaxKind.IdentifierToken, tokens[0].Kind);
        Assert.Contains(tokens[0].LeadingTrivia, t => t.Kind == SyntaxKind.CommentTrivia && t.Text == "Rem first");
        Assert.Equal(SyntaxKind.CallKeyword, tokens[3].Kind);
        Assert.Contains(tokens[3].LeadingTrivia, t => t.Kind == SyntaxKind.CommentTrivia && t.Text == "Rem second");
        Assert.Equal(SyntaxKind.RemKeyword, tokens[4].Kind);
    }

    [Fact]
    public void CommentContinuation_SwallowsTheNextLine()
    {
        var tokens = Tokens("' first _\r\nsecond\r\nx = 1");

        Assert.Equal(SyntaxKind.IdentifierToken, tokens[0].Kind);
        Assert.Equal("x", tokens[0].Text);
        var comment = Assert.Single(tokens[0].LeadingTrivia, t => t.Kind == SyntaxKind.CommentTrivia);
        Assert.Equal("' first _\r\nsecond", comment.Text);
    }

    [Fact]
    public void LineContinuation_JoinsLogicalLines()
    {
        var tokens = Tokens("x = 1 + _\r\n    2\r\n");

        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IntegerLiteralToken, SyntaxKind.PlusToken, SyntaxKind.IntegerLiteralToken],
            tokens.Select(t => t.Kind));
        var plus = tokens[3];
        Assert.Contains(plus.TrailingTrivia, t => t.Kind == SyntaxKind.LineContinuationTrivia && t.Text == "_\r\n");
        Assert.False(plus.EndsStatement);
        Assert.True(tokens[4].EndsLine);
    }

    [Fact]
    public void Colon_SeparatesStatementsAsTrivia()
    {
        var tokens = Tokens("x = 1: y = 2");

        Assert.Equal(5 + 1, tokens.Count);
        Assert.Contains(tokens[2].TrailingTrivia, t => t.Kind == SyntaxKind.ColonTrivia);
        Assert.True(tokens[2].EndsStatement);
        Assert.False(tokens[2].EndsLine);
    }

    [Fact]
    public void Label_FirstOnTheLine_KeepsItsColonAsAToken()
    {
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.ColonToken, SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IntegerLiteralToken],
            Kinds("Retry: x = 1"));
        Assert.Equal(
            [SyntaxKind.IntegerLiteralToken, SyntaxKind.ColonToken],
            Kinds("100:"));
        Assert.Equal(
            [SyntaxKind.IntegerLiteralToken, SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IntegerLiteralToken],
            Kinds("100 x = 1"));
    }

    [Fact]
    public void Colon_AfterAKeywordOrMidLine_IsNotALabel()
    {
        Assert.Equal([SyntaxKind.CaseKeyword, SyntaxKind.IntegerLiteralToken, SyntaxKind.IdentifierToken], Kinds("Case 1: x"));
        Assert.Equal([SyntaxKind.ElseKeyword, SyntaxKind.IdentifierToken], Kinds("Else: x"));
        Assert.Equal([SyntaxKind.IdentifierToken, SyntaxKind.EqualsToken, SyntaxKind.IdentifierToken, SyntaxKind.IdentifierToken], Kinds("x = y: z"));
    }

    [Fact]
    public void NamedArgument_ColonEquals_IsOneToken()
    {
        Assert.Equal(
            [SyntaxKind.IdentifierToken, SyntaxKind.IdentifierToken, SyntaxKind.ColonEqualsToken, SyntaxKind.IntegerLiteralToken],
            Kinds("Foo Bar:=1"));
    }

    [Fact]
    public void AlternativeOperatorSpellings_KeepTheirText()
    {
        var tokens = Tokens("a =< b => c >< d <= e >= f <> g");

        Assert.Equal(SyntaxKind.LessThanEqualsToken, tokens[1].Kind);
        Assert.Equal("=<", tokens[1].Text);
        Assert.Equal(SyntaxKind.GreaterThanEqualsToken, tokens[3].Kind);
        Assert.Equal("=>", tokens[3].Text);
        Assert.Equal(SyntaxKind.LessThanGreaterThanToken, tokens[5].Kind);
        Assert.Equal("><", tokens[5].Text);
        Assert.Equal(SyntaxKind.LessThanEqualsToken, tokens[7].Kind);
        Assert.Equal(SyntaxKind.GreaterThanEqualsToken, tokens[9].Kind);
        Assert.Equal(SyntaxKind.LessThanGreaterThanToken, tokens[11].Kind);
    }

    [Fact]
    public void ClassModuleHeader_IsRawLinesUntilEnd()
    {
        var tokens = Tokens("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\n");

        Assert.Equal(
            [SyntaxKind.HeaderLineToken, SyntaxKind.HeaderLineToken, SyntaxKind.HeaderLineToken, SyntaxKind.HeaderLineToken, SyntaxKind.AttributeKeyword],
            tokens.Take(5).Select(t => t.Kind));
        Assert.Equal("  MultiUse = -1  'True", tokens[2].Text);
    }

    [Fact]
    public void FormHeader_HandlesNestedBeginEnd()
    {
        var text = "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 \r\n   Caption         =   \"Form\"\r\n   BeginProperty Font \r\n      Name            =   \"Tahoma\"\r\n   EndProperty\r\nEnd\r\nAttribute VB_Name = \"UserForm1\"\r\n";
        var tokens = Tokens(text);

        Assert.Equal(7, tokens.Count(t => t.Kind == SyntaxKind.HeaderLineToken));
        Assert.Equal(SyntaxKind.AttributeKeyword, tokens[7].Kind);
    }

    [Fact]
    public void ProceduralModule_HasNoHeaderTokens()
    {
        var tokens = Tokens("Attribute VB_Name = \"M\"\r\nOption Explicit\r\n");

        Assert.DoesNotContain(tokens, t => t.Kind == SyntaxKind.HeaderLineToken);
    }

    [Fact]
    public void UnexpectedCharacter_ReportsVba0001WithPosition()
    {
        var result = Lex("x = 1\r\n  y = ?\r\n");

        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(FilePath + "(2,7): error VBA0001: Unexpected character '?'.", diagnostic.ToString());
        Assert.Contains(result.Tokens, t => t.Kind == SyntaxKind.BadToken);
    }

    [Fact]
    public void Positions_AreOffsetsOfTheTokenText()
    {
        var tokens = Tokens("  x = 1");

        Assert.Equal(2, tokens[0].Start);
        Assert.Equal(0, tokens[0].FullStart);
        Assert.Equal(4, tokens[1].Start);
        Assert.Equal(6, tokens[2].Start);
    }
}
