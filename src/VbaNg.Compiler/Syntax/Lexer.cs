using System.Globalization;
using System.Text;

namespace VbaNg.Compiler.Syntax;

/// <summary>The tokens of one module, ending with an end-of-file token, plus lexical diagnostics.</summary>
public sealed record LexResult(IReadOnlyList<SyntaxToken> Tokens, IReadOnlyList<Diagnostic> Diagnostics);

/// <summary>
/// Tokenizer for MS-VBAL section 3: physical and logical lines (3.2), separators and comments
/// (3.3.1), number, date, string, and identifier tokens (3.3.2 through 3.3.5). Conditional
/// compilation directives are left as tokens for the <see cref="Preprocessor"/>. Never throws:
/// unexpected input becomes a <see cref="SyntaxKind.BadToken"/> with a VBA0001 diagnostic.
/// </summary>
public sealed class Lexer
{
    private readonly string text;
    private readonly string filePath;
    private readonly SourceText source;
    private readonly List<Diagnostic> diagnostics = [];
    private readonly List<SyntaxToken> tokens = [];
    private int position;

    /// <summary>True when the next token is the first on its logical line.</summary>
    private bool atLineStart = true;

    /// <summary>True when the next token starts a statement (after a line terminator or a colon).</summary>
    private bool atStatementStart = true;

    private SyntaxKind previousKind = SyntaxKind.None;

    private Lexer(SourceText source, string filePath)
    {
        this.source = source;
        text = source.Text;
        this.filePath = filePath;
    }

    public static LexResult Tokenize(SourceText source, string filePath)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(filePath);
        var lexer = new Lexer(source, filePath);
        lexer.Run();
        return new LexResult(lexer.tokens, lexer.diagnostics);
    }

    private char Peek(int offset = 0) => position + offset < text.Length ? text[position + offset] : '\0';

    private bool AtEnd => position >= text.Length;

    private void Run()
    {
        if (StartsHeader())
        {
            LexHeader();
        }

        while (true)
        {
            var leading = ScanLeadingTrivia();
            var wasLineStart = atLineStart;
            var start = position;
            var token = ScanToken();
            var isLabelCandidate = wasLineStart && token.Kind is SyntaxKind.IdentifierToken or SyntaxKind.IntegerLiteralToken;
            var trailing = token.Kind == SyntaxKind.EndOfFileToken ? [] : ScanTrailingTrivia(isLabelCandidate);
            token = new SyntaxToken(token.Kind, token.Text, start, leading, trailing, token.Value);
            tokens.Add(token);
            previousKind = token.Kind;
            if (token.Kind == SyntaxKind.EndOfFileToken)
            {
                return;
            }

            if (isLabelCandidate && Peek() == ':' && Peek(1) != '=')
            {
                // The label's colon (MS-VBAL 5.4.1.1); a statement may follow it on the same line.
                var colonStart = position;
                position++;
                var colonTrailing = ScanTrailingTrivia(labelCandidate: false);
                var colon = new SyntaxToken(SyntaxKind.ColonToken, ":", colonStart, null, colonTrailing);
                tokens.Add(colon);
                previousKind = colon.Kind;
                atStatementStart = true;
            }
        }
    }

    // MS-VBAL 4.2: "the version and all text between BEGIN and END at the start of the file is
    // not part of the module body and is not required to conform to the VBA grammar."
    private bool StartsHeader() =>
        text.Length > 7
        && string.Compare(text, 0, "VERSION", 0, 7, StringComparison.OrdinalIgnoreCase) == 0
        && IsWhitespace(text[7]);

    private void LexHeader()
    {
        EmitHeaderLine();
        var depth = 0;
        if (!LineStartsWithWord("Begin") && !LineStartsWithWord("BeginProperty"))
        {
            return;
        }

        while (!AtEnd)
        {
            if (LineStartsWithWord("Begin") || LineStartsWithWord("BeginProperty"))
            {
                depth++;
            }
            else if (LineStartsWithWord("End") || LineStartsWithWord("EndProperty"))
            {
                depth--;
            }

            EmitHeaderLine();
            if (depth <= 0)
            {
                return;
            }
        }
    }

    private bool LineStartsWithWord(string word)
    {
        var i = position;
        while (i < text.Length && IsWhitespace(text[i]))
        {
            i++;
        }

        if (i + word.Length > text.Length || string.Compare(text, i, word, 0, word.Length, StringComparison.OrdinalIgnoreCase) != 0)
        {
            return false;
        }

        var after = i + word.Length;
        return after >= text.Length || !IsIdentifierPart(text[after]);
    }

    private void EmitHeaderLine()
    {
        var start = position;
        while (!AtEnd && SourceText.LineTerminatorLength(text, position) == 0)
        {
            position++;
        }

        var lineText = text[start..position];
        var trailing = new List<SyntaxTrivia>();
        var terminator = SourceText.LineTerminatorLength(text, position);
        if (terminator > 0)
        {
            trailing.Add(new SyntaxTrivia(SyntaxKind.EndOfLineTrivia, text.Substring(position, terminator)));
            position += terminator;
        }

        var token = new SyntaxToken(SyntaxKind.HeaderLineToken, lineText, start, null, trailing);
        tokens.Add(token);
        previousKind = token.Kind;
        atLineStart = true;
        atStatementStart = true;
    }

    // MS-VBAL 3.2.2 WSC: tab, U+0019, space, U+3000, and Unicode class Zs.
    private static bool IsWhitespace(char c) =>
        c is ' ' or '\t' or '\x19' or '\x3000' || (c > '\x7F' && CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.SpaceSeparator);

    private static bool IsIdentifierStart(char c) => char.IsLetter(c);

    private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

    private static bool IsDigit(char c) => c is >= '0' and <= '9';

    private static bool IsHexDigit(char c) => IsDigit(c) || c is >= 'A' and <= 'F' || c is >= 'a' and <= 'f';

    private static bool IsOctalDigit(char c) => c is >= '0' and <= '7';

    private List<SyntaxTrivia> ScanLeadingTrivia()
    {
        var trivia = new List<SyntaxTrivia>();
        while (!AtEnd)
        {
            var c = Peek();
            if (IsWhitespace(c))
            {
                trivia.Add(ScanWhitespace());
            }
            else if (c == '\'')
            {
                trivia.Add(ScanComment());
            }
            else if (atStatementStart && IsRemAt(position))
            {
                trivia.Add(ScanComment());
            }
            else if (SourceText.LineTerminatorLength(text, position) is var terminator && terminator > 0)
            {
                trivia.Add(new SyntaxTrivia(SyntaxKind.EndOfLineTrivia, text.Substring(position, terminator)));
                position += terminator;
                atLineStart = true;
                atStatementStart = true;
            }
            else if (c == '_' && IsContinuationAt(position))
            {
                trivia.Add(ScanContinuation());
            }
            else if (c == ':' && Peek(1) != '=')
            {
                trivia.Add(new SyntaxTrivia(SyntaxKind.ColonTrivia, ":"));
                position++;
                atStatementStart = true;
                atLineStart = false;
            }
            else
            {
                break;
            }
        }

        return trivia;
    }

    private List<SyntaxTrivia> ScanTrailingTrivia(bool labelCandidate)
    {
        var trivia = new List<SyntaxTrivia>();
        atLineStart = false;
        atStatementStart = false;
        while (!AtEnd)
        {
            var c = Peek();
            if (IsWhitespace(c))
            {
                trivia.Add(ScanWhitespace());
            }
            else if (c == '\'')
            {
                trivia.Add(ScanComment());
            }
            else if (c == '_' && IsContinuationAt(position))
            {
                trivia.Add(ScanContinuation());
            }
            else if (SourceText.LineTerminatorLength(text, position) is var terminator && terminator > 0)
            {
                trivia.Add(new SyntaxTrivia(SyntaxKind.EndOfLineTrivia, text.Substring(position, terminator)));
                position += terminator;
                atLineStart = true;
                atStatementStart = true;
                break;
            }
            else if (c == ':' && Peek(1) != '=')
            {
                if (labelCandidate)
                {
                    // MS-VBAL 5.4.1.1: a name or line number first on the line, followed by ":", is a label; the colon is its token.
                    break;
                }

                trivia.Add(new SyntaxTrivia(SyntaxKind.ColonTrivia, ":"));
                position++;
                atStatementStart = true;
                break;
            }
            else
            {
                break;
            }
        }

        return trivia;
    }

    private SyntaxTrivia ScanWhitespace()
    {
        var start = position;
        while (!AtEnd && IsWhitespace(Peek()))
        {
            position++;
        }

        return new SyntaxTrivia(SyntaxKind.WhitespaceTrivia, text[start..position]);
    }

    // MS-VBAL 3.3.1: comment-body = *(line-continuation / non-line-termination-character) LINE-END.
    // A line continuation inside a comment continues the comment.
    private SyntaxTrivia ScanComment()
    {
        var start = position;
        while (!AtEnd)
        {
            if (Peek() == '_' && IsContinuationAt(position) && position > start && IsWhitespace(text[position - 1]))
            {
                SkipContinuation();
                continue;
            }

            if (SourceText.LineTerminatorLength(text, position) > 0)
            {
                break;
            }

            position++;
        }

        return new SyntaxTrivia(SyntaxKind.CommentTrivia, text[start..position]);
    }

    private bool IsRemAt(int index) =>
        index + 3 <= text.Length
        && string.Compare(text, index, "Rem", 0, 3, StringComparison.OrdinalIgnoreCase) == 0
        && (index + 3 == text.Length || !IsIdentifierPart(text[index + 3]) && !SyntaxFacts.IsTypeSuffix(text[index + 3]));

    // MS-VBAL 3.2.2: line-continuation = 1*WSC "_" line-terminator. Whitespace before the underscore
    // is required, as the VBE requires it (a line starting with "_" is rejected, docs/vba-quirks.md);
    // trailing whitespace after the underscore is tolerated, as the VBE tolerates it.
    private bool IsContinuationAt(int index)
    {
        if (index >= text.Length || text[index] != '_' || index == 0 || !IsWhitespace(text[index - 1]))
        {
            return false;
        }

        var i = index + 1;
        while (i < text.Length && IsWhitespace(text[i]))
        {
            i++;
        }

        return i >= text.Length || SourceText.LineTerminatorLength(text, i) > 0;
    }

    private SyntaxTrivia ScanContinuation()
    {
        var start = position;
        SkipContinuation();
        return new SyntaxTrivia(SyntaxKind.LineContinuationTrivia, text[start..position]);
    }

    private void SkipContinuation()
    {
        position++;
        while (!AtEnd && IsWhitespace(Peek()))
        {
            position++;
        }

        position += SourceText.LineTerminatorLength(text, position);
    }

    private SyntaxToken ScanToken()
    {
        var start = position;
        if (AtEnd)
        {
            return new SyntaxToken(SyntaxKind.EndOfFileToken, string.Empty, start);
        }

        var c = Peek();
        if (IsIdentifierStart(c))
        {
            return ScanWord();
        }

        if (IsDigit(c) || (c == '.' && IsDigit(Peek(1))))
        {
            return ScanNumber();
        }

        // "&H" and "&O" always start a literal ("Err.Raise &H80040201" passes one as a bare
        // argument); a bare "&" followed by octal digits is a literal only in operand position.
        if (c == '&' && ((Peek(1) is 'H' or 'h' or 'O' or 'o' && IsHexDigit(Peek(2))) || (Peek(1) is >= '0' and <= '7')))
        {
            return ScanNumber();
        }

        if (c == '"')
        {
            return ScanString();
        }

        if (c == '#')
        {
            return ScanHash();
        }

        if (c == '[')
        {
            return ScanForeignName();
        }

        var kind = c switch
        {
            '(' => SyntaxKind.OpenParenToken,
            ')' => SyntaxKind.CloseParenToken,
            ',' => SyntaxKind.CommaToken,
            ';' => SyntaxKind.SemicolonToken,
            '.' => SyntaxKind.DotToken,
            '!' => SyntaxKind.BangToken,
            '+' => SyntaxKind.PlusToken,
            '-' => SyntaxKind.MinusToken,
            '*' => SyntaxKind.AsteriskToken,
            '/' => SyntaxKind.SlashToken,
            '\\' => SyntaxKind.BackslashToken,
            '^' => SyntaxKind.CaretToken,
            '&' => SyntaxKind.AmpersandToken,
            _ => SyntaxKind.None,
        };

        if (kind != SyntaxKind.None)
        {
            position++;
            return new SyntaxToken(kind, c.ToString(), start);
        }

        switch (c)
        {
            case '=':
                // VBA accepts "=<" and "=>" as spellings of "<=" and ">=".
                if (Peek(1) == '<')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.LessThanEqualsToken, "=<", start);
                }

                if (Peek(1) == '>')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.GreaterThanEqualsToken, "=>", start);
                }

                position++;
                return new SyntaxToken(SyntaxKind.EqualsToken, "=", start);

            case '<':
                if (Peek(1) == '=')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.LessThanEqualsToken, "<=", start);
                }

                if (Peek(1) == '>')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.LessThanGreaterThanToken, "<>", start);
                }

                position++;
                return new SyntaxToken(SyntaxKind.LessThanToken, "<", start);

            case '>':
                if (Peek(1) == '=')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.GreaterThanEqualsToken, ">=", start);
                }

                if (Peek(1) == '<')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.LessThanGreaterThanToken, "><", start);
                }

                position++;
                return new SyntaxToken(SyntaxKind.GreaterThanToken, ">", start);

            case ':':
                if (Peek(1) == '=')
                {
                    position += 2;
                    return new SyntaxToken(SyntaxKind.ColonEqualsToken, ":=", start);
                }

                position++;
                return new SyntaxToken(SyntaxKind.ColonToken, ":", start);

            default:
                position++;
                Report(start, string.Create(CultureInfo.InvariantCulture, $"Unexpected character '{c}'."));
                return new SyntaxToken(SyntaxKind.BadToken, c.ToString(), start);
        }
    }

    private bool PreviousIsOperand() =>
        !atStatementStart
        && (previousKind is SyntaxKind.IdentifierToken or SyntaxKind.ForeignNameToken or SyntaxKind.CloseParenToken
            || SyntaxFacts.IsLiteralToken(previousKind)
            || SyntaxFacts.IsExpressionKeyword(previousKind));

    // MS-VBAL 3.3.5: identifiers; 3.3.5.2 reserved identifiers; 3.3.5.3 typed names.
    private SyntaxToken ScanWord()
    {
        var start = position;
        while (!AtEnd && IsIdentifierPart(Peek()))
        {
            position++;
        }

        var name = text[start..position];
        if (!AtEnd && IsTypeSuffixHere(Peek()))
        {
            position++;
            return new SyntaxToken(SyntaxKind.IdentifierToken, text[start..position], start, value: name);
        }

        var keyword = SyntaxFacts.GetKeywordKind(name);
        return keyword == SyntaxKind.None
            ? new SyntaxToken(SyntaxKind.IdentifierToken, name, start, value: name)
            : new SyntaxToken(keyword, name, start);
    }

    /// <summary>
    /// A type-suffix character directly after a name is the suffix: the VBE rejects "a&amp;b",
    /// "a&amp;\"x\"", "a^b", "a^(2)", "x^2", and "2^3" alike, reading the suffix and then a stray
    /// operand. The one exception is "!" before a name, "rs!Field", which is dictionary access
    /// (Lexer golden, VBE probes, docs/vba-quirks.md).
    /// </summary>
    private bool IsTypeSuffixHere(char c)
    {
        if (!SyntaxFacts.IsTypeSuffix(c))
        {
            return false;
        }

        var next = Peek(1);
        return c switch
        {
            '!' => !(IsIdentifierStart(next) || next is '_' or '['),
            '&' => true,
            '^' => true,
            _ => true,
        };
    }

    // MS-VBAL 3.3.2 number tokens.
    private SyntaxToken ScanNumber()
    {
        var start = position;
        if (Peek() == '&')
        {
            position++;
            var radixChar = Peek();
            if (radixChar is 'H' or 'h')
            {
                position++;
                var digitStart = position;
                while (!AtEnd && IsHexDigit(Peek()))
                {
                    position++;
                }

                if (position == digitStart)
                {
                    Report(start, "Hexadecimal literal has no digits.");
                }
            }
            else
            {
                if (radixChar is 'O' or 'o')
                {
                    position++;
                }

                var digitStart = position;
                while (!AtEnd && IsOctalDigit(Peek()))
                {
                    position++;
                }

                if (position == digitStart)
                {
                    Report(start, "Octal literal has no digits.");
                }
            }

            ScanIntegerSuffix();
            return new SyntaxToken(SyntaxKind.IntegerLiteralToken, text[start..position], start);
        }

        while (!AtEnd && IsDigit(Peek()))
        {
            position++;
        }

        var isFloat = false;
        if (Peek() == '.')
        {
            isFloat = true;
            position++;
            while (!AtEnd && IsDigit(Peek()))
            {
                position++;
            }
        }

        if (Peek() is 'E' or 'e' or 'D' or 'd')
        {
            var exponentDigits = Peek(1) is '+' or '-' ? 2 : 1;
            if (IsDigit(Peek(exponentDigits)))
            {
                isFloat = true;
                position += exponentDigits;
                while (!AtEnd && IsDigit(Peek()))
                {
                    position++;
                }
            }
        }

        if (Peek() is '!' or '#' or '@')
        {
            isFloat = true;
            position++;
        }
        else if (!isFloat)
        {
            ScanIntegerSuffix();
        }

        return new SyntaxToken(isFloat ? SyntaxKind.FloatLiteralToken : SyntaxKind.IntegerLiteralToken, text[start..position], start);
    }

    private void ScanIntegerSuffix()
    {
        if (Peek() == '%' || (Peek() is '&' or '^' && IsTypeSuffixHere(Peek())))
        {
            position++;
        }
    }

    // MS-VBAL 3.3.4 string tokens: a doubled quote stands for one quote. VBA closes a string
    // left open at the end of the line (docs/vba-quirks.md).
    private SyntaxToken ScanString()
    {
        var start = position;
        position++;
        var value = new StringBuilder();
        while (!AtEnd && SourceText.LineTerminatorLength(text, position) == 0)
        {
            if (Peek() == '"')
            {
                if (Peek(1) == '"')
                {
                    value.Append('"');
                    position += 2;
                    continue;
                }

                position++;
                return new SyntaxToken(SyntaxKind.StringLiteralToken, text[start..position], start, value: value.ToString());
            }

            value.Append(Peek());
            position++;
        }

        return new SyntaxToken(SyntaxKind.StringLiteralToken, text[start..position], start, value: value.ToString());
    }

    // "#" starts a conditional compilation directive at line start (MS-VBAL 3.4), a date literal
    // (3.3.3) when a matching "#" closes a date-or-time on the same line, and otherwise stands
    // alone as the file-number marker of the file statements (5.4.5).
    private SyntaxToken ScanHash()
    {
        var start = position;
        if (atLineStart && IsDirectiveWordAt(position + 1))
        {
            position++;
            return new SyntaxToken(SyntaxKind.HashToken, "#", start);
        }

        var close = position + 1;
        while (close < text.Length && text[close] != '#' && SourceText.LineTerminatorLength(text, close) == 0)
        {
            close++;
        }

        if (close < text.Length && text[close] == '#' && DateLiteralSyntax.IsDateOrTime(text.AsSpan(position + 1, close - position - 1)))
        {
            position = close + 1;
            return new SyntaxToken(SyntaxKind.DateLiteralToken, text[start..position], start);
        }

        position++;
        return new SyntaxToken(SyntaxKind.HashToken, "#", start);
    }

    private bool IsDirectiveWordAt(int index)
    {
        while (index < text.Length && IsWhitespace(text[index]))
        {
            index++;
        }

        var end = index;
        while (end < text.Length && IsIdentifierPart(text[end]))
        {
            end++;
        }

        var word = text[index..end];
        return word.Equals("If", StringComparison.OrdinalIgnoreCase)
            || word.Equals("ElseIf", StringComparison.OrdinalIgnoreCase)
            || word.Equals("Else", StringComparison.OrdinalIgnoreCase)
            || word.Equals("End", StringComparison.OrdinalIgnoreCase)
            || word.Equals("EndIf", StringComparison.OrdinalIgnoreCase)
            || word.Equals("Const", StringComparison.OrdinalIgnoreCase);
    }

    // MS-VBAL 3.3.5.3: FOREIGN-NAME = "[" foreign-identifier "]".
    private SyntaxToken ScanForeignName()
    {
        var start = position;
        position++;
        while (!AtEnd && Peek() != ']' && SourceText.LineTerminatorLength(text, position) == 0)
        {
            position++;
        }

        if (Peek() == ']')
        {
            position++;
            return new SyntaxToken(SyntaxKind.ForeignNameToken, text[start..position], start, value: text[(start + 1)..(position - 1)]);
        }

        Report(start, "']' expected to close the bracketed name.");
        return new SyntaxToken(SyntaxKind.ForeignNameToken, text[start..position], start, value: text[(start + 1)..position]);
    }

    private void Report(int offset, string message)
    {
        var (line, column) = source.GetLinePosition(offset);
        diagnostics.Add(new Diagnostic(DiagnosticIds.SyntaxError, DiagnosticSeverity.Error, message, filePath, line, column));
    }
}
