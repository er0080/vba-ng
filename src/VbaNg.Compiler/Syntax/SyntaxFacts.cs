using System.Collections.Frozen;

namespace VbaNg.Compiler.Syntax;

/// <summary>Tables derived from MS-VBAL section 3.3: keywords, their categories, and token classification.</summary>
public static class SyntaxFacts
{
    private static readonly FrozenDictionary<string, SyntaxKind> Keywords = BuildKeywords();
    private static readonly FrozenDictionary<SyntaxKind, string> KeywordText = Keywords.ToFrozenDictionary(p => p.Value, p => p.Key);

    /// <summary>The keyword kind for a word, or <see cref="SyntaxKind.None"/> (MS-VBAL 3.3.5.2, case-insensitive).</summary>
    public static SyntaxKind GetKeywordKind(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return Keywords.TryGetValue(text, out var kind) ? kind : SyntaxKind.None;
    }

    /// <summary>Canonical spelling of a keyword kind, as written in MS-VBAL.</summary>
    public static string GetText(SyntaxKind kind) =>
        KeywordText.TryGetValue(kind, out var text) ? text : kind switch
        {
            SyntaxKind.OpenParenToken => "(",
            SyntaxKind.CloseParenToken => ")",
            SyntaxKind.CommaToken => ",",
            SyntaxKind.SemicolonToken => ";",
            SyntaxKind.DotToken => ".",
            SyntaxKind.BangToken => "!",
            SyntaxKind.HashToken => "#",
            SyntaxKind.EqualsToken => "=",
            SyntaxKind.LessThanToken => "<",
            SyntaxKind.GreaterThanToken => ">",
            SyntaxKind.LessThanEqualsToken => "<=",
            SyntaxKind.GreaterThanEqualsToken => ">=",
            SyntaxKind.LessThanGreaterThanToken => "<>",
            SyntaxKind.PlusToken => "+",
            SyntaxKind.MinusToken => "-",
            SyntaxKind.AsteriskToken => "*",
            SyntaxKind.SlashToken => "/",
            SyntaxKind.BackslashToken => "\\",
            SyntaxKind.CaretToken => "^",
            SyntaxKind.AmpersandToken => "&",
            SyntaxKind.ColonToken => ":",
            SyntaxKind.ColonEqualsToken => ":=",
            SyntaxKind.EndOfFileToken => "end of file",
            SyntaxKind.IdentifierToken => "identifier",
            _ => kind.ToString(),
        };

    public static bool IsKeyword(SyntaxKind kind) => kind >= SyntaxKind.CallKeyword && kind <= SyntaxKind.DefDecKeyword;

    public static bool IsTrivia(SyntaxKind kind) => kind >= SyntaxKind.WhitespaceTrivia && kind <= SyntaxKind.SkippedTokensTrivia;

    public static bool IsToken(SyntaxKind kind) => kind >= SyntaxKind.EndOfFileToken && kind <= SyntaxKind.DefDecKeyword;

    /// <summary>MS-VBAL 3.3.5.3 type-suffix.</summary>
    public static bool IsTypeSuffix(char c) => c is '%' or '&' or '^' or '!' or '#' or '@' or '$';

    /// <summary>Keywords that name a declared type in an As clause (reserved-type-identifier plus the future-reserved Decimal).</summary>
    public static bool IsBuiltinTypeKeyword(SyntaxKind kind) => kind is
        SyntaxKind.BooleanKeyword or SyntaxKind.ByteKeyword or SyntaxKind.CurrencyKeyword or SyntaxKind.DateKeyword or
        SyntaxKind.DoubleKeyword or SyntaxKind.IntegerKeyword or SyntaxKind.LongKeyword or SyntaxKind.LongLongKeyword or
        SyntaxKind.LongPtrKeyword or SyntaxKind.SingleKeyword or SyntaxKind.StringKeyword or SyntaxKind.VariantKeyword or
        SyntaxKind.DecimalKeyword or SyntaxKind.AnyKeyword;

    /// <summary>
    /// Keywords that may start a primary expression as if they were program-defined names:
    /// reserved-name and special-form (MS-VBAL 3.3.5.2), the literal identifiers, and Seek, which
    /// VBA also exposes as a function.
    /// </summary>
    public static bool IsExpressionKeyword(SyntaxKind kind) => kind is
        SyntaxKind.AbsKeyword or SyntaxKind.CBoolKeyword or SyntaxKind.CByteKeyword or SyntaxKind.CCurKeyword or
        SyntaxKind.CDateKeyword or SyntaxKind.CDblKeyword or SyntaxKind.CDecKeyword or SyntaxKind.CIntKeyword or
        SyntaxKind.CLngKeyword or SyntaxKind.CLngLngKeyword or SyntaxKind.CLngPtrKeyword or SyntaxKind.CSngKeyword or
        SyntaxKind.CStrKeyword or SyntaxKind.CVarKeyword or SyntaxKind.CVErrKeyword or SyntaxKind.DateKeyword or
        SyntaxKind.DebugKeyword or SyntaxKind.DoEventsKeyword or SyntaxKind.FixKeyword or SyntaxKind.IntKeyword or
        SyntaxKind.LenKeyword or SyntaxKind.LenBKeyword or SyntaxKind.MeKeyword or SyntaxKind.PSetKeyword or
        SyntaxKind.ScaleKeyword or SyntaxKind.SgnKeyword or SyntaxKind.StringKeyword or
        SyntaxKind.ArrayKeyword or SyntaxKind.CircleKeyword or SyntaxKind.InputKeyword or SyntaxKind.InputBKeyword or
        SyntaxKind.LBoundKeyword or SyntaxKind.UBoundKeyword or SyntaxKind.SeekKeyword or
        SyntaxKind.TrueKeyword or SyntaxKind.FalseKeyword or SyntaxKind.NothingKeyword or SyntaxKind.EmptyKeyword or SyntaxKind.NullKeyword;

    public static bool IsLiteralKeyword(SyntaxKind kind) => kind is
        SyntaxKind.TrueKeyword or SyntaxKind.FalseKeyword or SyntaxKind.NothingKeyword or SyntaxKind.EmptyKeyword or SyntaxKind.NullKeyword;

    public static bool IsLiteralToken(SyntaxKind kind) => kind is
        SyntaxKind.IntegerLiteralToken or SyntaxKind.FloatLiteralToken or SyntaxKind.StringLiteralToken or SyntaxKind.DateLiteralToken;

    /// <summary>Binary operator precedence per MS-VBAL 5.6.9 (higher binds tighter); 0 for non-operators.</summary>
    public static int GetBinaryOperatorPrecedence(SyntaxKind kind) => kind switch
    {
        SyntaxKind.CaretToken => 14,
        // 13 is unary negation.
        SyntaxKind.AsteriskToken or SyntaxKind.SlashToken => 12,
        SyntaxKind.BackslashToken => 11,
        SyntaxKind.ModKeyword => 10,
        SyntaxKind.PlusToken or SyntaxKind.MinusToken => 9,
        SyntaxKind.AmpersandToken => 8,
        SyntaxKind.EqualsToken or SyntaxKind.LessThanGreaterThanToken or SyntaxKind.LessThanToken or SyntaxKind.GreaterThanToken or
        SyntaxKind.LessThanEqualsToken or SyntaxKind.GreaterThanEqualsToken or SyntaxKind.LikeKeyword or SyntaxKind.IsKeyword => 7,
        // 6 is Not.
        SyntaxKind.AndKeyword => 5,
        SyntaxKind.OrKeyword => 4,
        SyntaxKind.XorKeyword => 3,
        SyntaxKind.EqvKeyword => 2,
        SyntaxKind.ImpKeyword => 1,
        _ => 0,
    };

    public const int UnaryNegationPrecedence = 13;
    public const int NotPrecedence = 6;

    private static FrozenDictionary<string, SyntaxKind> BuildKeywords()
    {
        var map = new Dictionary<string, SyntaxKind>(StringComparer.OrdinalIgnoreCase);
        for (var kind = SyntaxKind.CallKeyword; kind <= SyntaxKind.DefDecKeyword; kind++)
        {
            var name = kind.ToString();
            map[name[..^"Keyword".Length]] = kind;
        }

        return map.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
    }
}
