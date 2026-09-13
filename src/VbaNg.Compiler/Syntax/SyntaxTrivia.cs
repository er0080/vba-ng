namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Source text that carries no syntactic meaning: whitespace, comments, line continuations,
/// line terminators, statement-separating colons, conditional compilation directives, and
/// excluded source (MS-VBAL 3.2.2, 3.3.1, 3.4). Every character of a file belongs to exactly one
/// token or trivia, which is what makes the tree round-trip byte for byte.
/// </summary>
public class SyntaxTrivia
{
    public SyntaxTrivia(SyntaxKind kind, string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        Kind = kind;
        Text = text;
    }

    public SyntaxKind Kind { get; }

    public string Text { get; }

    public int Width => Text.Length;

    public bool IsEndOfLine => Kind == SyntaxKind.EndOfLineTrivia;

    /// <summary>True for trivia that ends a statement: a line terminator or a colon separator.</summary>
    public bool IsStatementTerminator => Kind is SyntaxKind.EndOfLineTrivia or SyntaxKind.ColonTrivia;

    public override string ToString() => Text;
}

/// <summary>Tokens the parser skipped during error recovery, kept so the tree still round-trips.</summary>
public sealed class SkippedTokensTrivia : SyntaxTrivia
{
    public SkippedTokensTrivia(IReadOnlyList<SyntaxToken> tokens)
        : base(SyntaxKind.SkippedTokensTrivia, string.Concat((tokens ?? throw new ArgumentNullException(nameof(tokens))).Select(t => t.ToFullString())))
    {
        Tokens = tokens;
    }

    public IReadOnlyList<SyntaxToken> Tokens { get; }
}
