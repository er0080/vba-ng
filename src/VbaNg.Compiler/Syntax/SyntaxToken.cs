using System.Diagnostics;

namespace VbaNg.Compiler.Syntax;

/// <summary>
/// A lexical token with its surrounding trivia. Leading trivia is everything between the previous
/// token's trailing trivia and this token; trailing trivia runs from the token to the end of its
/// line, including the line terminator or the colon that ends the statement (MS-VBAL 3.3:
/// "Lexical tokens encompass any white space characters that immediately precede them"; the
/// trailing side is this implementation's choice so that a statement owns its own line).
/// </summary>
[DebuggerDisplay("{Kind} {Text}")]
public sealed class SyntaxToken
{
    private static readonly IReadOnlyList<SyntaxTrivia> NoTrivia = [];

    public SyntaxToken(
        SyntaxKind kind,
        string text,
        int start,
        IReadOnlyList<SyntaxTrivia>? leadingTrivia = null,
        IReadOnlyList<SyntaxTrivia>? trailingTrivia = null,
        object? value = null,
        bool isMissing = false)
    {
        ArgumentNullException.ThrowIfNull(text);
        Kind = kind;
        Text = text;
        Start = start;
        LeadingTrivia = leadingTrivia ?? NoTrivia;
        TrailingTrivia = trailingTrivia ?? NoTrivia;
        Value = value;
        IsMissing = isMissing;
    }

    public SyntaxKind Kind { get; }

    /// <summary>The token's own text, exactly as written; empty for missing tokens.</summary>
    public string Text { get; }

    /// <summary>Offset of <see cref="Text"/> in the source.</summary>
    public int Start { get; }

    public int End => Start + Text.Length;

    public int FullStart => Start - LeadingTrivia.Sum(t => t.Width);

    public int FullEnd => End + TrailingTrivia.Sum(t => t.Width);

    public IReadOnlyList<SyntaxTrivia> LeadingTrivia { get; }

    public IReadOnlyList<SyntaxTrivia> TrailingTrivia { get; }

    /// <summary>
    /// Decoded value: the name without type suffix for identifiers, the unescaped text for string
    /// literals, the text inside the brackets for foreign names; null otherwise.
    /// </summary>
    public object? Value { get; }

    /// <summary>A token the parser expected but did not find. Missing tokens have no text.</summary>
    public bool IsMissing { get; }

    public bool IsKeyword => SyntaxFacts.IsKeyword(Kind);

    /// <summary>The type suffix character of a typed name (MS-VBAL 3.3.5.3), or '\0'.</summary>
    public char TypeSuffix =>
        Kind == SyntaxKind.IdentifierToken && Text.Length > 0 && SyntaxFacts.IsTypeSuffix(Text[^1]) ? Text[^1] : '\0';

    /// <summary>The name value of an identifier, typed name, foreign name, or keyword used as a name (MS-VBAL 3.3.5.1).</summary>
    public string NameValue => Value as string ?? Text;

    /// <summary>True if this token's trailing trivia ends the statement it belongs to.</summary>
    public bool EndsStatement => TrailingTrivia.Any(t => t.IsStatementTerminator);

    /// <summary>True if this token's trailing trivia ends the logical line.</summary>
    public bool EndsLine => TrailingTrivia.Any(t => t.IsEndOfLine);

    public SyntaxToken WithLeadingTrivia(IReadOnlyList<SyntaxTrivia> trivia) =>
        new(Kind, Text, Start, trivia, TrailingTrivia, Value, IsMissing);

    public SyntaxToken WithTrailingTrivia(IReadOnlyList<SyntaxTrivia> trivia) =>
        new(Kind, Text, Start, LeadingTrivia, trivia, Value, IsMissing);

    public static SyntaxToken Missing(SyntaxKind kind, int position) =>
        new(kind, string.Empty, position, isMissing: true);

    public void WriteTo(TextWriter writer)
    {
        ArgumentNullException.ThrowIfNull(writer);
        foreach (var trivia in LeadingTrivia)
        {
            writer.Write(trivia.Text);
        }

        writer.Write(Text);
        foreach (var trivia in TrailingTrivia)
        {
            writer.Write(trivia.Text);
        }
    }

    /// <summary>The token with its trivia, as written in the source.</summary>
    public string ToFullString()
    {
        using var writer = new StringWriter();
        WriteTo(writer);
        return writer.ToString();
    }

    public override string ToString() => Text;
}
