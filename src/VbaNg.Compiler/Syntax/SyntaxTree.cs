namespace VbaNg.Compiler.Syntax;

/// <summary>A parsed module: the full-fidelity root, its source, and every lexical and syntactic diagnostic.</summary>
public sealed class SyntaxTree
{
    private SyntaxTree(string filePath, SourceText text, ModuleSyntax root, IReadOnlyList<Diagnostic> diagnostics)
    {
        FilePath = filePath;
        Text = text;
        Root = root;
        Diagnostics = diagnostics;
    }

    public string FilePath { get; }

    public SourceText Text { get; }

    public ModuleSyntax Root { get; }

    public IReadOnlyList<Diagnostic> Diagnostics { get; }

    public bool HasErrors => Diagnostics.Any(d => d.IsError);

    /// <summary>Lexes, preprocesses (MS-VBAL 3.4), and parses one module.</summary>
    public static SyntaxTree Parse(SourceFile file, ParseOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(file);
        return Parse(file.Text, file.Path, options);
    }

    public static SyntaxTree Parse(string text, string filePath, ParseOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(filePath);
        options ??= ParseOptions.Default;

        var source = new SourceText(text);
        var diagnostics = new List<Diagnostic>();
        var lexed = Lexer.Tokenize(source, filePath);
        var tokens = Preprocessor.Process(lexed.Tokens, lexed.Diagnostics, source, filePath, options, diagnostics);
        var root = new Parser(tokens, source, filePath, diagnostics).ParseModule();
        return new SyntaxTree(filePath, source, root, diagnostics);
    }

    /// <summary>1-based line and column of a token's text.</summary>
    public (int Line, int Column) GetLinePosition(SyntaxToken token)
    {
        ArgumentNullException.ThrowIfNull(token);
        return Text.GetLinePosition(token.Start);
    }
}
