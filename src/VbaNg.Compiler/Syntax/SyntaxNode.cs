using System.Runtime.CompilerServices;

namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Base of every node in the full-fidelity tree. Children are enumerated in source order, so
/// writing them out reproduces the original text exactly (ARCHITECTURE.md section 4, step 2).
/// </summary>
public abstract class SyntaxNode
{
    protected SyntaxNode(SyntaxKind kind)
    {
        Kind = kind;
    }

    public SyntaxKind Kind { get; }

    /// <summary>Child nodes and tokens in source order. Null optional children are omitted.</summary>
    public abstract IEnumerable<SyntaxNodeOrToken> ChildNodesAndTokens();

    public IEnumerable<SyntaxNode> ChildNodes() =>
        ChildNodesAndTokens().Where(c => c.IsNode).Select(c => c.AsNode()!);

    public IEnumerable<SyntaxToken> DescendantTokens()
    {
        foreach (var child in ChildNodesAndTokens())
        {
            if (child.IsToken)
            {
                yield return child.AsToken()!;
            }
            else
            {
                foreach (var token in child.AsNode()!.DescendantTokens())
                {
                    yield return token;
                }
            }
        }
    }

    public IEnumerable<SyntaxNode> DescendantNodes()
    {
        foreach (var child in ChildNodes())
        {
            yield return child;
            foreach (var descendant in child.DescendantNodes())
            {
                yield return descendant;
            }
        }
    }

    public SyntaxToken? FirstToken() => DescendantTokens().FirstOrDefault();

    public SyntaxToken? LastToken()
    {
        SyntaxToken? last = null;
        foreach (var token in DescendantTokens())
        {
            last = token;
        }

        return last;
    }

    /// <summary>Offset of the first token's text, ignoring leading trivia.</summary>
    public int Start => FirstToken()?.Start ?? 0;

    public int End => LastToken()?.End ?? 0;

    public bool ContainsMissingTokens => DescendantTokens().Any(t => t.IsMissing);

    public void WriteTo(TextWriter writer)
    {
        ArgumentNullException.ThrowIfNull(writer);
        foreach (var child in ChildNodesAndTokens())
        {
            if (child.IsToken)
            {
                child.AsToken()!.WriteTo(writer);
            }
            else
            {
                child.AsNode()!.WriteTo(writer);
            }
        }
    }

    /// <summary>The node with all its trivia, exactly as written in the source.</summary>
    public string ToFullString()
    {
        using var writer = new StringWriter();
        WriteTo(writer);
        return writer.ToString();
    }

    public override string ToString() => ToFullString();

    /// <summary>Helper for <see cref="ChildNodesAndTokens"/> implementations: yields the non-null items in order.</summary>
    protected static IEnumerable<SyntaxNodeOrToken> Children(params SyntaxNodeOrToken[] items)
    {
        foreach (var item in items)
        {
            if (!item.IsDefault)
            {
                yield return item;
            }
        }
    }

    /// <summary>Helper for nodes whose children include lists: flattens nodes, tokens, and lists in order.</summary>
    protected static IEnumerable<SyntaxNodeOrToken> Children(params IEnumerable<SyntaxNodeOrToken>[] groups)
    {
        foreach (var group in groups)
        {
            foreach (var item in group)
            {
                if (!item.IsDefault)
                {
                    yield return item;
                }
            }
        }
    }

    /// <summary>Wraps a single optional node or token as a one-item child group.</summary>
    protected static IEnumerable<SyntaxNodeOrToken> One(SyntaxNodeOrToken item)
    {
        if (!item.IsDefault)
        {
            yield return item;
        }
    }
}

/// <summary>Either a node or a token, for enumerating children in source order.</summary>
public readonly struct SyntaxNodeOrToken : IEquatable<SyntaxNodeOrToken>
{
    private readonly object? item;

    private SyntaxNodeOrToken(object? item)
    {
        this.item = item;
    }

    public bool IsDefault => item is null;

    public bool IsNode => item is SyntaxNode;

    public bool IsToken => item is SyntaxToken;

    public SyntaxNode? AsNode() => item as SyntaxNode;

    public SyntaxToken? AsToken() => item as SyntaxToken;

    public static implicit operator SyntaxNodeOrToken(SyntaxNode? node) => new(node);

    public static implicit operator SyntaxNodeOrToken(SyntaxToken? token) => new(token);

    public static SyntaxNodeOrToken FromNode(SyntaxNode? node) => new(node);

    public static SyntaxNodeOrToken FromToken(SyntaxToken? token) => new(token);

    public bool Equals(SyntaxNodeOrToken other) => ReferenceEquals(item, other.item);

    public override bool Equals(object? obj) => obj is SyntaxNodeOrToken other && Equals(other);

    public override int GetHashCode() => item is null ? 0 : RuntimeHelpers.GetHashCode(item);

    public static bool operator ==(SyntaxNodeOrToken left, SyntaxNodeOrToken right) => left.Equals(right);

    public static bool operator !=(SyntaxNodeOrToken left, SyntaxNodeOrToken right) => !left.Equals(right);

    public override string ToString() => item?.ToString() ?? string.Empty;
}
