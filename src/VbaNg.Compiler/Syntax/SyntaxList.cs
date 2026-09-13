using System.Collections;

namespace VbaNg.Compiler.Syntax;

/// <summary>An immutable list of nodes, itself enumerable as children in source order.</summary>
public sealed class SyntaxList<TNode> : IReadOnlyList<TNode>
    where TNode : SyntaxNode
{
    private readonly TNode[] nodes;

    public SyntaxList(IEnumerable<TNode> nodes)
    {
        ArgumentNullException.ThrowIfNull(nodes);
        this.nodes = nodes.ToArray();
    }

    public int Count => nodes.Length;

    public TNode this[int index] => nodes[index];

    public IEnumerator<TNode> GetEnumerator() => ((IEnumerable<TNode>)nodes).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerable<SyntaxNodeOrToken> AsChildren() => nodes.Select(n => (SyntaxNodeOrToken)n);
}

/// <summary>An immutable list of tokens, for example the modifiers of a declaration.</summary>
public sealed class SyntaxTokenList : IReadOnlyList<SyntaxToken>
{
    private readonly SyntaxToken[] tokens;

    public SyntaxTokenList(IEnumerable<SyntaxToken> tokens)
    {
        ArgumentNullException.ThrowIfNull(tokens);
        this.tokens = tokens.ToArray();
    }

    public int Count => tokens.Length;

    public SyntaxToken this[int index] => tokens[index];

    public bool Any(SyntaxKind kind) => tokens.Any(t => t.Kind == kind);

    public IEnumerator<SyntaxToken> GetEnumerator() => ((IEnumerable<SyntaxToken>)tokens).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IEnumerable<SyntaxNodeOrToken> AsChildren() => tokens.Select(t => (SyntaxNodeOrToken)t);
}

/// <summary>
/// Nodes separated by tokens (usually commas), stored interleaved so that both the nodes and the
/// separators are children in source order.
/// </summary>
public sealed class SeparatedSyntaxList<TNode> : IReadOnlyList<TNode>
    where TNode : SyntaxNode
{
    private readonly SyntaxNodeOrToken[] items;

    /// <summary>Items alternate node, separator, node, ...; a trailing separator is allowed (an omitted last element).</summary>
    public SeparatedSyntaxList(IEnumerable<SyntaxNodeOrToken> items)
    {
        ArgumentNullException.ThrowIfNull(items);
        this.items = items.ToArray();
    }

    /// <summary>Number of nodes.</summary>
    public int Count => items.Count(i => i.IsNode);

    public int SeparatorCount => items.Length - Count;

    public TNode this[int index] => Nodes.ElementAt(index);

    public IEnumerable<TNode> Nodes => items.Where(i => i.IsNode).Select(i => (TNode)i.AsNode()!);

    public IEnumerable<SyntaxToken> Separators => items.Where(i => i.IsToken).Select(i => i.AsToken()!);

    public IEnumerable<SyntaxNodeOrToken> AsChildren() => items;

    public IEnumerator<TNode> GetEnumerator() => Nodes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
