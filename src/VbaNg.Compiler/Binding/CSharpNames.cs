using System.Collections.Frozen;

namespace VbaNg.Compiler.Binding;

/// <summary>Turns VBA names into C# identifiers: the same spelling, prefixed with @ when C# reserves the word (CLAUDE.md R10: identifiers stay recognizable).</summary>
public static class CSharpNames
{
    private static readonly FrozenSet<string> Reserved = new[]
    {
        "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class", "const", "continue",
        "decimal", "default", "delegate", "do", "double", "else", "enum", "event", "explicit", "extern", "false", "finally",
        "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int", "interface", "internal", "is", "lock",
        "long", "namespace", "new", "null", "object", "operator", "out", "override", "params", "private", "protected",
        "public", "readonly", "ref", "return", "sbyte", "sealed", "short", "sizeof", "stackalloc", "static", "string",
        "struct", "switch", "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked", "unsafe", "ushort",
        "using", "virtual", "void", "volatile", "while",
    }.ToFrozenSet(StringComparer.Ordinal);

    public static string Identifier(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return Reserved.Contains(name) ? "@" + name : name;
    }
}
