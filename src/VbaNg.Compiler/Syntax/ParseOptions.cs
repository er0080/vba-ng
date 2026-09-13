using System.Collections.Frozen;

namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Options that affect parsing: the conditional compilation constants (MS-VBAL 3.4.1). The
/// defaults describe 64-bit Excel on Windows: Win64, Win32, VBA6, and VBA7 are True; Win16 and
/// Mac are False.
/// </summary>
public sealed class ParseOptions
{
    // Static fields initialize in textual order: Predefined must come before Default.
    private static readonly FrozenDictionary<string, object?> Predefined = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
    {
        ["Win64"] = true,
        ["Win32"] = true,
        ["Win16"] = false,
        ["Mac"] = false,
        ["VBA6"] = true,
        ["VBA7"] = true,
    }.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);

    public static readonly ParseOptions Default = new(FrozenDictionary<string, object?>.Empty);

    public ParseOptions(IReadOnlyDictionary<string, object?> conditionalConstants)
    {
        ArgumentNullException.ThrowIfNull(conditionalConstants);
        var merged = new Dictionary<string, object?>(Predefined, StringComparer.OrdinalIgnoreCase);
        foreach (var (name, value) in conditionalConstants)
        {
            merged[name] = value;
        }

        ConditionalConstants = merged.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Predefined constants plus the project's own; values are Boolean, Int64, Double, or String.</summary>
    public IReadOnlyDictionary<string, object?> ConditionalConstants { get; }
}
