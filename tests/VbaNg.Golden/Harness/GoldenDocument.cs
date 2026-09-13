using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaNg.Golden.Harness;

/// <summary>
/// A value recorded by real VBA. <see cref="Type"/> is VBA's <c>TypeName</c> ("Integer", "String",
/// "Empty", "Null", "Nothing", "Variant()", ...) or "Object" with the class in <see cref="Class"/>.
/// <see cref="Value"/> is the readable form (<c>Str</c> for numbers, so it is locale independent;
/// the text itself for strings; <c>True</c>/<c>False</c>). Exact bits travel separately:
/// <see cref="Bits"/> holds the IEEE bits of Single and Double, the serial of Date, and the scaled
/// integer of Currency as big-endian hex; <see cref="Bytes"/> holds the 16 VARIANT bytes of a
/// Decimal. A String with a lone surrogate, which JSON cannot carry, shows U+FFFD in
/// <see cref="Value"/> and its exact UTF-16 code units, four hex digits each, in <see cref="Units"/>.
/// Arrays carry <see cref="Bounds"/> in VBA notation ("0 To 2, 1 To 3", empty when unallocated)
/// and, for one and two dimensions, their <see cref="Items"/> in row-major order.
/// </summary>
public sealed record GoldenValue(
    string Type,
    string? Value = null,
    string? Bits = null,
    string? Bytes = null,
    string? Units = null,
    string? Class = null,
    string? Bounds = null,
    IReadOnlyList<GoldenValue>? Items = null);

/// <summary>One recorded value: which <c>?</c> line produced it (0-based) and what it was.</summary>
public sealed record GoldenResult(int Index, GoldenValue Value);

/// <summary>The VBA error that ended a case, as <c>Err</c> reported it.</summary>
public sealed record GoldenError(int Number, string Description, string Source);

/// <summary>A case and what real VBA did with it: the values in execution order, then the error if one ended the case.</summary>
public sealed record GoldenCase(string Name, IReadOnlyList<string> Source, IReadOnlyList<GoldenResult> Results, GoldenError? Error);

/// <summary>
/// One golden file: a library area recorded by a given Excel build under a given culture. Committed
/// under tests/VbaNg.Golden/Goldens and only ever rewritten by the regen tool (CLAUDE.md R16).
/// <see cref="Declarations"/> holds the module-level declarations of the case file, or null when it has none;
/// <see cref="Classes"/> the class modules recorded with it and <see cref="Modules"/> the standard
/// modules, each null when it has none.
/// </summary>
public sealed record GoldenArea(string Area, string Excel, string Culture, IReadOnlyList<string> Options, IReadOnlyList<string>? Declarations, IReadOnlyList<GoldenModule>? Classes, IReadOnlyList<GoldenModule>? Modules, IReadOnlyList<GoldenCase> Cases)
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        NewLine = "\n",
        // Keep VBA source readable in diffs: no + for a plus sign. Control characters stay escaped.
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static GoldenArea Load(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        using var stream = File.OpenRead(path);
        return JsonSerializer.Deserialize<GoldenArea>(stream, JsonOptions)
            ?? throw new JsonException($"Golden file is empty: {path}");
    }

    public string ToJson() => JsonSerializer.Serialize(this, JsonOptions) + "\n";

    public void Save(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, ToJson());
    }

    /// <summary>
    /// Combines the case sources with the lines the recorder module wrote. Every case appears in the
    /// golden, with or without recorded values, so a case that recorded nothing is visible.
    /// </summary>
    public static GoldenArea Merge(CaseArea area, IReadOnlyList<RecordLine> lines, string excel, string culture)
    {
        ArgumentNullException.ThrowIfNull(area);
        ArgumentNullException.ThrowIfNull(lines);
        if (!lines.Any(l => l.Done == true))
        {
            throw new InvalidOperationException($"{area.Name}: the recorder did not finish; the results end after {lines.Count} line(s).");
        }

        var results = new List<GoldenResult>[area.Cases.Count];
        var errors = new GoldenError?[area.Cases.Count];
        for (var i = 0; i < results.Length; i++)
        {
            results[i] = [];
        }

        foreach (var line in lines)
        {
            if (line.Case is not { } number)
            {
                continue;
            }

            if (number < 1 || number > area.Cases.Count)
            {
                throw new InvalidOperationException($"{area.Name}: the recorder reported case {number}, but there are {area.Cases.Count} cases.");
            }

            if (line.Error is not null)
            {
                errors[number - 1] = errors[number - 1] is null
                    ? line.Error
                    : throw new InvalidOperationException($"{area.Name}: case {number} reported two errors.");
            }
            else if (line.Value is not null && line.Index is { } index)
            {
                results[number - 1].Add(new GoldenResult(index, line.Value));
            }
        }

        var cases = area.Cases
            .Select((source, i) => new GoldenCase(source.Name, source.Lines, results[i], errors[i]))
            .ToList();
        return new GoldenArea(
            area.Name,
            excel,
            culture,
            area.Options,
            area.Declarations.Count == 0 ? null : area.Declarations,
            Recorded(area.Companions.Where(c => c.IsClass)),
            Recorded(area.Companions.Where(c => !c.IsClass)),
            cases);
    }

    private static List<GoldenModule>? Recorded(IEnumerable<CompanionSource> companions)
    {
        var recorded = companions.Select(c => new GoldenModule(c.Name, c.Lines)).ToList();
        return recorded.Count == 0 ? null : recorded;
    }

    /// <summary>Parses the recorder's output: one JSON object per line.</summary>
    public static IReadOnlyList<RecordLine> ReadRecordLines(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var lines = new List<RecordLine>();
        foreach (var raw in text.Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0)
            {
                continue;
            }

            lines.Add(JsonSerializer.Deserialize<RecordLine>(line, JsonOptions)
                ?? throw new JsonException("Empty record line."));
        }

        return lines;
    }
}

/// <summary>One line written by the recorder module: a value, an error, or the final done marker.</summary>
public sealed record RecordLine(int? Case, int? Index, GoldenValue? Value, GoldenError? Error, bool? Done);

/// <summary>A companion module recorded with an area, class or standard: its name and its source lines as the VBE exported them.</summary>
public sealed record GoldenModule(string Name, IReadOnlyList<string> Source);
