using System.Globalization;

namespace VbaNg.Golden.Harness;

/// <summary>
/// One golden case: a few lines of VBA run inside a procedure of the recorder module. Lines that
/// start with <c>?</c> (the Immediate window's print shorthand) record the value of the expression
/// that follows; every other line runs as written (ARCHITECTURE.md section 11).
/// </summary>
public sealed record GoldenCaseSource(string Name, int Line, IReadOnlyList<string> Lines)
{
    /// <summary>The expressions of the <c>?</c> lines, in order.</summary>
    public IEnumerable<string> Expressions => Lines.Where(IsResultLine).Select(ResultExpression);

    public static bool IsResultLine(string line)
    {
        ArgumentNullException.ThrowIfNull(line);
        return line.TrimStart().StartsWith('?');
    }

    public static string ResultExpression(string line)
    {
        ArgumentNullException.ThrowIfNull(line);
        return line.TrimStart()[1..].Trim();
    }
}

/// <summary>
/// A companion module an area's cases use: a <c>&lt;Area&gt;.&lt;Name&gt;.cls</c> class module or a
/// <c>&lt;Area&gt;.&lt;Name&gt;.bas</c> standard module in the VBE's export format next to the case
/// file, so attribute-only declarations (default members, <c>NewEnum</c>, predeclared instances)
/// read exactly as the VBE reads them, and a second standard module gives scope its other side.
/// </summary>
public sealed record CompanionSource(string Name, string Extension, IReadOnlyList<string> Lines)
{
    public bool IsClass => string.Equals(Extension, ".cls", StringComparison.OrdinalIgnoreCase);

    public string FileName => Name + Extension;
}

/// <summary>
/// A case file: one library area, its module-level <c>Option</c> lines, the module-level
/// declarations its cases share (procedures, types, enums, module variables), the companion
/// modules next to it, and its cases.
/// </summary>
public sealed record CaseArea(string Name, IReadOnlyList<string> Options, IReadOnlyList<string> Declarations, IReadOnlyList<CompanionSource> Companions, IReadOnlyList<GoldenCaseSource> Cases);

/// <summary>
/// Reads <c>.cases</c> files. Blank lines separate blocks. Comment lines (<c>'</c>) at the top of a
/// block describe it; the last one is the case name. Before the first case, a block made only of
/// <c>Option</c> lines sets the recorder module's options, and a block that starts with a
/// procedure, Type, Enum, or module-level declaration (Public, Private, Global, Friend, Declare)
/// is emitted at module level, in order, so cases can call helper procedures; module variables
/// must come before procedures there, as VBA requires. Sources are ASCII only, so the text
/// survives the trip through the VBE unchanged; use <c>ChrW</c> for anything else.
/// </summary>
public static class CaseFile
{
    public const string Extension = ".cases";

    private static readonly string[] DeclarationKeywords = ["Public", "Private", "Friend", "Global", "Sub", "Function", "Property", "Type", "Enum", "Declare"];

    public static CaseArea Load(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        var areaName = Path.GetFileNameWithoutExtension(path);
        var area = Parse(areaName, File.ReadAllText(path));
        return area with { Companions = LoadCompanions(Path.GetDirectoryName(Path.GetFullPath(path))!, areaName) };
    }

    /// <summary>The companion modules of an area: every <c>&lt;Area&gt;.&lt;Name&gt;.cls</c> beside the case file, then every <c>.bas</c>, by name.</summary>
    private static List<CompanionSource> LoadCompanions(string directory, string areaName)
    {
        var companions = new List<CompanionSource>();
        foreach (var extension in new[] { ".cls", ".bas" })
        {
            foreach (var file in Directory.GetFiles(directory, areaName + ".*" + extension).OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                var name = Path.GetFileNameWithoutExtension(file)[(areaName.Length + 1)..];
                if (!IsIdentifier(name))
                {
                    throw new FormatException($"Companion file '{Path.GetFileName(file)}' must be named <Area>.<Module>{extension} with a VBA identifier as the module name.");
                }

                var lines = File.ReadAllText(file).Split("\r\n").ToList();
                if (lines.Count > 0 && lines[^1].Length == 0)
                {
                    lines.RemoveAt(lines.Count - 1);
                }

                for (var i = 0; i < lines.Count; i++)
                {
                    var offending = lines[i].FirstOrDefault(c => c > '\x7E' || (c < ' ' && c != '\t'));
                    if (offending != '\0')
                    {
                        throw new FormatException(string.Create(
                            CultureInfo.InvariantCulture,
                            $"{Path.GetFileName(file)}({i + 1}): character U+{(int)offending:X4} is not ASCII; write it with ChrW."));
                    }
                }

                // The VBE names an imported module after this attribute, and the cases qualify names with the file's name.
                if (!lines.Contains("Attribute VB_Name = \"" + name + "\""))
                {
                    throw new FormatException($"Companion file '{Path.GetFileName(file)}' must carry Attribute VB_Name = \"{name}\", the name the VBE gives it on import.");
                }

                companions.Add(new CompanionSource(name, extension, lines));
            }
        }

        return companions;
    }

    public static CaseArea Parse(string areaName, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(areaName);
        ArgumentNullException.ThrowIfNull(text);
        if (!IsIdentifier(areaName))
        {
            throw new FormatException($"Area name '{areaName}' must be a VBA identifier; it names the recorder module.");
        }

        var options = new List<string>();
        var declarations = new List<string>();
        var cases = new List<GoldenCaseSource>();
        foreach (var (startLine, block) in Blocks(text))
        {
            var description = block.TakeWhile(IsComment).ToList();
            var code = block.Skip(description.Count).ToList();
            if (code.Count == 0)
            {
                continue;
            }

            var codeLine = startLine + description.Count;
            for (var i = 0; i < code.Count; i++)
            {
                var offending = code[i].FirstOrDefault(c => c > '\x7E' || (c < ' ' && c != '\t'));
                if (offending != '\0')
                {
                    throw new FormatException(string.Create(
                        CultureInfo.InvariantCulture,
                        $"{areaName}{Extension}({codeLine + i}): character U+{(int)offending:X4} is not ASCII; write it with ChrW."));
                }
            }

            if (code.All(l => l.TrimStart().StartsWith("Option ", StringComparison.OrdinalIgnoreCase)))
            {
                if (cases.Count > 0)
                {
                    throw new FormatException($"{areaName}{Extension}({codeLine}): Option lines must come before the first case.");
                }

                options.AddRange(code.Select(l => l.Trim()));
                continue;
            }

            if (IsDeclaration(code[0]))
            {
                if (cases.Count > 0)
                {
                    throw new FormatException($"{areaName}{Extension}({codeLine}): declarations must come before the first case.");
                }

                if (declarations.Count > 0)
                {
                    declarations.Add(string.Empty);
                }

                declarations.AddRange(code.Select(l => l.TrimEnd()));
                continue;
            }

            var name = description.Count > 0
                ? description[^1].TrimStart().TrimStart('\'').Trim()
                : code.Where(GoldenCaseSource.IsResultLine).Select(GoldenCaseSource.ResultExpression).FirstOrDefault() ?? code[0].Trim();
            cases.Add(new GoldenCaseSource(name, codeLine, code));
        }

        return new CaseArea(areaName, options, declarations, [], cases);
    }

    private static bool IsDeclaration(string line)
    {
        var trimmed = line.TrimStart();
        var end = trimmed.IndexOf(' ', StringComparison.Ordinal);
        var word = end < 0 ? trimmed : trimmed[..end];
        if (word.Equals("Static", StringComparison.OrdinalIgnoreCase))
        {
            var rest = trimmed[(end + 1)..].TrimStart();
            return rest.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) || rest.StartsWith("Function ", StringComparison.OrdinalIgnoreCase);
        }

        return DeclarationKeywords.Any(k => k.Equals(word, StringComparison.OrdinalIgnoreCase));
    }

    private static IEnumerable<(int StartLine, List<string> Lines)> Blocks(string text)
    {
        var lines = text.Split('\n').Select(l => l.TrimEnd('\r')).ToList();
        List<string>? block = null;
        var blockStart = 0;
        for (var i = 0; i < lines.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                if (block is not null)
                {
                    yield return (blockStart, block);
                    block = null;
                }

                continue;
            }

            if (block is null)
            {
                block = [];
                blockStart = i + 1;
            }

            block.Add(lines[i]);
        }

        if (block is not null)
        {
            yield return (blockStart, block);
        }
    }

    private static bool IsComment(string line) => line.TrimStart().StartsWith('\'');

    private static bool IsIdentifier(string name) =>
        char.IsAsciiLetter(name[0]) && name.All(c => char.IsAsciiLetterOrDigit(c) || c == '_');
}
