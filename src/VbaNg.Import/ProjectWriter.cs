using System.Globalization;
using System.Text;

namespace VbaNg.Import;

/// <summary>What an import wrote, and what it could not bring over.</summary>
public sealed record ImportResult(string ProjectDir, IReadOnlyList<string> Files, IReadOnlyList<string> Report)
{
    public string ToText()
    {
        var text = new StringBuilder();
        text.Append(CultureInfo.InvariantCulture, $"Imported {Files.Count.ToString(CultureInfo.InvariantCulture)} file(s) into {ProjectDir}");
        foreach (var file in Files)
        {
            text.Append(CultureInfo.InvariantCulture, $"{Environment.NewLine}  {file}");
        }

        foreach (var line in Report)
        {
            text.Append(CultureInfo.InvariantCulture, $"{Environment.NewLine}warning: {line}");
        }

        return text.ToString();
    }
}

/// <summary>
/// Writes a project the importer read as vba-ng source files (ARCHITECTURE.md section 10): one
/// file per module in the VBE's export format, a manifest with the project's references, and a
/// report of what did not come over. The workbook itself is never touched (D17).
/// </summary>
public static class ProjectWriter
{
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>The class header the VBE writes above a class module's attributes; the module stream does not carry it.</summary>
    private const string ClassHeader = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n";

    /// <summary>
    /// The libraries the compiler supplies on its own (ARCHITECTURE.md section 3), which the manifest
    /// does not name. Office is not among them: it is a default reference of every Excel project,
    /// not an implicit one, and the manifest says so when the project carries it.
    /// </summary>
    private static readonly string[] Implicit = ["stdole", "VBA"];

    /// <param name="documentKinds">
    /// CodeName to document kind for the workbook the project came out of, which only its caller
    /// can read; without it a document module is a Workbook when it is named ThisWorkbook and a
    /// Worksheet otherwise, the same rule the compiler falls back to (D19).
    /// </param>
    public static ImportResult Write(VbaProject project, string projectDir, string? hostLibrary = "Excel", IReadOnlyDictionary<string, string>? documentKinds = null)
    {
        ArgumentNullException.ThrowIfNull(project);
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        Directory.CreateDirectory(projectDir);

        var files = new List<string>();
        var report = new List<string>();
        foreach (var module in project.Modules)
        {
            var path = Path.Combine(projectDir, module.FileName);
            File.WriteAllText(path, Source(module), Utf8);
            files.Add(module.FileName);
            if (module.Kind == VbaModuleKind.Form)
            {
                report.Add($"{module.FileName}: only the code behind the form was imported; its layout and .frx resources stay in the workbook, and the build reports the file until UserForms are supported.");
            }

            // A Declare comes over and builds since M7 closed, so it is not something that did not
            // come over; a type the build cannot marshal is the build's own diagnostic to report.
        }

        var manifestPath = Path.Combine(projectDir, "vbang.json");
        File.WriteAllText(manifestPath, Manifest(project, hostLibrary, documentKinds), Utf8);
        files.Add("vbang.json");
        return new ImportResult(projectDir, files, report);
    }

    /// <summary>A module's file content: the class header where the VBE writes one, then the module's own text.</summary>
    private static string Source(VbaModule module) =>
        module.Kind is VbaModuleKind.Class or VbaModuleKind.Document ? ClassHeader + module.Source : module.Source;

    /// <summary>
    /// The manifest (ARCHITECTURE.md section 3): the project's name, the host's own library, and
    /// every referenced type library that carries an identity, in the order the project records
    /// them. The libraries every project has are left out. The document modules the PROJECT stream
    /// named are written with their kind, so the folder says what its .cls files are once it is in
    /// git without the workbook beside it, where the attribute pairing cannot tell a sheet module
    /// from a PublicNotCreatable class (D19).
    /// </summary>
    private static string Manifest(VbaProject project, string? hostLibrary, IReadOnlyDictionary<string, string>? documentKinds)
    {
        var references = new List<string>();
        if (!string.IsNullOrEmpty(hostLibrary))
        {
            references.Add($"    {{ \"name\": \"{hostLibrary}\" }}");
        }

        foreach (var reference in project.References)
        {
            if (reference.Name.Length == 0
                || Implicit.Contains(reference.Name, StringComparer.OrdinalIgnoreCase)
                || reference.Name.Equals(hostLibrary, StringComparison.OrdinalIgnoreCase)
                || reference.LibraryGuid is null)
            {
                continue;
            }

            var version = reference.Version is { Length: > 0 } declared ? $", \"version\": \"{declared}\"" : string.Empty;
            references.Add($"    {{ \"name\": \"{reference.Name}\", \"guid\": \"{reference.LibraryGuid}\"{version} }}");
        }

        var documents = new List<string>();
        foreach (var module in project.Modules)
        {
            if (module.Kind != VbaModuleKind.Document)
            {
                continue;
            }

            var kind = documentKinds is not null && documentKinds.TryGetValue(module.Name, out var known) && known.Length > 0
                ? known
                : module.Name.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase) ? "Workbook" : "Worksheet";
            documents.Add($"    \"{module.Name}\": \"{kind}\"");
        }

        var text = new StringBuilder();
        text.Append("{\n");
        text.Append(CultureInfo.InvariantCulture, $"  \"name\": \"{project.Name}\",\n");
        text.Append("  \"references\": [\n");
        text.Append(string.Join(",\n", references).Replace("    ", "    ", StringComparison.Ordinal));
        text.Append("\n  ]");
        if (documents.Count > 0)
        {
            text.Append(",\n  \"documents\": {\n");
            text.Append(string.Join(",\n", documents));
            text.Append("\n  }");
        }

        text.Append("\n}\n");
        return text.ToString().ReplaceLineEndings("\n");
    }
}
