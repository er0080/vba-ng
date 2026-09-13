using System.Text.Json;
using System.Text.Json.Serialization;

using VbaNg.Runtime.Hosting;

namespace VbaNg.Compiler;

/// <summary>
/// A manifest reference (ARCHITECTURE.md section 3): a type library by name, guid, and version,
/// or a vba-ng library by name alone. The <c>vbang</c> library brings the <c>Assert</c> module of
/// test procedures (section 9). Type-library references resolve in M4 and pass through unread.
/// </summary>
public sealed record ManifestReference(
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("guid")] string? Guid = null,
    [property: JsonPropertyName("version")] string? Version = null);

/// <summary>The project manifest, <c>vbang.json</c> in the project folder (ARCHITECTURE.md section 3). A missing manifest means no references.</summary>
public sealed record ProjectManifest([property: JsonPropertyName("references")] IReadOnlyList<ManifestReference>? References)
{
    /// <summary>The name of vba-ng's own library in a manifest reference.</summary>
    public const string VbangLibrary = "vbang";

    /// <summary>
    /// Document module kinds by module name (ARCHITECTURE.md section 3), as <c>"Workbook"</c>,
    /// <c>"Worksheet"</c>, or <c>"Chart"</c>. Present, it is the whole list: every document module
    /// of the project, and a class module that is not in it is a class whatever its attributes say.
    /// Absent, the workbook beside the folder decides, and without one the attribute pairing does.
    /// </summary>
    [JsonPropertyName("documents")]
    public IReadOnlyDictionary<string, string>? Documents { get; init; }

    /// <summary>
    /// The CodeNames of the workbook next to the project folder by document kind, read by
    /// <see cref="WorkbookCodeNames"/> (ARCHITECTURE.md D19); not part of the manifest file.
    /// </summary>
    [JsonIgnore]
    public IReadOnlyDictionary<string, string>? WorkbookDocuments { get; init; }

    /// <summary>
    /// Whether a .cls with the VBE's class header is a document module (ARCHITECTURE.md D19): the
    /// manifest's documents map decides first; then the workbook beside the folder, whose CodeNames
    /// are the only modules the host can bind; and only without a workbook the attribute pairing
    /// (VB_PredeclaredId with VB_Exposed), which cannot tell a sheet module from a
    /// PublicNotCreatable class with a predeclared instance.
    /// </summary>
    public bool IsDocumentModule(string moduleName, bool predeclared, bool exposed)
    {
        ArgumentNullException.ThrowIfNull(moduleName);
        if (Documents is not null)
        {
            // The map is the whole list rather than a set of exceptions: the importer writes every
            // module the PROJECT stream named a Document, which is what lets a project folder in
            // git without its workbook say which .cls files are document modules. A predeclared and
            // exposed class module that is not listed is a class, which the attributes alone cannot
            // tell (ARCHITECTURE.md D19).
            return Documents.Keys.Any(name => name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));
        }

        if (WorkbookDocuments is { } known)
        {
            return known.Keys.Any(name => name.Equals(moduleName, StringComparison.OrdinalIgnoreCase));
        }

        return predeclared && exposed;
    }

    /// <summary>The document kind of a document module: the manifest's word for it, else the workbook's, else Workbook for ThisWorkbook and Worksheet otherwise.</summary>
    public string DocumentKind(string moduleName)
    {
        ArgumentNullException.ThrowIfNull(moduleName);
        if (Documents is not null)
        {
            foreach (var (name, kind) in Documents)
            {
                if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    return kind;
                }
            }
        }

        if (WorkbookDocuments is not null)
        {
            foreach (var (name, kind) in WorkbookDocuments)
            {
                if (name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    return kind;
                }
            }
        }

        return moduleName.Equals("ThisWorkbook", StringComparison.OrdinalIgnoreCase) ? "Workbook" : "Worksheet";
    }

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    public static ProjectManifest Empty { get; } = new((IReadOnlyList<ManifestReference>?)null);

    public bool HasReference(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return References?.Any(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase)) ?? false;
    }

    /// <summary>Reads the project's manifest; a malformed one is reported as VBA0022 and read as empty.</summary>
    public static ProjectManifest Load(string projectDir, List<Diagnostic> diagnostics)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        ArgumentNullException.ThrowIfNull(diagnostics);
        var path = ProjectPaths.ManifestPath(projectDir);
        if (!File.Exists(path))
        {
            return Empty;
        }

        ProjectManifest? manifest;
        try
        {
            manifest = JsonSerializer.Deserialize<ProjectManifest>(File.ReadAllText(path), Options);
        }
        catch (JsonException ex)
        {
            var line = (int)((ex.LineNumber ?? 0) + 1);
            var column = (int)((ex.BytePositionInLine ?? 0) + 1);
            diagnostics.Add(new Diagnostic(DiagnosticIds.InvalidManifest, DiagnosticSeverity.Error, "Invalid manifest: " + ex.Message, path, line, column));
            return Empty;
        }

        manifest ??= Empty;
        if (manifest.References?.Any(r => r is null || string.IsNullOrWhiteSpace(r.Name)) == true)
        {
            diagnostics.Add(new Diagnostic(DiagnosticIds.InvalidManifest, DiagnosticSeverity.Error, "Invalid manifest: every reference needs a name.", path, 1, 1));
            return Empty;
        }

        return manifest;
    }
}
