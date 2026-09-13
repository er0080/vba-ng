using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

using VbaNg.Compiler.Binding;
using VbaNg.Compiler.Emit;
using VbaNg.Compiler.Syntax;
using VbaNg.Runtime.Hosting;
using VbaNg.Runtime.TypeLibraries;

using InformationalVersion = System.Reflection.AssemblyInformationalVersionAttribute;

namespace VbaNg.Compiler;

public sealed record BuildResult(
    bool Success,
    string ProjectName,
    string? AssemblyPath,
    IReadOnlyList<Diagnostic> Diagnostics);

/// <summary>
/// Builds one project folder: parses every module, binds the project, emits C# into <c>out/gen</c>,
/// compiles it into <c>out/&lt;Project&gt;.dll</c> with a portable PDB, and records input hashes in
/// <c>out/build.json</c> so a host can tell when the build is stale (ARCHITECTURE.md section 4).
/// </summary>
public static class ProjectCompiler
{
    private static readonly JsonSerializerOptions BuildInfoOptions = new() { WriteIndented = true };

    /// <summary>
    /// Builds a project folder. <paramref name="resolveReference"/> supplies the model of each
    /// type library the manifest references (the CLI reads them through Interop and caches them);
    /// without it, or when it returns null, the reference is reported as VBA0023.
    /// </summary>
    /// <summary>
    /// A form's code as a class module (ROADMAP.md D-D). The VBE writes the form designer's block
    /// above the attributes and the importer writes none; the block becomes blank lines rather than
    /// going, so every statement keeps the line it is on and its <c>#line</c> with it (R10).
    /// </summary>
    private static SourceFile FormSource(string path)
    {
        var file = SourceFile.Load(path);
        var start = file.Text.IndexOf("Attribute VB_Name", StringComparison.Ordinal);
        if (start <= 0)
        {
            return file;
        }

        var designer = file.Text[..start];
        return file with { Text = new string('\n', designer.AsSpan().Count('\n')) + file.Text[start..] };
    }

    public static BuildResult Build(string projectDir, Func<ManifestReference, ComLibrary?>? resolveReference = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        var fullDir = Path.GetFullPath(projectDir);
        if (!Directory.Exists(fullDir))
        {
            throw new DirectoryNotFoundException($"Project folder not found: {fullDir}");
        }

        var name = ProjectPaths.ProjectName(fullDir);
        var sources = Directory.GetFiles(fullDir, "*.bas").Concat(Directory.GetFiles(fullDir, "*.cls"))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Select(SourceFile.Load)
            .ToList();

        var diagnostics = new List<Diagnostic>();
        foreach (var form in Directory.GetFiles(fullDir, "*.frm").OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            // The alpha does not carry UserForms, but the rest of a workbook that has one still has
            // to build, so the form is a warning and only its signatures compile.
            diagnostics.Add(new Diagnostic(DiagnosticIds.NotSupported, DiagnosticSeverity.Warning, "UserForms are not supported yet: only the form's signatures compile, its controls are Nothing, and every procedure of it raises error 438.", form, 1, 1));
            sources.Add(FormSource(form));
        }
        // The workbook beside the folder says which .cls files are its document modules (ARCHITECTURE.md D19).
        var manifest = ProjectManifest.Load(fullDir, diagnostics) with { WorkbookDocuments = WorkbookCodeNames.ReadBeside(fullDir) };
        if (diagnostics.Any(d => d.IsError))
        {
            return new BuildResult(false, name, null, diagnostics);
        }

        var libraries = ResolveReferences(manifest, resolveReference, ProjectPaths.ManifestPath(fullDir), diagnostics);
        if (diagnostics.Any(d => d.IsError))
        {
            return new BuildResult(false, name, null, diagnostics);
        }

        var generated = Generate(name, sources, diagnostics, manifest: manifest, libraries: libraries);
        if (generated is null)
        {
            return new BuildResult(false, name, null, diagnostics);
        }

        var generatedDir = ProjectPaths.GeneratedDir(fullDir);
        Directory.CreateDirectory(generatedDir);
        var files = new List<GeneratedFile>();
        foreach (var (moduleName, text) in generated)
        {
            var generatedPath = Path.Combine(generatedDir, moduleName + ".cs");
            File.WriteAllText(generatedPath, text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            files.Add(new GeneratedFile(generatedPath, text));
        }

        var assemblyPath = ProjectPaths.AssemblyPath(fullDir);
        var symbolsPath = ProjectPaths.SymbolsPath(fullDir);
        diagnostics.AddRange(CSharpCompiler.Compile(name, files, assemblyPath, symbolsPath));
        if (diagnostics.Any(d => d.IsError))
        {
            return new BuildResult(false, name, null, diagnostics);
        }

        WriteBuildInfo(fullDir, name, sources);
        return new BuildResult(true, name, assemblyPath, diagnostics);
    }

    /// <summary>
    /// Parses, binds, and emits a set of modules, returning the C# per module, or null after
    /// reporting errors. Hosts that compile in memory (the golden replay) use this directly.
    /// </summary>
    public static IReadOnlyList<(string Module, string Text)>? Generate(string projectName, IReadOnlyList<SourceFile> sources, List<Diagnostic> diagnostics, ParseOptions? options = null, ProjectManifest? manifest = null, IReadOnlyList<ComLibrary>? libraries = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectName);
        ArgumentNullException.ThrowIfNull(sources);
        ArgumentNullException.ThrowIfNull(diagnostics);
        manifest ??= ProjectManifest.Empty;

        // The strings the binder folds and the emitter reads are Variants of the runtime, temporaries of this compilation (ARCHITECTURE.md D20).
        using var frame = VbaNg.Runtime.ObjectRefs.Frame();
        var trees = new List<SyntaxTree>();
        foreach (var source in sources)
        {
            var tree = SyntaxTree.Parse(source, options);
            diagnostics.AddRange(tree.Diagnostics);
            trees.Add(tree);
        }

        if (diagnostics.Any(d => d.IsError))
        {
            return null;
        }

        var bound = Binder.Bind(projectName, trees, diagnostics, manifest, libraries);
        if (bound is null)
        {
            return null;
        }

        var generated = new List<(string, string)>();
        foreach (var module in bound.Modules)
        {
            var text = CSharpEmitter.Emit(module, diagnostics);
            if (text is not null)
            {
                generated.Add((module.Symbol.Name, text));
            }
        }

        return diagnostics.Any(d => d.IsError) ? null : generated;
    }

    /// <summary>The type libraries the manifest references, in manifest order; the vbang library is the binder's own and needs no model.</summary>
    private static List<ComLibrary> ResolveReferences(ProjectManifest manifest, Func<ManifestReference, ComLibrary?>? resolveReference, string manifestPath, List<Diagnostic> diagnostics)
    {
        var libraries = new List<ComLibrary>();
        foreach (var reference in manifest.References ?? [])
        {
            if (string.Equals(reference.Name, ProjectManifest.VbangLibrary, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var library = resolveReference?.Invoke(reference);
            if (library is null)
            {
                var description = reference.Guid is null ? reference.Name : $"{reference.Name} {reference.Guid} {reference.Version}";
                diagnostics.Add(new Diagnostic(DiagnosticIds.ReferenceNotFound, DiagnosticSeverity.Error, $"Can't find project or library: {description}.", manifestPath, 1, 1));
                continue;
            }

            libraries.Add(library);
        }

        // Every VBA project references stdole, and the manifest does not name it (section 4), so it
        // is resolved here rather than left out: its interfaces name the pictures, fonts, and
        // enumerators projects declare (stdole.IPicture, stdole.StdFont, stdole.IEnumVARIANT), which
        // the VBE resolves and vba-ng could not. Last, so a manifest reference of the same name
        // still wins, and best effort, because dotnet test must not need a type library (R9).
        if (resolveReference is not null
            && !libraries.Any(library => library.Name.Equals("stdole", StringComparison.OrdinalIgnoreCase))
            && resolveReference(new ManifestReference("stdole")) is { } stdole)
        {
            libraries.Add(stdole);
        }

        return libraries;
    }

    private static void WriteBuildInfo(string projectDir, string name, IReadOnlyList<SourceFile> sources)
    {
        var inputs = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in sources)
        {
            inputs[Path.GetFileName(source.Path)] = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source.Text)));
        }

        var manifestPath = ProjectPaths.ManifestPath(projectDir);
        if (File.Exists(manifestPath))
        {
            inputs[Path.GetFileName(manifestPath)] = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(manifestPath)));
        }

        var info = new BuildInfo(name, CompilerVersion, inputs);
        File.WriteAllText(ProjectPaths.BuildInfoPath(projectDir), JsonSerializer.Serialize(info, BuildInfoOptions));
    }

    /// <summary>
    /// The compiler's own version, read off the assembly so no version is written down twice. It
    /// stamps out/build.json, so a project built by an older compiler is stale.
    /// </summary>
    internal static string CompilerVersion =>
        System.Reflection.CustomAttributeExtensions.GetCustomAttribute<InformationalVersion>(typeof(ProjectCompiler).Assembly)?.InformationalVersion
        ?? typeof(ProjectCompiler).Assembly.GetName().Version?.ToString()
        ?? "0";

    private sealed record BuildInfo(string Project, string Compiler, SortedDictionary<string, string> Inputs);
}
