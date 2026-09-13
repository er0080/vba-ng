using System.Globalization;
using System.Text;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.CodeAnalysis.Text;

using RoslynDiagnostic = Microsoft.CodeAnalysis.Diagnostic;
using RoslynSeverity = Microsoft.CodeAnalysis.DiagnosticSeverity;

namespace VbaNg.Compiler;

/// <summary>
/// The Roslyn driver: compiles generated C# to an assembly plus a portable PDB (ARCHITECTURE.md D2).
/// Compilation is deterministic so identical inputs produce identical bytes (CLAUDE.md R10).
/// </summary>
public static class CSharpCompiler
{
    private static readonly Lazy<List<MetadataReference>> References = new(LoadReferences);

    /// <summary>
    /// Compiles the files and writes the assembly and symbols only when compilation succeeds.
    /// Roslyn errors are reported as VBA0003 at the mapped VBA location; they always indicate a compiler bug.
    /// </summary>
    public static IReadOnlyList<Diagnostic> Compile(
        string assemblyName,
        IReadOnlyList<GeneratedFile> files,
        string assemblyPath,
        string symbolsPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(assemblyName);
        ArgumentNullException.ThrowIfNull(files);
        ArgumentException.ThrowIfNullOrWhiteSpace(assemblyPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(symbolsPath);

        var parseOptions = new CSharpParseOptions(LanguageVersion.Latest);
        var trees = files
            .Select(f => CSharpSyntaxTree.ParseText(SourceText.From(f.Text, Encoding.UTF8), parseOptions, f.Path))
            .ToList();

        var compilationOptions = new CSharpCompilationOptions(
            OutputKind.DynamicallyLinkedLibrary,
            optimizationLevel: OptimizationLevel.Debug,
            deterministic: true,
            nullableContextOptions: NullableContextOptions.Enable);

        var compilation = CSharpCompilation.Create(assemblyName, trees, References.Value, compilationOptions);
        var emitOptions = new EmitOptions(debugInformationFormat: DebugInformationFormat.PortablePdb, pdbFilePath: symbolsPath);

        using var assemblyStream = new MemoryStream();
        using var symbolsStream = new MemoryStream();
        var result = compilation.Emit(assemblyStream, symbolsStream, options: emitOptions);

        var diagnostics = result.Diagnostics
            .Where(d => d.Severity == RoslynSeverity.Error)
            .Select(Map)
            .ToList();

        if (result.Success)
        {
            File.WriteAllBytes(assemblyPath, assemblyStream.ToArray());
            File.WriteAllBytes(symbolsPath, symbolsStream.ToArray());
        }

        return diagnostics;
    }

    private static Diagnostic Map(RoslynDiagnostic diagnostic)
    {
        var span = diagnostic.Location.GetMappedLineSpan();
        var message = string.Create(
            CultureInfo.InvariantCulture,
            $"Internal compiler error {diagnostic.Id}: {diagnostic.GetMessage(CultureInfo.InvariantCulture)}");
        return new Diagnostic(
            DiagnosticIds.InternalError,
            DiagnosticSeverity.Error,
            message,
            span.Path,
            span.StartLinePosition.Line + 1,
            span.StartLinePosition.Character + 1);
    }

    private static List<MetadataReference> LoadReferences()
    {
        // Compile against the framework assemblies of the running process plus VbaNg.Runtime.
        // A reference-assembly pack replaces this once builds must target a runtime other than the CLI's.
        var trusted = (AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") as string ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        var wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "System.Runtime.dll",
            "System.Private.CoreLib.dll",
            "netstandard.dll",
            "System.Collections.dll",
            "System.Linq.dll",
        };

        var references = trusted
            .Where(path => wanted.Contains(Path.GetFileName(path)))
            .Select(path => (MetadataReference)MetadataReference.CreateFromFile(path))
            .ToList();

        references.Add(MetadataReference.CreateFromFile(typeof(Runtime.Host).Assembly.Location));
        return references;
    }
}
