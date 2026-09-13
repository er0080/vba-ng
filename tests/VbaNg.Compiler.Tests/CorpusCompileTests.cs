using System.Globalization;
using System.Text;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// The compatibility scorecard on real-world projects (ROADMAP.md M6, ARCHITECTURE.md section 11):
/// every project of the corpus is bound and emitted, and the report says how many compile and what
/// stops the rest. The corpus is not checked in; point VBANG_CORPUS at a folder whose immediate
/// subfolders are projects of .bas, .cls, and .frm files.
/// </summary>
public sealed class CorpusCompileTests(ITestOutputHelper output)
{
    private static readonly string[] Extensions = [".bas", ".cls", ".frm"];

    private static readonly string[] WorkbookExtensions = [".xlsm", ".xlam", ".xlsb", ".xltm", ".xlsx"];

    /// <summary>
    /// The workbooks a corpus project carries name its document modules (ARCHITECTURE.md D19):
    /// stdVBA's build and test workbooks say that <c>xlsProjectBuilder</c> is a sheet and that its
    /// seven classes with a predeclared instance are classes. A project without a workbook (xlwings
    /// ships its add-in elsewhere) is left to the attribute rule, as a folder without one is.
    /// </summary>
    private static ProjectManifest ManifestOf(string directory)
    {
        IReadOnlyDictionary<string, string>? documents = null;
        foreach (var workbook in Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories)
            .Where(f => WorkbookExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
            .Where(f => !f.Contains(Path.DirectorySeparatorChar + ".git" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            var names = WorkbookCodeNames.Read(workbook);
            if (names.Count == 0)
            {
                continue;
            }

            var union = documents is null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>(documents, StringComparer.OrdinalIgnoreCase);
            foreach (var (name, kind) in names)
            {
                union.TryAdd(name, kind);
            }

            documents = union;
        }

        return ProjectManifest.Empty with { WorkbookDocuments = documents };
    }

    /// <summary>
    /// The libraries the corpus projects reference; without them every library type would count as
    /// an error the compiler cannot help. Two fixture models are checked in (an Excel subset and
    /// Scripting), and every model the CLI has cached on this machine (ARCHITECTURE.md section 6:
    /// one JSON per library the registry provided) replaces or joins them, so a machine that has
    /// built against the real Excel, Office, MSForms, VBIDE, or ADODB libraries scores with those.
    /// The compiler never reads the registry itself (R9); it only loads what is already cached.
    /// </summary>
    private static (List<VbaNg.Runtime.TypeLibraries.ComLibrary> Libraries, string Description) LoadLibraries()
    {
        var fixtures = Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures", "TypeLibs");
        var libraries = new List<VbaNg.Runtime.TypeLibraries.ComLibrary>
        {
            VbaNg.Runtime.TypeLibraries.ComLibrary.Load(Path.Combine(fixtures, "ExcelSubset.json")),
            VbaNg.Runtime.TypeLibraries.ComLibrary.Load(Path.Combine(fixtures, "Scripting.json")),
        };
        var sources = libraries.Select(l => l.Name + " (fixture)").ToList();

        var cache = VbaNg.Runtime.TypeLibraries.ComLibrary.CacheDirectory;
        if (Directory.Exists(cache))
        {
            foreach (var file in Directory.EnumerateFiles(cache, "*.json").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                VbaNg.Runtime.TypeLibraries.ComLibrary cached;
                try
                {
                    cached = VbaNg.Runtime.TypeLibraries.ComLibrary.Load(file);
                }
                catch (Exception ex) when (ex is IOException or System.Text.Json.JsonException)
                {
                    continue;
                }

                var replaced = libraries.FindIndex(l => l.Guid == cached.Guid);
                if (replaced >= 0)
                {
                    libraries[replaced] = cached;
                    sources[replaced] = cached.Name + " (cached, replaces the fixture)";
                }
                else
                {
                    libraries.Add(cached);
                    sources.Add(cached.Name + " (cached)");
                }
            }
        }

        return (libraries, string.Join(", ", sources));
    }

    [Fact]
    public void Corpus_CompileRateIsReported()
    {
        var setting = Environment.GetEnvironmentVariable("VBANG_CORPUS");
        Assert.SkipWhen(string.IsNullOrWhiteSpace(setting), "Set VBANG_CORPUS to a folder of VBA projects (or several, separated by ';') to run the compile scorecard.");
        var roots = setting!.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var root in roots)
        {
            Assert.SkipUnless(Directory.Exists(root), $"VBANG_CORPUS folder not found: {root}");
        }

        var projects = roots.SelectMany(Projects).ToList();
        Assert.NotEmpty(projects);
        var (libraries, libraryDescription) = LoadLibraries();

        var report = new StringBuilder();
        report.AppendLine(CultureInfo.InvariantCulture, $"Corpus: {setting}");
        report.AppendLine(CultureInfo.InvariantCulture, $"Libraries: {libraryDescription}");
        report.AppendLine();

        var byMessage = new Dictionary<string, int>(StringComparer.Ordinal);
        var samples = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var byId = new Dictionary<string, int>(StringComparer.Ordinal);
        var compiled = 0;
        var files = 0;
        var filesWithoutErrors = 0;
        foreach (var (name, directory, sources) in projects)
        {
            files += sources.Count;
            var diagnostics = new List<Diagnostic>();
            ProjectCompiler.Generate(Identifier(name), sources, diagnostics, manifest: ManifestOf(directory), libraries: libraries);
            var errors = diagnostics.Where(d => d.IsError).ToList();
            var clean = sources.Count(s => !errors.Any(e => string.Equals(e.FilePath, s.Path, StringComparison.OrdinalIgnoreCase)));
            filesWithoutErrors += clean;
            if (errors.Count == 0)
            {
                compiled++;
            }

            report.AppendLine(CultureInfo.InvariantCulture, $"{name,-28} {sources.Count,4} file(s)  {clean,4} without errors  {errors.Count,5} error(s)");
            foreach (var error in errors)
            {
                byId[error.Id] = byId.GetValueOrDefault(error.Id) + 1;
                var shape = Shorten(error.Message);
                byMessage[shape] = byMessage.GetValueOrDefault(shape) + 1;
                var sample = samples.TryGetValue(shape, out var list) ? list : samples[shape] = [];
                if (sample.Count < 3)
                {
                    sample.Add(Path.GetFileName(error.FilePath) + "(" + error.Line.ToString(CultureInfo.InvariantCulture) + "): " + error.Message);
                }
            }
        }

        report.AppendLine();
        report.AppendLine(CultureInfo.InvariantCulture, $"Projects: {projects.Count}, compiling: {compiled} ({compiled / (double)projects.Count:P0})");
        report.AppendLine(CultureInfo.InvariantCulture, $"Files: {files}, without errors: {filesWithoutErrors} ({filesWithoutErrors / (double)files:P0})");
        report.AppendLine();
        report.AppendLine("Errors by id:");
        foreach (var (id, count) in byId.OrderByDescending(p => p.Value))
        {
            report.AppendLine(CultureInfo.InvariantCulture, $"  {count,5}  {id}");
        }

        report.AppendLine();
        report.AppendLine("Errors by message:");
        foreach (var (message, count) in byMessage.OrderByDescending(p => p.Value).Take(25))
        {
            report.AppendLine(CultureInfo.InvariantCulture, $"  {count,5}  {message}");
            foreach (var sample in samples.GetValueOrDefault(message) ?? [])
            {
                report.AppendLine("           " + sample);
            }
        }

        output.WriteLine(report.ToString());
        var reportPath = Path.Combine(AppContext.BaseDirectory, "corpus-compile-report.txt");
        File.WriteAllText(reportPath, report.ToString());
        output.WriteLine("Report written to " + reportPath);
    }

    /// <summary>A project is an immediate subfolder of the corpus root that holds sources, or the root itself when it does.</summary>
    private static IEnumerable<(string Name, string Directory, List<SourceFile> Sources)> Projects(string root)
    {
        var atRoot = Sources(root, SearchOption.TopDirectoryOnly).ToList();
        if (atRoot.Count > 0)
        {
            yield return (Path.GetFileName(Path.TrimEndingDirectorySeparator(root)), root, atRoot);
        }

        foreach (var directory in Directory.EnumerateDirectories(root).OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
        {
            var sources = Sources(directory, SearchOption.AllDirectories).ToList();
            if (sources.Count > 0)
            {
                yield return (Path.GetFileName(directory), directory, sources);
            }
        }
    }

    private static IEnumerable<SourceFile> Sources(string directory, SearchOption option) =>
        Directory.EnumerateFiles(directory, "*", option)
            .Where(f => Extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
            .Where(f => !f.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .Where(f => !f.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .Select(f => new SourceFile(f, ReadSource(f)));

    /// <summary>The corpus folders are named after their repository, which is not always a VBA identifier.</summary>
    private static string Identifier(string name)
    {
        var text = new StringBuilder();
        foreach (var character in name)
        {
            text.Append(char.IsAsciiLetterOrDigit(character) ? character : '_');
        }

        return text.Length == 0 || !char.IsAsciiLetter(text[0]) ? "P" + text : text.ToString();
    }

    /// <summary>Messages carry names and paths; the shape of the message is what the scorecard counts.</summary>
    private static string Shorten(string message)
    {
        var quote = message.IndexOf('\'', StringComparison.Ordinal);
        return quote < 0 ? message : message[..quote].TrimEnd() + " ...";
    }

    private static string ReadSource(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        try
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.Latin1.GetString(bytes);
        }
    }
}
