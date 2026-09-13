using System.Globalization;
using System.IO.Compression;
using System.Text;

using VbaNg.Compiler;
using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang report</c> (ROADMAP.md WP7): one zip a bug report can carry. It holds what is needed to
/// reproduce a failure and nothing else: the manifest, the build's own diagnostics, the generated
/// C#, the project's log, and the running add-in's status when Excel is reachable. Never the
/// workbook, so the bundle can go on a public issue without the data going with it.
/// </summary>
internal static class ReportCommand
{
    public static int Run(CommandLineOptions options)
    {
        var positional = options.Positionals.Count > 0 ? options.Positionals[0] : null;
        var projectDir = ProjectLocator.Resolve(options.Project ?? positional, out var error);
        if (projectDir is null)
        {
            Console.Error.WriteLine("vbang: " + error);
            return ExitCodes.Usage;
        }

        var name = ProjectPaths.ProjectName(projectDir);
        var zipPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            string.Create(CultureInfo.InvariantCulture, $"vbang-report-{name}-{DateTime.Now:yyyyMMdd-HHmmss}.zip"));

        // The build runs here so the report carries the diagnostics as they are now, rather than
        // whatever the last build happened to leave behind.
        var build = ProjectCompiler.Build(projectDir, TypeLibraryCache.Resolve);

        // Asked once: the reason Excel cannot be reached is worth printing, but only the once.
        var status = HostCommandClient.Call("vbang.Status", [], out var statusJson) == ExitCodes.Ok ? statusJson : null;

        using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
        {
            Add(zip, "versions.txt", VersionsFile(projectDir, build, status));
            Add(zip, "build.log", BuildLog(build));
            AddFile(zip, "vbang.json", ProjectPaths.ManifestPath(projectDir));
            AddFile(zip, "out/build.json", ProjectPaths.BuildInfoPath(projectDir));
            AddFile(zip, "out/output.log", ProjectPaths.OutputLogPath(projectDir));

            var generatedDir = ProjectPaths.GeneratedDir(projectDir);
            if (Directory.Exists(generatedDir))
            {
                foreach (var file in Directory.EnumerateFiles(generatedDir, "*.cs").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
                {
                    AddFile(zip, "out/gen/" + Path.GetFileName(file), file);
                }
            }

            foreach (var source in Directory.EnumerateFiles(projectDir)
                .Where(f => f.EndsWith(".bas", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".cls", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".frm", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
            {
                AddFile(zip, "source/" + Path.GetFileName(source), source);
            }

            if (status is not null)
            {
                Add(zip, "status.json", status);
            }
        }

        Console.WriteLine($"Wrote {zipPath}");
        Console.WriteLine("It holds the manifest, the sources, the build's diagnostics, the generated C#, the project's log, and the add-in's status. It does not hold the workbook.");
        return ExitCodes.Ok;
    }

    /// <summary>What a maintainer needs before reading anything else: which build of what, on which Excel.</summary>
    private static string VersionsFile(string projectDir, BuildResult build, string? status)
    {
        var text = new StringBuilder();
        text.Append(CultureInfo.InvariantCulture, $"vbang: {Versions.Vbang}\n");
        text.Append(CultureInfo.InvariantCulture, $"runtime: {Versions.Runtime}\n");
        text.Append(CultureInfo.InvariantCulture, $"os: {Environment.OSVersion.VersionString} ({(Environment.Is64BitProcess ? "x64" : "x86")})\n");
        text.Append(CultureInfo.InvariantCulture, $"culture: {CultureInfo.CurrentCulture.Name}\n");
        text.Append(CultureInfo.InvariantCulture, $"project: {projectDir}\n");
        text.Append(CultureInfo.InvariantCulture, $"build: {(build.Success ? "succeeded" : "failed")}, {build.Diagnostics.Count} diagnostic(s)\n");
        text.Append(CultureInfo.InvariantCulture, $"excel: {(status is null ? "not reachable; the add-in's status is not in this report" : RunResponse.FromJson(status).Output)}\n");
        return text.ToString();
    }

    private static string BuildLog(BuildResult build)
    {
        var text = new StringBuilder();
        foreach (var diagnostic in build.Diagnostics)
        {
            text.Append(diagnostic).Append('\n');
        }

        if (build.Diagnostics.Count == 0)
        {
            text.Append("No diagnostics.\n");
        }

        return text.ToString();
    }

    private static void Add(ZipArchive zip, string entryName, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(entryName).Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    /// <summary>A file the project may not have; a report says what is there rather than failing over what is not.</summary>
    private static void AddFile(ZipArchive zip, string entryName, string path)
    {
        if (File.Exists(path))
        {
            zip.CreateEntryFromFile(path, entryName);
        }
    }
}
