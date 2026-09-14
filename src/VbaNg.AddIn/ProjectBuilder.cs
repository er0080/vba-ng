using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

internal enum BuildOutcome
{
    /// <summary>out/build.json matches the sources; nothing to do.</summary>
    UpToDate,

    /// <summary>vbang build ran and succeeded.</summary>
    Built,

    /// <summary>vbang build ran and failed; <see cref="ProjectBuilder.LastOutput"/> has its diagnostics.</summary>
    Failed,

    /// <summary>The sources changed but no vbang CLI was found; the existing output, if any, is used as is.</summary>
    NoCli,
}

/// <summary>
/// Keeps a project's build current before the add-in loads it (ARCHITECTURE.md section 7,
/// lifecycle step 2, and D1): when <c>out/build.json</c> is missing or its input hashes differ
/// from the sources on disk, <c>vbang build</c> runs out of process and the add-in waits. The
/// CLI is found through VBANG_CLI, next to the add-in, or in this repository's build output.
/// </summary>
internal static class ProjectBuilder
{
    private static readonly TimeSpan BuildTimeout = TimeSpan.FromMinutes(3);

    /// <summary>The output of the last build that ran, for the status command and the debug console.</summary>
    public static string LastOutput { get; private set; } = string.Empty;

    public static BuildOutcome EnsureBuilt(string projectDir)
    {
        if (!IsStale(projectDir))
        {
            return BuildOutcome.UpToDate;
        }

        var cli = FindCli();
        if (cli is null)
        {
            LastOutput = "vbang CLI not found (set VBANG_CLI); using the existing build output.";
            return BuildOutcome.NoCli;
        }

        var info = new ProcessStartInfo(cli)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        info.ArgumentList.Add("build");
        info.ArgumentList.Add("--project");
        info.ArgumentList.Add(projectDir);
        using var process = Process.Start(info);
        if (process is null)
        {
            LastOutput = "vbang could not be started: " + cli;
            return BuildOutcome.Failed;
        }

        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(BuildTimeout))
        {
            process.Kill(entireProcessTree: true);
            LastOutput = "vbang build did not finish within " + BuildTimeout;
            return BuildOutcome.Failed;
        }

        LastOutput = (output.Result + error.Result).Trim();
        return process.ExitCode == 0 ? BuildOutcome.Built : BuildOutcome.Failed;
    }

    /// <summary>
    /// True when out/build.json is missing, another version of vba-ng wrote it, or any source (.bas, .cls, .frm,
    /// vbang.json) was added, removed, or changed since. The add-in and the CLI ship as one version, so a project an
    /// older one built is built again after an upgrade.
    /// </summary>
    public static bool IsStale(string projectDir)
    {
        var infoPath = ProjectPaths.BuildInfoPath(projectDir);
        if (!File.Exists(infoPath) || !File.Exists(ProjectPaths.AssemblyPath(projectDir)))
        {
            return true;
        }

        Dictionary<string, string> recorded;
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(infoPath));
            if (document.RootElement.GetProperty("Compiler").GetString() != typeof(ProjectBuilder).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion)
            {
                return true;
            }

            recorded = document.RootElement.GetProperty("Inputs").EnumerateObject()
                .ToDictionary(p => p.Name, p => p.Value.GetString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException)
        {
            return true;
        }

        var current = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in Directory.GetFiles(projectDir, "*.bas").Concat(Directory.GetFiles(projectDir, "*.cls")).Concat(Directory.GetFiles(projectDir, "*.frm")))
        {
            current[Path.GetFileName(source)] = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(File.ReadAllText(source))));
        }

        var manifest = ProjectPaths.ManifestPath(projectDir);
        if (File.Exists(manifest))
        {
            current[Path.GetFileName(manifest)] = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(manifest)));
        }

        return current.Count != recorded.Count || current.Any(pair => !recorded.TryGetValue(pair.Key, out var hash) || !string.Equals(hash, pair.Value, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The folder of the .xll; the managed assembly has no location of its own when Excel-DNA loads it from the packed add-in.</summary>
    private static string AddInDirectory()
    {
        var xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
        var directory = string.IsNullOrEmpty(xll) ? null : Path.GetDirectoryName(xll);
        return directory ?? AppContext.BaseDirectory;
    }

    private static string? FindCli()
    {
        var configured = Environment.GetEnvironmentVariable("VBANG_CLI");
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
        {
            return configured;
        }

        var addInDir = AddInDirectory();
        var beside = Path.Combine(addInDir, "vbang.exe");
        if (File.Exists(beside))
        {
            return beside;
        }

        // The repository layout: src/VbaNg.AddIn/bin/<Configuration>/net10.0-windows next to src/VbaNg.Cli/bin/...
        var configuration = typeof(ProjectBuilder).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? "Debug";
        for (var directory = new DirectoryInfo(addInDir); directory is not null; directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, "src", "VbaNg.Cli", "bin", configuration, "net10.0-windows", "vbang.exe");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }
}
