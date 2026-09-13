namespace VbaNg.Runtime.Hosting;

/// <summary>Folder and file conventions for a project (ARCHITECTURE.md section 3 and D4).</summary>
public static class ProjectPaths
{
    /// <summary>Suffix of a project folder: <c>Sales.xlsx</c> binds to <c>Sales.vbang/</c>.</summary>
    public const string FolderSuffix = ".vbang";

    /// <summary>Build output folder inside the project folder.</summary>
    public const string OutputFolderName = "out";

    /// <summary>Generated C# folder inside the output folder.</summary>
    public const string GeneratedFolderName = "gen";

    /// <summary>The project name: the folder name without the <c>.vbang</c> suffix.</summary>
    public static string ProjectName(string projectDir)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        var name = Path.GetFileName(Path.TrimEndingDirectorySeparator(Path.GetFullPath(projectDir)));
        return name.EndsWith(FolderSuffix, StringComparison.OrdinalIgnoreCase)
            ? name[..^FolderSuffix.Length]
            : name;
    }

    public static string OutputDir(string projectDir) =>
        Path.Combine(Path.GetFullPath(projectDir), OutputFolderName);

    public static string GeneratedDir(string projectDir) =>
        Path.Combine(OutputDir(projectDir), GeneratedFolderName);

    public static string AssemblyPath(string projectDir) =>
        Path.Combine(OutputDir(projectDir), ProjectName(projectDir) + ".dll");

    public static string SymbolsPath(string projectDir) =>
        Path.Combine(OutputDir(projectDir), ProjectName(projectDir) + ".pdb");

    public static string BuildInfoPath(string projectDir) =>
        Path.Combine(OutputDir(projectDir), "build.json");

    /// <summary>The project's log, written while Excel runs the project (see <see cref="ProjectLog"/>).</summary>
    public static string OutputLogPath(string projectDir) =>
        Path.Combine(OutputDir(projectDir), "output.log");

    /// <summary>The manifest, <c>vbang.json</c>, optional.</summary>
    public static string ManifestPath(string projectDir) =>
        Path.Combine(Path.GetFullPath(projectDir), "vbang.json");

    /// <summary>Where a host command writes a response too long for <c>Application.Run</c> to return (see <see cref="RunResponse"/>).</summary>
    public static string ResponsePath(string projectDir) =>
        Path.Combine(OutputDir(projectDir), "response.json");
}
