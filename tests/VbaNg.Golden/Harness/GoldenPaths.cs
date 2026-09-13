namespace VbaNg.Golden.Harness;

/// <summary>Locations of the case files and goldens, found from the repository root.</summary>
public static class GoldenPaths
{
    public static string RepositoryRoot { get; } = FindRepositoryRoot();

    public static string CasesDirectory => Path.Combine(RepositoryRoot, "tests", "VbaNg.Golden", "Cases");

    public static string GoldensDirectory => Path.Combine(RepositoryRoot, "tests", "VbaNg.Golden", "Goldens");

    /// <summary>Every case file, ordered by name so runs and reports are deterministic.</summary>
    public static IReadOnlyList<string> CaseFiles() =>
        Directory.EnumerateFiles(CasesDirectory, "*" + CaseFile.Extension)
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();

    public static string CasePath(string area) => Path.Combine(CasesDirectory, area + CaseFile.Extension);

    public static string GoldenPath(string area) => Path.Combine(GoldensDirectory, area + ".json");

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "VbaNg.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Repository root (VbaNg.slnx) not found above " + AppContext.BaseDirectory);
    }
}
