using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>Finds the project folder: an explicit folder, the current folder if it is one, or the single <c>*.vbang</c> folder here.</summary>
internal static class ProjectLocator
{
    public static string? Resolve(string? explicitDir, out string? error)
    {
        error = null;
        if (explicitDir is not null)
        {
            var full = Path.GetFullPath(explicitDir);
            if (Directory.Exists(full))
            {
                return full;
            }

            error = $"Project folder not found: {full}";
            return null;
        }

        var current = Path.TrimEndingDirectorySeparator(Directory.GetCurrentDirectory());
        if (current.EndsWith(ProjectPaths.FolderSuffix, StringComparison.OrdinalIgnoreCase))
        {
            return current;
        }

        var candidates = Directory.GetDirectories(current, "*" + ProjectPaths.FolderSuffix);
        if (candidates.Length == 1)
        {
            return candidates[0];
        }

        error = candidates.Length == 0
            ? "No project folder found. Run inside a *.vbang folder or pass --project <folder>."
            : "Several *.vbang folders found here; pass --project <folder>.";
        return null;
    }
}
