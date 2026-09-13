using System.Globalization;
using System.Security.Cryptography;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Loads built project assemblies into collectible contexts and reloads a project when its
/// assembly on disk changes (ARCHITECTURE.md section 7, "Hot reload").
/// Assemblies are shadow-copied before loading so a rebuild can overwrite the output folder
/// while the previous version is still loaded.
/// </summary>
public sealed class ProjectHost : IDisposable
{
    private readonly Dictionary<string, LoadedProject> projects = new(StringComparer.OrdinalIgnoreCase);
    private readonly string shadowRoot;

    public ProjectHost()
        : this(Path.Combine(Path.GetTempPath(), "vbang", "shadow"))
    {
    }

    public ProjectHost(string shadowRoot)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shadowRoot);
        this.shadowRoot = shadowRoot;
    }

    public IReadOnlyCollection<LoadedProject> Projects => projects.Values;

    /// <summary>
    /// Returns the loaded project for a project folder, loading it on first use and reloading it
    /// when the built assembly has changed since it was loaded.
    /// </summary>
    public LoadedProject Load(string projectDir)
    {
        var fullDir = Path.GetFullPath(projectDir);
        var assemblyPath = ProjectPaths.AssemblyPath(fullDir);
        if (!File.Exists(assemblyPath))
        {
            throw new FileNotFoundException($"Project has not been built; expected {assemblyPath}.", assemblyPath);
        }

        var stamp = HashFile(assemblyPath);
        if (projects.TryGetValue(fullDir, out var existing))
        {
            if (string.Equals(existing.Stamp, stamp, StringComparison.Ordinal))
            {
                return existing;
            }

            Unload(fullDir);
        }

        var name = ProjectPaths.ProjectName(fullDir);
        var shadowDir = Path.Combine(shadowRoot, name, Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(shadowDir);

        var shadowAssembly = Path.Combine(shadowDir, Path.GetFileName(assemblyPath));
        File.Copy(assemblyPath, shadowAssembly, overwrite: true);

        var symbolsPath = ProjectPaths.SymbolsPath(fullDir);
        if (File.Exists(symbolsPath))
        {
            File.Copy(symbolsPath, Path.Combine(shadowDir, Path.GetFileName(symbolsPath)), overwrite: true);
        }

        var context = new ProjectLoadContext(name);
        var assembly = context.LoadFromAssemblyPath(shadowAssembly);
        var loaded = new LoadedProject(name, fullDir, assembly, context, stamp);
        projects[fullDir] = loaded;
        return loaded;
    }

    public void Unload(string projectDir)
    {
        if (projects.Remove(Path.GetFullPath(projectDir), out var project))
        {
            project.Unload();
        }
    }

    public void UnloadAll()
    {
        foreach (var project in projects.Values)
        {
            project.Unload();
        }

        projects.Clear();
    }

    public void Dispose() => UnloadAll();

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
