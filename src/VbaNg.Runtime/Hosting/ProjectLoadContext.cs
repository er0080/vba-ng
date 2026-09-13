using System.Reflection;
using System.Runtime.Loader;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Collectible load context for one compiled project (ARCHITECTURE.md section 7, "Hot reload").
/// Resolves <c>VbaNg.Runtime</c> to the copy already loaded by the host, so <see cref="Host"/>
/// state is shared between the host and the project instead of duplicated per context.
/// Everything else falls back to the default context.
/// </summary>
public sealed class ProjectLoadContext : AssemblyLoadContext
{
    private static readonly Assembly RuntimeAssembly = typeof(ProjectLoadContext).Assembly;
    private static readonly string RuntimeAssemblyName = RuntimeAssembly.GetName().Name!;

    public ProjectLoadContext(string name)
        : base(name, isCollectible: true)
    {
    }

    protected override Assembly? Load(AssemblyName assemblyName)
    {
        ArgumentNullException.ThrowIfNull(assemblyName);
        return string.Equals(assemblyName.Name, RuntimeAssemblyName, StringComparison.OrdinalIgnoreCase)
            ? RuntimeAssembly
            : null;
    }
}
