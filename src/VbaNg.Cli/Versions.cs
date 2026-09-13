using System.Reflection;

namespace VbaNg.Cli;

/// <summary>
/// The versions vbang reports. Every one is read off an assembly at run time: the only place a
/// version is written down is Directory.Build.props, which a release edits once and CI overrides
/// with -p:Version. Nothing here, in the docs, or in a script repeats a version literal.
/// </summary>
internal static class Versions
{
    /// <summary>This build of vbang, as Directory.Build.props declared it.</summary>
    public static string Vbang { get; } = Of(typeof(Versions).Assembly);

    /// <summary>The .NET runtime this process is on.</summary>
    public static string Runtime => Environment.Version.ToString();

    /// <summary>The informational version of an assembly, falling back to its assembly version.</summary>
    public static string Of(Assembly assembly) =>
        assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? assembly.GetName().Version?.ToString()
        ?? "unknown";
}
