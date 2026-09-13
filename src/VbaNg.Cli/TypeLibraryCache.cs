using VbaNg.Compiler;
using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Cli;

/// <summary>The CLI's view of the type library cache in Interop: manifest references in, models out (ARCHITECTURE.md section 6).</summary>
internal static class TypeLibraryCache
{
    public static ComLibrary? Resolve(ManifestReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        try
        {
            return Interop.TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // A library the reader cannot read is a diagnostic (VBA0023 from the compiler), never a dead CLI.
            Console.Error.WriteLine($"vbang: could not read the type library '{reference.Name}': {ex.Message}");
            return null;
        }
    }
}
