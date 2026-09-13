using System.Globalization;

using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Interop;

/// <summary>
/// Resolves a manifest reference to a type library model (ARCHITECTURE.md section 6): a library
/// is read from the registry once per version and kept as JSON under the user profile
/// (<see cref="ComLibrary.CachePath"/>). A reference with a name and no guid resolves through
/// the table of libraries vba-ng knows.
/// </summary>
public static class TypeLibraryCache
{
    private static readonly Dictionary<string, (Guid Guid, int Major, int Minor)> WellKnown = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Excel"] = (new Guid("00020813-0000-0000-C000-000000000046"), 1, 9),
        ["Office"] = (new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"), 2, 8),
        ["Scripting"] = (new Guid("420B2830-E718-11CF-893D-00A0C9054228"), 1, 0),
        ["stdole"] = (new Guid("00020430-0000-0000-C000-000000000046"), 2, 0),
        ["VBIDE"] = (new Guid("0002E157-0000-0000-C000-000000000046"), 5, 3),
        ["MSForms"] = (new Guid("0D452EE1-E08F-101A-852E-02608C4D0BB4"), 2, 0),
    };

    /// <summary>The Excel type library, whose application object the add-in supplies.</summary>
    public static Guid ExcelLibraryId => WellKnown["Excel"].Guid;

    /// <summary>The model for a reference, or null when the library cannot be found; the caller reports VBA0023.</summary>
    public static ComLibrary? Resolve(string? name, string? libraryGuid, string? version)
    {
        if (!string.IsNullOrWhiteSpace(libraryGuid) && Guid.TryParse(libraryGuid, out var libraryId))
        {
            var (major, minor) = ParseVersion(version);
            if (Read(libraryId, major, minor) is { } found)
            {
                return found;
            }

            // A workbook with a UserForm references a type library generated for that document,
            // whose guid is the document's own and is registered nowhere: VBA reads it out of the
            // workbook, which vba-ng does not open. Its members are the library the name gives,
            // so the name decides when the guid resolves to nothing (ROADMAP.md WP6).
        }

        if (name is null || !WellKnown.TryGetValue(name, out var known))
        {
            return null;
        }

        var (wellKnownId, wellKnownMajor, wellKnownMinor) = known;
        if (!string.IsNullOrWhiteSpace(version))
        {
            (wellKnownMajor, wellKnownMinor) = ParseVersion(version);
        }

        return Read(wellKnownId, wellKnownMajor, wellKnownMinor);
    }

    /// <summary>The library from the cache, or read from the registry and cached; null when it is registered nowhere.</summary>
    private static ComLibrary? Read(Guid libraryId, int major, int minor)
    {
        var path = ComLibrary.CachePath(libraryId, major, minor);
        if (File.Exists(path))
        {
            try
            {
                var cached = ComLibrary.Load(path);
                if (cached.Format == ComLibrary.CurrentFormat)
                {
                    return cached;
                }

                // A file of an older shape lacks what the binder reads now; rebuilt below.
            }
            catch (Exception ex) when (ex is IOException or System.Text.Json.JsonException)
            {
                // A damaged cache file is rebuilt below.
            }
        }

        ComLibrary library;
        try
        {
            library = TypeLibraryReader.Read(libraryId, major, minor);
        }
        catch (FileNotFoundException)
        {
            return null;
        }

        try
        {
            library.Save(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // The cache is a convenience; the build goes on without it.
        }

        return library;
    }

    private static (int Major, int Minor) ParseVersion(string? version)
    {
        if (string.IsNullOrWhiteSpace(version))
        {
            return (1, 0);
        }

        var parts = version.Split('.');
        var major = int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var m) ? m : 1;
        var minor = parts.Length > 1 && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        return (major, minor);
    }
}
