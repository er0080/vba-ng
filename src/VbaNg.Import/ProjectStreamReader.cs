namespace VbaNg.Import;

/// <summary>
/// The <c>PROJECT</c> stream (MS-OVBA 2.3.1): plain text that lists what kind of module each name
/// is, which is the only place a class module is told from a document module.
/// </summary>
public static class ProjectStreamReader
{
    /// <summary>The kind of every module the stream names, by module name.</summary>
    public static Dictionary<string, VbaModuleKind> ModuleKinds(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var kinds = new Dictionary<string, VbaModuleKind>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in text.Split('\n'))
        {
            var trimmed = line.Trim('\r', ' ', '\t');
            var split = trimmed.IndexOf('=', StringComparison.Ordinal);
            if (split <= 0)
            {
                continue;
            }

            var keyword = trimmed[..split].Trim();
            var value = trimmed[(split + 1)..].Trim();

            // A Document line carries the module name and its cookie: Document=Sheet1/&H00000000.
            var name = value.Split('/')[0].Trim();
            if (name.Length == 0)
            {
                continue;
            }

            switch (keyword.ToUpperInvariant())
            {
                case "MODULE":
                    kinds[name] = VbaModuleKind.Standard;
                    break;
                case "CLASS":
                    kinds[name] = VbaModuleKind.Class;
                    break;
                case "DOCUMENT":
                    kinds[name] = VbaModuleKind.Document;
                    break;
                case "BASECLASS":
                    kinds[name] = VbaModuleKind.Form;
                    break;
            }
        }

        return kinds;
    }
}
