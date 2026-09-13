using System.IO.Compression;
using System.Xml;

using VbaNg.Runtime.Hosting;

namespace VbaNg.Compiler;

/// <summary>
/// The CodeNames of the workbook next to a project folder (ARCHITECTURE.md D19, section 4), read
/// from the OPC parts with no Excel: <c>xl/workbook.xml</c> carries ThisWorkbook's under
/// <c>workbookPr</c>, and each sheet or chart sheet part carries its own under <c>sheetPr</c> once
/// Excel has assigned one. A <c>.cls</c> whose VB_Name is one of them is the document module of
/// that object; the manifest's <c>documents</c> map overrides, and the attribute pairing is the
/// last resort when neither the workbook nor the manifest says (<see cref="ProjectManifest.IsDocumentModule"/>).
/// </summary>
public static class WorkbookCodeNames
{
    /// <summary>The workbook formats that are OPC packages; the .xls family is not, and its CodeNames come with the importer's manifest.</summary>
    private static readonly string[] Extensions = [".xlsx", ".xlsm", ".xlsb", ".xlam", ".xltx", ".xltm"];

    /// <summary>The workbook a project folder belongs to (D4): the folder's name with a workbook extension, in the same directory; null when there is none.</summary>
    public static string? FindWorkbook(string projectDir)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        var fullDir = Path.GetFullPath(projectDir);
        var parent = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(fullDir));
        if (parent is null)
        {
            return null;
        }

        var name = ProjectPaths.ProjectName(fullDir);
        return Extensions.Select(extension => Path.Combine(parent, name + extension)).FirstOrDefault(File.Exists);
    }

    /// <summary>The CodeNames of the workbook beside a project folder, by document kind; null when there is no workbook, which leaves the attribute rule in charge.</summary>
    public static IReadOnlyDictionary<string, string>? ReadBeside(string projectDir) =>
        FindWorkbook(projectDir) is { } workbook ? Read(workbook) : null;

    /// <summary>
    /// CodeName to document kind, <c>"Workbook"</c>, <c>"Worksheet"</c>, or <c>"Chart"</c>. ThisWorkbook
    /// is always present, since Excel reports it whether or not the part names it. A sheet part
    /// without a <c>codeName</c> (a workbook that never carried VBA) contributes nothing, and the
    /// binary parts of an .xlsb are not read, so such workbooks leave the attribute pairing in
    /// charge. A file that cannot be opened reads as empty.
    /// </summary>
    public static IReadOnlyDictionary<string, string> Read(string workbookPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var archive = ZipFile.OpenRead(workbookPath);
            var workbook = archive.GetEntry("xl/workbook.xml");
            if (workbook is null)
            {
                return names;
            }

            names[CodeNameOf(workbook, "workbookPr") ?? "ThisWorkbook"] = "Workbook";
            foreach (var entry in archive.Entries)
            {
                var kind = entry.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) ? "Worksheet"
                    : entry.FullName.StartsWith("xl/chartsheets/", StringComparison.OrdinalIgnoreCase) ? "Chart"
                    : null;
                if (kind is null || !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || entry.FullName.Contains("/_rels/", StringComparison.Ordinal))
                {
                    continue;
                }

                if (CodeNameOf(entry, "sheetPr") is { } codeName)
                {
                    names.TryAdd(codeName, kind);
                }
            }
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or XmlException)
        {
            names.Clear();
        }

        return names;
    }

    /// <summary>The codeName attribute of the first element of the given local name, in any namespace.</summary>
    private static string? CodeNameOf(ZipArchiveEntry entry, string elementName)
    {
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == elementName)
            {
                var codeName = reader.GetAttribute("codeName");
                return string.IsNullOrEmpty(codeName) ? null : codeName;
            }
        }

        return null;
    }
}
