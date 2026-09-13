using System.Globalization;
using System.Text;

namespace VbaNg.Import;

/// <summary>What a module is (MS-OVBA 2.3.4.2.3.2.9 MODULETYPE, and the PROJECT stream's lines).</summary>
public enum VbaModuleKind
{
    /// <summary>A standard module, written as a .bas file.</summary>
    Standard,

    /// <summary>A class module, written as a .cls file.</summary>
    Class,

    /// <summary>A module bound to a host object (a sheet, the workbook), written as a .cls file.</summary>
    Document,

    /// <summary>A UserForm, written as a .frm file with its .frx resources.</summary>
    Form,
}

/// <summary>One module of the project: its name, kind, and source as the VBE would export it.</summary>
public sealed record VbaModule(string Name, VbaModuleKind Kind, string Source)
{
    /// <summary>The file name the importer writes, following the VBE's own extensions.</summary>
    public string FileName => Name + Kind switch
    {
        VbaModuleKind.Standard => ".bas",
        VbaModuleKind.Form => ".frm",
        _ => ".cls",
    };
}

/// <summary>A library the project references (MS-OVBA 2.3.4.2.2 REFERENCE records).</summary>
public sealed record VbaReference(string Name, string? LibraryId)
{
    /// <summary>The type library identifier of the moniker (<c>*G{guid}#major.minor#lcid#path#description</c>).</summary>
    public string? LibraryGuid
    {
        get
        {
            var moniker = Part(0);
            var open = moniker?.IndexOf('{', StringComparison.Ordinal) ?? -1;
            return open >= 0 && Guid.TryParse(moniker![open..], out var value) ? value.ToString("D", CultureInfo.InvariantCulture) : null;
        }
    }

    /// <summary>The version the moniker names, as "major.minor".</summary>
    public string? Version => Part(1);

    private string? Part(int index)
    {
        var parts = LibraryId?.Split('#');
        return parts is not null && parts.Length > index ? parts[index] : null;
    }
}

/// <summary>The VBA project of a workbook, as MS-OVBA stores it.</summary>
public sealed record VbaProject(string Name, int CodePage, IReadOnlyList<VbaReference> References, IReadOnlyList<VbaModule> Modules);

/// <summary>
/// Reads a <c>vbaProject.bin</c> (ARCHITECTURE.md section 10): the <c>dir</c> stream for the
/// project's name, code page, references, and modules, the <c>PROJECT</c> stream for what kind of
/// module each one is, and every module stream for its source. No Excel, and no "trust access to
/// the VBA project" setting.
/// </summary>
public static class VbaProjectReader
{
    /// <summary>Reads the project from the storage that holds it, normally the <c>VBA</c> storage of a vbaProject.bin.</summary>
    public static VbaProject Read(CompoundFile file, string vbaStorage = "VBA")
    {
        ArgumentNullException.ThrowIfNull(file);
        var dir = DirStreamReader.Read(VbaCompression.Decompress(file.ReadStream(vbaStorage + "/dir")));
        var encoding = Encodings.ForCodePage(dir.CodePage);
        var kinds = file.Exists("PROJECT") ? ProjectStreamReader.ModuleKinds(encoding.GetString(file.ReadStream("PROJECT"))) : [];

        var modules = new List<VbaModule>();
        foreach (var record in dir.Modules)
        {
            var streamPath = vbaStorage + "/" + record.StreamName;
            if (!file.Exists(streamPath))
            {
                continue;
            }

            var stream = file.ReadStream(streamPath);
            var source = record.TextOffset < stream.Length
                ? encoding.GetString(VbaCompression.Decompress(stream, record.TextOffset))
                : string.Empty;
            var kind = kinds.GetValueOrDefault(record.Name, record.IsDocument ? VbaModuleKind.Class : VbaModuleKind.Standard);
            modules.Add(new VbaModule(record.Name, kind, Normalize(source)));
        }

        return new VbaProject(dir.Name, dir.CodePage, dir.References, modules);
    }

    /// <summary>Reads the project from a workbook file: the .xlsm and .xlsb families keep the project as an OPC part, the .xls family in the workbook's own compound file.</summary>
    public static VbaProject? FromWorkbook(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var extension = Path.GetExtension(path);
        if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xla", StringComparison.OrdinalIgnoreCase))
        {
            var workbook = CompoundFile.Open(path);
            return workbook.Exists("_VBA_PROJECT_CUR/VBA/dir") ? Read(workbook, "_VBA_PROJECT_CUR/VBA") : null;
        }

        using var archive = System.IO.Compression.ZipFile.OpenRead(path);
        var part = archive.Entries.FirstOrDefault(e => e.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        if (part is null)
        {
            return null;
        }

        using var stream = part.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return Read(CompoundFile.Read(memory.ToArray()));
    }

    /// <summary>Module source keeps the VBE's CRLF line ends and loses the trailing blank line the streams carry.</summary>
    private static string Normalize(string source) => source.ReplaceLineEndings("\r\n").TrimEnd('\r', '\n') + "\r\n";
}

/// <summary>The code pages MS-OVBA names, resolved once (the provider is not registered by default on .NET).</summary>
internal static class Encodings
{
    private static readonly Lock Gate = new();
    private static bool registered;

    public static Encoding ForCodePage(int codePage)
    {
        lock (Gate)
        {
            if (!registered)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                registered = true;
            }
        }

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
        {
            return Encoding.Latin1;
        }
    }
}
