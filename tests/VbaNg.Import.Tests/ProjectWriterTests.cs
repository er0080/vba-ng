using Xunit;

namespace VbaNg.Import.Tests;

/// <summary>
/// What an import writes (ARCHITECTURE.md section 10): the VBE's export format per module, a
/// manifest with the project's references, and a report of what did not come over.
/// </summary>
public sealed class ProjectWriterTests : IDisposable
{
    private readonly string workDir = Path.Combine(Path.GetTempPath(), "vbang-import-tests", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try
        {
            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private ImportResult Import()
    {
        var project = VbaProjectReader.FromWorkbook(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Legacy.xlsm"))!;
        return ProjectWriter.Write(project, Path.Combine(workDir, "Legacy.vbang"));
    }

    [Fact]
    public void Write_PutsEveryModuleInItsOwnFile()
    {
        var result = Import();

        Assert.Contains("Sales.bas", result.Files, StringComparer.Ordinal);
        Assert.Contains("Basket.cls", result.Files, StringComparer.Ordinal);
        Assert.Contains("Sheet1.cls", result.Files, StringComparer.Ordinal);
        Assert.Contains("Dialog.frm", result.Files, StringComparer.Ordinal);
        Assert.Contains("vbang.json", result.Files, StringComparer.Ordinal);
        Assert.All(result.Files, f => Assert.True(File.Exists(Path.Combine(result.ProjectDir, f)), f));
    }

    /// <summary>A class module's file starts with the header the VBE writes above the attributes, which the module stream does not carry.</summary>
    [Fact]
    public void Write_AddsTheClassHeaderToClassAndDocumentModules()
    {
        var result = Import();

        var basket = File.ReadAllText(Path.Combine(result.ProjectDir, "Basket.cls"));
        Assert.StartsWith("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Basket\"", basket, StringComparison.Ordinal);
        Assert.StartsWith("VERSION 1.0 CLASS\r\n", File.ReadAllText(Path.Combine(result.ProjectDir, "Sheet1.cls")), StringComparison.Ordinal);
        Assert.StartsWith("Attribute VB_Name = \"Sales\"\r\n", File.ReadAllText(Path.Combine(result.ProjectDir, "Sales.bas")), StringComparison.Ordinal);
    }

    /// <summary>The manifest names the host library and every referenced library that carries an identity (ARCHITECTURE.md section 3).</summary>
    [Fact]
    public void Write_ManifestNamesTheHostAndTheReferencedLibraries()
    {
        var result = Import();

        var manifest = File.ReadAllText(Path.Combine(result.ProjectDir, "vbang.json"));
        Assert.Contains("\"name\": \"LegacyBook\"", manifest, StringComparison.Ordinal);
        Assert.Contains("{ \"name\": \"Excel\" }", manifest, StringComparison.Ordinal);
        Assert.Contains("\"name\": \"Scripting\", \"guid\": \"420b2830-e718-11cf-893d-00a0c9054228\", \"version\": \"1.0\"", manifest, StringComparison.Ordinal);
        // Office is a default reference of every Excel project, not an implicit one: the manifest names it (ARCHITECTURE.md section 3).
        Assert.Contains("\"name\": \"Office\", \"guid\": \"2df8d04c-5bfa-101b-bde5-00aa0044de52\"", manifest, StringComparison.Ordinal);
        Assert.DoesNotContain("stdole", manifest, StringComparison.Ordinal);
    }

    /// <summary>
    /// The manifest names every document module the PROJECT stream listed, so the folder says what
    /// its .cls files are once it is in git without the workbook beside it (D19). Without kinds
    /// from the caller the name decides, which is the rule the compiler falls back to anyway.
    /// </summary>
    [Fact]
    public void Write_ManifestNamesTheDocumentModulesAndTheirKinds()
    {
        var result = Import();

        var manifest = File.ReadAllText(Path.Combine(result.ProjectDir, "vbang.json"));
        Assert.Contains("\"documents\": {", manifest, StringComparison.Ordinal);
        Assert.Contains("\"ThisWorkbook\": \"Workbook\"", manifest, StringComparison.Ordinal);
        Assert.Contains("\"Sheet1\": \"Worksheet\"", manifest, StringComparison.Ordinal);

        // A class module is not a document module, so it stays out of the map.
        Assert.DoesNotContain("\"Basket\"", manifest, StringComparison.Ordinal);

        using var json = System.Text.Json.JsonDocument.Parse(manifest);
        Assert.Equal("Workbook", json.RootElement.GetProperty("documents").GetProperty("ThisWorkbook").GetString());
    }

    /// <summary>
    /// A chart sheet's module is a document module the vbaProject.bin cannot tell from a sheet's,
    /// so the kinds its caller read out of the workbook win over the name rule.
    /// </summary>
    [Fact]
    public void Write_ManifestTakesTheDocumentKindsTheCallerSupplies()
    {
        var project = VbaProjectReader.FromWorkbook(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Legacy.xlsm"))!;
        var kinds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Sheet1"] = "Chart" };

        var result = ProjectWriter.Write(project, Path.Combine(workDir, "Kinds.vbang"), documentKinds: kinds);

        var manifest = File.ReadAllText(Path.Combine(result.ProjectDir, "vbang.json"));
        Assert.Contains("\"Sheet1\": \"Chart\"", manifest, StringComparison.Ordinal);
        Assert.Contains("\"ThisWorkbook\": \"Workbook\"", manifest, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_ReportsWhatDidNotComeOver()
    {
        var result = Import();

        Assert.Contains(result.Report, line => line.Contains("Dialog.frm", StringComparison.Ordinal) && line.Contains("UserForms", StringComparison.Ordinal));

        // A Declare comes over and builds, so the report of what did not come over leaves it out.
        Assert.DoesNotContain(result.Report, line => line.Contains("Declare", StringComparison.Ordinal));
        Assert.Contains("Imported 6 file(s)", result.ToText(), StringComparison.Ordinal);
    }
}
