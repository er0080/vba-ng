using Xunit;

namespace VbaNg.Import.Tests;

/// <summary>
/// The MS-OVBA reader against a real workbook (ARCHITECTURE.md section 10). Fixtures/Legacy.xlsm
/// was written by tools/New-ImportFixtureWorkbook.ps1 in a throwaway Excel; reading it needs no
/// Excel and no "trust access to the VBA project" setting (CLAUDE.md R9).
/// </summary>
public sealed class VbaProjectReaderTests
{
    private static VbaProject Fixture() =>
        VbaProjectReader.FromWorkbook(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Legacy.xlsm"))
        ?? throw new InvalidOperationException("The fixture workbook has no VBA project.");

    [Fact]
    public void Read_ReportsTheProjectNameAndCodePage()
    {
        var project = Fixture();

        Assert.Equal("LegacyBook", project.Name);
        Assert.Equal(1252, project.CodePage);
    }

    [Fact]
    public void Read_FindsEveryModuleWithItsKind()
    {
        var project = Fixture();

        Assert.Equal(VbaModuleKind.Standard, Module(project, "Sales").Kind);
        Assert.Equal(VbaModuleKind.Class, Module(project, "Basket").Kind);
        Assert.Equal(VbaModuleKind.Document, Module(project, "Sheet1").Kind);
        Assert.Equal(VbaModuleKind.Document, Module(project, "ThisWorkbook").Kind);
        Assert.Equal(VbaModuleKind.Form, Module(project, "Dialog").Kind);
        Assert.Equal("Sales.bas", Module(project, "Sales").FileName);
        Assert.Equal("Basket.cls", Module(project, "Basket").FileName);
        Assert.Equal("Dialog.frm", Module(project, "Dialog").FileName);
    }

    [Fact]
    public void Read_DecompressesTheSourceOfEveryModule()
    {
        var project = Fixture();

        var sales = Module(project, "Sales").Source;
        Assert.StartsWith("Attribute VB_Name = \"Sales\"\r\n", sales, StringComparison.Ordinal);
        Assert.Contains("Public Function Gross(ByVal net As Currency) As Currency\r\n", sales, StringComparison.Ordinal);
        Assert.Contains("Private Declare PtrSafe Sub Sleep Lib \"kernel32\"", sales, StringComparison.Ordinal);
        Assert.EndsWith("End Sub\r\n", sales, StringComparison.Ordinal);
        Assert.DoesNotContain('\0', sales);

        Assert.Contains("Private Sub Class_Initialize()", Module(project, "Basket").Source, StringComparison.Ordinal);
        Assert.Contains("Attribute VB_PredeclaredId = True", Module(project, "Sheet1").Source, StringComparison.Ordinal);
        Assert.Contains("Private Sub Worksheet_Change(ByVal Target As Range)", Module(project, "Sheet1").Source, StringComparison.Ordinal);
        Assert.Contains("Private Sub Workbook_Open()", Module(project, "ThisWorkbook").Source, StringComparison.Ordinal);
        Assert.Contains("Private Sub UserForm_Click()", Module(project, "Dialog").Source, StringComparison.Ordinal);
    }

    /// <summary>
    /// The references the dir stream records: a registered library keeps its moniker, and a control
    /// library (the MSForms behind a UserForm) arrives as an original plus an extended entry. VBA
    /// and the host's own library are implicit and are not recorded.
    /// </summary>
    [Fact]
    public void Read_ReportsTheReferencesWithTheirLibraryIdentity()
    {
        var project = Fixture();

        var scripting = Assert.Single(project.References, r => r.Name.Equals("Scripting", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("420b2830-e718-11cf-893d-00a0c9054228", scripting.LibraryGuid);
        Assert.Equal("1.0", scripting.Version);

        var stdole = Assert.Single(project.References, r => r.Name.Equals("stdole", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("00020430-0000-0000-c000-000000000046", stdole.LibraryGuid);
        Assert.Equal("2.0", stdole.Version);

        Assert.Contains(project.References, r => r.Name.Equals("Office", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(project.References, r => r.Name.Equals("MSForms", StringComparison.OrdinalIgnoreCase));
    }

    private static VbaModule Module(VbaProject project, string name) =>
        project.Modules.Single(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}
