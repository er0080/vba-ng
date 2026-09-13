using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// The workbook beside a project folder names its document modules (ARCHITECTURE.md D19): the
/// CodeNames are read from the OPC parts, with no Excel. The Sheets sample workbook has carried
/// an ActiveX control, so Excel wrote its sheet's codeName; the Quickstart workbook never has.
/// </summary>
public sealed class WorkbookCodeNamesTests
{
    private static string Sample(string folder, string file) => Path.Combine(TestPaths.RepositoryRoot, "samples", folder, file);

    [Fact]
    public void Read_SheetsWorkbook_NamesThisWorkbookAndSheet1()
    {
        var names = WorkbookCodeNames.Read(Sample("Sheets", "Sheets.xlsx"));

        Assert.Equal("Workbook", names["ThisWorkbook"]);
        Assert.Equal("Worksheet", names["sheet1"]);
        Assert.Equal(2, names.Count);
    }

    [Fact]
    public void Read_WorkbookThatNeverCarriedVba_NamesOnlyThisWorkbook()
    {
        var names = WorkbookCodeNames.Read(Sample("Quickstart", "Quickstart.xlsx"));

        Assert.Equal(["ThisWorkbook"], names.Keys);
    }

    [Fact]
    public void FindWorkbook_FollowsTheFolderName()
    {
        Assert.Equal(Sample("Sheets", "Sheets.xlsx"), WorkbookCodeNames.FindWorkbook(Sample("Sheets", "Sheets.vbang")));
        Assert.Null(WorkbookCodeNames.FindWorkbook(Sample("Classes", "Classes.vbang")));
    }

    [Fact]
    public void Read_NotAWorkbook_IsEmpty()
    {
        Assert.Empty(WorkbookCodeNames.Read(Sample("Hello", Path.Combine("Hello.vbang", "Hello.bas"))));
    }
}
