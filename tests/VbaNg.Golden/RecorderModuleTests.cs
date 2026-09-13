using VbaNg.Compiler.Syntax;
using VbaNg.Golden.Harness;

using Xunit;

namespace VbaNg.Golden;

public sealed class RecorderModuleTests
{
    public static TheoryData<string> Areas()
    {
        var data = new TheoryData<string>();
        foreach (var file in GoldenPaths.CaseFiles())
        {
            data.Add(Path.GetFileNameWithoutExtension(file));
        }

        return data;
    }

    [Fact]
    public void Generate_RewritesResultLinesAndWrapsEachCaseInAHandler()
    {
        var area = CaseFile.Parse("Test", "Option Base 1\n\n' Sum\nDim a As Integer, b As Integer\na = 1: b = 2\n? a + b\n\n? 1\n? 2\n");

        var module = RecorderModule.Generate(area, @"C:\temp\Test.jsonl");

        Assert.Equal(2, module.CaseLines.Count);
        Assert.DoesNotContain("\n", module.Text.Replace("\r\n", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);
        Assert.StartsWith("Option Base 1\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("Private Const GoldenPath As String = \"C:\\temp\\Test.jsonl\"\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("    Open GoldenPath For Output As #goldenFile: Close #goldenFile\r\n    GoldenCase1\r\n    If Err.Number <> 0 Then GoldenError 1, Err.Number, Err.Description, Err.Source: Err.Clear\r\n    GoldenCase2\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("    Open GoldenPath For Append As #goldenFile\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("    GoldenRecord 1, 0, a + b\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("    GoldenRecord 2, 0, 1\r\n    GoldenRecord 2, 1, 2\r\n", module.Text, StringComparison.Ordinal);
        Assert.Contains("    GoldenError 2, Err.Number, Err.Description, Err.Source\r\n", module.Text, StringComparison.Ordinal);

        var lines = module.Text.Split("\r\n");
        Assert.Equal("    Dim a As Integer, b As Integer", lines[module.CaseLines[0] - 1]);
        Assert.Equal("    GoldenRecord 2, 0, 1", lines[module.CaseLines[1] - 1]);

        var tree = SyntaxTree.Parse(module.Text, "Test.bas");
        Assert.Empty(tree.Diagnostics.Select(d => d.ToString()));
        Assert.Equal(module.Text, tree.Root.ToFullString());
    }

    [Fact]
    public void Validate_ReportsSyntaxErrorsAtTheCaseFileLine()
    {
        var area = CaseFile.Parse("Test", "? 1\n\n' Broken\nDim x\n? (1\n");

        var problems = RecorderModule.Validate(area);

        Assert.Equal(["Test.cases(5): case 'Broken': Expected: )"], problems);
    }

    [Fact]
    public void Validate_ReportsResultLinesWithMoreThanOneExpression()
    {
        var area = CaseFile.Parse("Test", "? 1, 2\n");

        var problems = RecorderModule.Validate(area);

        Assert.Equal(["Test.cases(1): case '1, 2': a ? line records exactly one expression."], problems);
    }

    [Theory]
    [MemberData(nameof(Areas))]
    public void CaseFile_GeneratesAModuleTheParserAccepts(string area)
    {
        var problems = RecorderModule.Validate(CaseFile.Load(GoldenPaths.CasePath(area)));

        Assert.Empty(problems);
    }
}
