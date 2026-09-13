using VbaNg.Golden.Harness;

using Xunit;

namespace VbaNg.Golden;

public sealed class CaseFileTests
{
    [Fact]
    public void Parse_SplitsBlocksNamesCasesAndReadsOptions()
    {
        const string text = """
            ' Strings: a file comment, not a case

            Option Compare Text

            ' First comment
            ' Upper case
            Dim s As String
            s = "a"
            ? UCase(s)

            ? 1 + 1
            ? 2 + 2
            """;

        var area = CaseFile.Parse("Strings", text);

        Assert.Equal("Strings", area.Name);
        Assert.Equal(["Option Compare Text"], area.Options);
        Assert.Equal(2, area.Cases.Count);

        Assert.Equal("Upper case", area.Cases[0].Name);
        Assert.Equal(7, area.Cases[0].Line);
        Assert.Equal(["Dim s As String", "s = \"a\"", "? UCase(s)"], area.Cases[0].Lines);
        Assert.Equal(["UCase(s)"], area.Cases[0].Expressions);

        Assert.Equal("1 + 1", area.Cases[1].Name);
        Assert.Equal(11, area.Cases[1].Line);
        Assert.Equal(["1 + 1", "2 + 2"], area.Cases[1].Expressions);
    }

    [Fact]
    public void Parse_AcceptsCrLfAndIndentedResultLines()
    {
        var area = CaseFile.Parse("A", "Dim x\r\n  ? x\r\n");

        var source = Assert.Single(area.Cases);
        Assert.Equal(["Dim x", "  ? x"], source.Lines);
        Assert.Equal(["x"], source.Expressions);
    }

    [Fact]
    public void Parse_NamesACaseWithoutCommentsOrResultsAfterItsFirstLine()
    {
        var area = CaseFile.Parse("A", "Dim x\nx = 1\n");

        Assert.Equal("Dim x", Assert.Single(area.Cases).Name);
    }

    [Fact]
    public void Parse_RejectsNonAsciiSource()
    {
        var ex = Assert.Throws<FormatException>(() => CaseFile.Parse("Strings", "? Len(\"café\")\n"));

        Assert.Equal("Strings.cases(1): character U+00E9 is not ASCII; write it with ChrW.", ex.Message);
    }

    [Fact]
    public void Parse_RejectsOptionsAfterTheFirstCase()
    {
        var ex = Assert.Throws<FormatException>(() => CaseFile.Parse("A", "? 1\n\nOption Base 1\n"));

        Assert.Equal("A.cases(3): Option lines must come before the first case.", ex.Message);
    }

    [Fact]
    public void Parse_ReadsDeclarationBlocksBeforeTheFirstCase()
    {
        const string text = "Option Explicit\n\nPublic Type T\n    X As Long\nEnd Type\n\n' helper\nPublic Function F() As Long\n    F = 1\nEnd Function\n\n? F()\n";

        var area = CaseFile.Parse("A", text);

        Assert.Equal(["Option Explicit"], area.Options);
        Assert.Equal(["Public Type T", "    X As Long", "End Type", "", "Public Function F() As Long", "    F = 1", "End Function"], area.Declarations);
        Assert.Equal("F()", Assert.Single(area.Cases).Name);
    }

    [Fact]
    public void Load_ReadsCompanionModulesNamedAfterTheArea()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "vbang-golden-tests", Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            File.WriteAllText(Path.Combine(dir, "A.cases"), "? 1\n");
            File.WriteAllText(Path.Combine(dir, "A.Counter.cls"), "VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Counter\"\r\nPublic Value As Long\r\n");
            File.WriteAllText(Path.Combine(dir, "A.Helpers.bas"), "Attribute VB_Name = \"Helpers\"\r\nPublic Total As Long\r\n");
            File.WriteAllText(Path.Combine(dir, "B.Other.cls"), "VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Other\"\r\n");

            var area = CaseFile.Load(Path.Combine(dir, "A.cases"));

            Assert.Equal(2, area.Companions.Count);
            var source = area.Companions[0];
            Assert.Equal("Counter", source.Name);
            Assert.True(source.IsClass);
            Assert.Equal(["VERSION 1.0 CLASS", "Attribute VB_Name = \"Counter\"", "Public Value As Long"], source.Lines);
            Assert.Equal("VERSION 1.0 CLASS\r\nAttribute VB_Name = \"Counter\"\r\nPublic Value As Long\r\n", RecorderModule.CompanionText(source));
            var module = area.Companions[1];
            Assert.Equal("Helpers.bas", module.FileName);
            Assert.False(module.IsClass);
            Assert.Equal(["Attribute VB_Name = \"Helpers\"", "Public Total As Long"], module.Lines);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Load_RejectsACompanionWhoseAttributeNamesAnotherModule()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "vbang-golden-tests", Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            File.WriteAllText(Path.Combine(dir, "A.cases"), "? 1\n");
            File.WriteAllText(Path.Combine(dir, "A.Helpers.bas"), "Attribute VB_Name = \"Helper\"\r\nPublic Total As Long\r\n");

            var ex = Assert.Throws<FormatException>(() => CaseFile.Load(Path.Combine(dir, "A.cases")));
            Assert.Contains("Attribute VB_Name = \"Helpers\"", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Parse_RejectsDeclarationsAfterTheFirstCase()
    {
        var ex = Assert.Throws<FormatException>(() => CaseFile.Parse("A", "? 1\n\nPublic Sub S()\nEnd Sub\n"));

        Assert.Equal("A.cases(3): declarations must come before the first case.", ex.Message);
    }

    [Fact]
    public void Parse_RejectsAreaNamesThatAreNotIdentifiers()
    {
        Assert.Throws<FormatException>(() => CaseFile.Parse("Bad-Name", "? 1\n"));
    }
}
