using VbaNg.Golden.Harness;

using Xunit;

namespace VbaNg.Golden;

public sealed class GoldenDocumentTests
{
    private const string RecordedLines = """
        {"case":1,"index":0,"value":{"type":"Long","value":"32768"}}
        {"case":2,"error":{"number":6,"description":"Overflow","source":"VBAProject"}}
        {"case":3,"index":0,"value":{"type":"Variant()","bounds":"0 To 1","items":[{"type":"Integer","value":"1"},{"type":"String","value":"a\"b"}]}}
        {"case":3,"index":1,"value":{"type":"Double","value":"1.5","bits":"3FF8000000000000"}}
        {"case":3,"error":{"number":13,"description":"Type mismatch","source":"VBAProject"}}
        {"case":4,"index":0,"value":{"type":"String","value":"�","units":"D83D"}}
        {"done":true}
        """;

    private static CaseArea Area() => CaseFile.Parse("Test", "Dim v\nv = 32767\n? v + 1\n\n' Typed overflow\nDim i As Integer\ni = 32767\n? i + 1\n\n? Array(1, \"a\"\"b\")\n? 1.5\n? \"x\" + 1\n\n? 0\n");

    [Fact]
    public void Merge_AssignsValuesAndErrorsToTheirCases()
    {
        var golden = GoldenArea.Merge(Area(), GoldenArea.ReadRecordLines(RecordedLines), "16.0.20326", "en-US");

        Assert.Equal("Test", golden.Area);
        Assert.Equal("16.0.20326", golden.Excel);
        Assert.Equal("en-US", golden.Culture);
        Assert.Equal(4, golden.Cases.Count);

        var first = golden.Cases[0];
        Assert.Equal("v + 1", first.Name);
        Assert.Equal(["Dim v", "v = 32767", "? v + 1"], first.Source);
        var result = Assert.Single(first.Results);
        Assert.Equal(0, result.Index);
        Assert.Equal(new GoldenValue("Long", "32768"), result.Value);
        Assert.Null(first.Error);

        var second = golden.Cases[1];
        Assert.Equal("Typed overflow", second.Name);
        Assert.Empty(second.Results);
        Assert.Equal(new GoldenError(6, "Overflow", "VBAProject"), second.Error);

        var third = golden.Cases[2];
        Assert.Equal(2, third.Results.Count);
        var array = third.Results[0].Value;
        Assert.Equal("Variant()", array.Type);
        Assert.Equal("0 To 1", array.Bounds);
        Assert.Equal(["1", "a\"b"], array.Items!.Select(i => i.Value));
        Assert.Equal("3FF8000000000000", third.Results[1].Value.Bits);
        Assert.Equal(13, third.Error!.Number);

        var fourth = golden.Cases[3];
        var surrogate = Assert.Single(fourth.Results).Value;
        Assert.Equal("�", surrogate.Value);
        Assert.Equal("D83D", surrogate.Units);
        Assert.Null(fourth.Error);
    }

    [Fact]
    public void Merge_RequiresTheDoneMarker()
    {
        var lines = GoldenArea.ReadRecordLines("{\"case\":1,\"index\":0,\"value\":{\"type\":\"Empty\"}}\n");

        var ex = Assert.Throws<InvalidOperationException>(() => GoldenArea.Merge(Area(), lines, "16.0", "en-US"));

        Assert.Equal("Test: the recorder did not finish; the results end after 1 line(s).", ex.Message);
    }

    [Fact]
    public void Merge_RejectsCaseNumbersOutsideTheArea()
    {
        var lines = GoldenArea.ReadRecordLines("{\"case\":9,\"index\":0,\"value\":{\"type\":\"Empty\"}}\n{\"done\":true}\n");

        Assert.Throws<InvalidOperationException>(() => GoldenArea.Merge(Area(), lines, "16.0", "en-US"));
    }

    [Fact]
    public void SaveAndLoad_RoundTripWithDeterministicText()
    {
        var golden = GoldenArea.Merge(Area(), GoldenArea.ReadRecordLines(RecordedLines), "16.0.20326", "en-US");
        var path = Path.Combine(Path.GetTempPath(), "vbang-golden-tests", Path.GetRandomFileName() + ".json");
        try
        {
            golden.Save(path);
            var text = File.ReadAllText(path);
            var loaded = GoldenArea.Load(path);

            Assert.Equal(text, loaded.ToJson());
            Assert.DoesNotContain("\r", text, StringComparison.Ordinal);
            Assert.EndsWith("}\n", text, StringComparison.Ordinal);
            Assert.Contains("\"bounds\": \"0 To 1\"", text, StringComparison.Ordinal);
            Assert.Contains("\"description\": \"Overflow\"", text, StringComparison.Ordinal);
            Assert.DoesNotContain("\"bits\": null", text, StringComparison.Ordinal);
            Assert.Equal(golden.Cases.Select(c => c.Name), loaded.Cases.Select(c => c.Name));
            Assert.Equal(golden.Cases[2].Results[0].Value.Items!.Select(i => i.Value), loaded.Cases[2].Results[0].Value.Items!.Select(i => i.Value));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
