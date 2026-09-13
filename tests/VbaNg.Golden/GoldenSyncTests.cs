using VbaNg.Golden.Harness;

using Xunit;

namespace VbaNg.Golden;

/// <summary>
/// Every case file has a golden recorded from the same sources. A case file edited without
/// regeneration fails here; an area whose golden has not been recorded yet is skipped as
/// pending-golden rather than guessed (CLAUDE.md R2, R16).
/// </summary>
public sealed class GoldenSyncTests
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

    [Theory]
    [MemberData(nameof(Areas))]
    public void Golden_MatchesItsCaseFile(string area)
    {
        var cases = CaseFile.Load(GoldenPaths.CasePath(area));
        var goldenPath = GoldenPaths.GoldenPath(area);
        Assert.SkipUnless(File.Exists(goldenPath), $"pending-golden: record {area} with `dotnet run --project tests/VbaNg.Golden.Regen -- {area}`.");

        var golden = GoldenArea.Load(goldenPath);
        var regenerate = $"stale golden: regenerate with `dotnet run --project tests/VbaNg.Golden.Regen -- {area}`.";

        Assert.True(cases.Options.SequenceEqual(golden.Options), $"{area}: module options differ; {regenerate}");
        Assert.True(cases.Declarations.SequenceEqual(golden.Declarations ?? []), $"{area}: module declarations differ; {regenerate}");
        var recordedCompanions = (golden.Classes ?? []).Select(c => (c.Name + ".cls", c.Source))
            .Concat((golden.Modules ?? []).Select(m => (m.Name + ".bas", m.Source)))
            .ToList();
        Assert.True(cases.Companions.Count == recordedCompanions.Count, $"{area}: {cases.Companions.Count} companion module(s) but {recordedCompanions.Count} recorded; {regenerate}");
        for (var i = 0; i < cases.Companions.Count; i++)
        {
            Assert.True(cases.Companions[i].FileName == recordedCompanions[i].Item1 && cases.Companions[i].Lines.SequenceEqual(recordedCompanions[i].Source), $"{area}: companion module {cases.Companions[i].FileName} changed; {regenerate}");
        }

        Assert.True(cases.Cases.Count == golden.Cases.Count, $"{area}: {cases.Cases.Count} cases but {golden.Cases.Count} recorded; {regenerate}");
        for (var i = 0; i < cases.Cases.Count; i++)
        {
            var source = cases.Cases[i];
            var recorded = golden.Cases[i];
            Assert.True(source.Name == recorded.Name, $"{area} case {i + 1}: '{source.Name}' was recorded as '{recorded.Name}'; {regenerate}");
            Assert.True(source.Lines.SequenceEqual(recorded.Source), $"{area} case '{source.Name}': the source changed; {regenerate}");
        }
    }

    [Theory]
    [MemberData(nameof(Areas))]
    public void Golden_EveryCaseRecordedAValueOrAnError(string area)
    {
        var goldenPath = GoldenPaths.GoldenPath(area);
        Assert.SkipUnless(File.Exists(goldenPath), $"pending-golden: record {area} with `dotnet run --project tests/VbaNg.Golden.Regen -- {area}`.");

        var golden = GoldenArea.Load(goldenPath);

        var silent = golden.Cases.Where(c => c.Results.Count == 0 && c.Error is null).Select(c => c.Name).ToList();
        Assert.Empty(silent);
    }
}
