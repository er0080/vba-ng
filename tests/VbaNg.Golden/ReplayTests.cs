using System.Globalization;
using System.Text;

using VbaNg.Golden.Harness;
using VbaNg.Golden.Replay;
using VbaNg.Runtime;

using Xunit;

[assembly: AssemblyFixture(typeof(VbaNg.Golden.ReplayReport))]

namespace VbaNg.Golden;

/// <summary>
/// Collects the per-area pass rates (the M2 exit criterion, ROADMAP.md) and writes them to
/// golden-report.txt next to the test assembly when the run ends.
/// </summary>
public sealed class ReplayReport : IDisposable
{
    private readonly List<string> lines = [];
    private readonly List<string> leaks = [];
    private readonly Lock gate = new();

    /// <summary>
    /// Adds an area's line; <paramref name="tolerated"/> is how many of its failures Expected/Ulps.txt accepts as one ulp
    /// from VBA's (ARCHITECTURE.md D22), counted with the passes in the rate; <paramref name="leaked"/> is how many BSTRs
    /// and arrays the area left allocated, a number M7 drives to zero (ROADMAP.md WP2). The cases that leaked are listed
    /// in golden-leaks.txt.
    /// </summary>
    public void Add(string area, IReadOnlyList<CaseOutcome> outcomes, long leaked, int tolerated = 0)
    {
        ArgumentNullException.ThrowIfNull(outcomes);
        var pass = outcomes.Count(o => o.Verdict == CaseVerdict.Pass);
        var fail = outcomes.Count(o => o.Verdict == CaseVerdict.Fail) - tolerated;
        var unsupported = outcomes.Count(o => o.Verdict == CaseVerdict.Unsupported);
        var rate = outcomes.Count == 0 ? 0 : 100.0 * (pass + tolerated) / outcomes.Count;
        lock (gate)
        {
            lines.Add(string.Create(CultureInfo.InvariantCulture, $"{area,-14} {pass,5} pass {tolerated,3} ulp {fail,5} fail {unsupported,5} unsupported  {rate,6:F1} %  {leaked,6} left"));
            leaks.AddRange(outcomes.Where(o => o.Leaked != 0).Select(o => string.Create(CultureInfo.InvariantCulture, $"{area}: {o.Leaked,4}  {o.Name}")));
        }
    }

    public void Dispose()
    {
        var report = new StringBuilder();
        report.Append("Golden replay: pass rate per library area\n\n");
        lock (gate)
        {
            foreach (var line in lines.Order(StringComparer.Ordinal))
            {
                report.Append(line).Append('\n');
            }

            File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "golden-leaks.txt"), string.Join('\n', leaks.Order(StringComparer.Ordinal)) + "\n");
        }

        File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "golden-report.txt"), report.ToString());
    }
}

/// <summary>
/// Replays every recorded case against the runtime. A case may fail only when it is listed in
/// tests/VbaNg.Golden/Expected/&lt;Area&gt;.txt (a known gap, one case name per line), or when
/// Expected/Ulps.txt lists it and its only differences are Doubles or Singles one ulp from VBA's
/// (ARCHITECTURE.md D22); a listed case that passes must be removed from its list, so the lists
/// always say exactly what is missing.
/// </summary>
public sealed class ReplayTests(ReplayReport report, ITestOutputHelper output)
{
    public static TheoryData<string> Areas()
    {
        var data = new TheoryData<string>();
        foreach (var file in GoldenPaths.CaseFiles())
        {
            var area = Path.GetFileNameWithoutExtension(file);
            if (File.Exists(GoldenPaths.GoldenPath(area)))
            {
                data.Add(area);
            }
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(Areas))]
    public void Area_ReplaysAgainstTheRuntime(string area)
    {
        var golden = GoldenArea.Load(GoldenPaths.GoldenPath(area));
        var liveBefore = Bstr.LiveCount + VbaArray.LiveCount;
        var outcomes = CompiledReplay.Run(golden).ToList();

        // A case Expected/Ulps.txt names may miss VBA's Doubles or Singles by one ulp and by nothing else (D22).
        var ulpCases = LoadNames(ExpectedPath("Ulps.txt"));
        bool Tolerated(CaseOutcome o) => o.WithinUlp && ulpCases.Contains($"{area}: {o.Name}");
        report.Add(area, outcomes, Bstr.LiveCount + VbaArray.LiveCount - liveBefore, outcomes.Count(Tolerated));

        var expectedFailures = LoadNames(ExpectedPath(area + ".txt"));
        var unexpectedFailures = outcomes.Where(o => o.Verdict != CaseVerdict.Pass && !expectedFailures.Contains(o.Name) && !Tolerated(o)).ToList();
        var unexpectedPasses = outcomes.Where(o => o.Verdict == CaseVerdict.Pass && expectedFailures.Contains(o.Name)).ToList();
        var exactNow = outcomes.Where(o => o.Verdict == CaseVerdict.Pass && ulpCases.Contains($"{area}: {o.Name}")).ToList();

        // Every case leaves the live-allocation count where it found it (ROADMAP.md M7 B5), except the ones Expected/Leaks.txt names.
        var expectedLeaks = LoadNames(ExpectedPath("Leaks.txt"));
        var unexpectedLeaks = outcomes.Where(o => o.Leaked != 0 && !expectedLeaks.Contains($"{area}: {o.Name}")).ToList();
        var unexpectedTight = outcomes.Where(o => o.Leaked == 0 && expectedLeaks.Contains($"{area}: {o.Name}")).ToList();

        foreach (var outcome in outcomes.Where(o => o.Verdict != CaseVerdict.Pass))
        {
            output.WriteLine(outcome.ToString());
        }

        var actualPath = Path.Combine(AppContext.BaseDirectory, "expected", area + ".txt");
        Directory.CreateDirectory(Path.GetDirectoryName(actualPath)!);
        File.WriteAllLines(actualPath, outcomes.Where(o => o.Verdict != CaseVerdict.Pass && !Tolerated(o)).Select(o => o.Name));

        var message = new StringBuilder();
        if (unexpectedFailures.Count > 0)
        {
            message.Append(CultureInfo.InvariantCulture, $"{area}: {unexpectedFailures.Count} case(s) fail that Expected/{area}.txt does not list:\n");
            foreach (var outcome in unexpectedFailures.Take(25))
            {
                message.Append("  ").Append(outcome).Append('\n');
            }
        }

        if (unexpectedPasses.Count > 0)
        {
            message.Append(CultureInfo.InvariantCulture, $"{area}: {unexpectedPasses.Count} case(s) now pass; remove them from Expected/{area}.txt:\n");
            foreach (var outcome in unexpectedPasses.Take(25))
            {
                message.Append("  ").Append(outcome.Name).Append('\n');
            }
        }

        if (exactNow.Count > 0)
        {
            message.Append(CultureInfo.InvariantCulture, $"{area}: {exactNow.Count} case(s) listed in Expected/Ulps.txt now match exactly; remove them:\n");
            foreach (var outcome in exactNow.Take(25))
            {
                message.Append("  ").Append(outcome.Name).Append('\n');
            }
        }

        if (unexpectedLeaks.Count > 0)
        {
            message.Append(CultureInfo.InvariantCulture, $"{area}: {unexpectedLeaks.Count} case(s) leave BSTRs or arrays allocated that Expected/Leaks.txt does not list:\n");
            foreach (var outcome in unexpectedLeaks.Take(25))
            {
                message.Append(CultureInfo.InvariantCulture, $"  {outcome.Leaked,4}  {outcome.Name}\n");
            }
        }

        if (unexpectedTight.Count > 0)
        {
            message.Append(CultureInfo.InvariantCulture, $"{area}: {unexpectedTight.Count} case(s) listed in Expected/Leaks.txt leave nothing allocated any more; remove them:\n");
            foreach (var outcome in unexpectedTight.Take(25))
            {
                message.Append("  ").Append(outcome.Name).Append('\n');
            }
        }

        Assert.True(message.Length == 0, message.ToString());
    }

    private static string ExpectedPath(string file) => Path.Combine(GoldenPaths.RepositoryRoot, "tests", "VbaNg.Golden", "Expected", file);

    private static HashSet<string> LoadNames(string path)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (!File.Exists(path))
        {
            return names;
        }

        foreach (var line in File.ReadLines(path))
        {
            var trimmed = line.Trim();
            if (trimmed.Length > 0 && !trimmed.StartsWith('#'))
            {
                names.Add(trimmed);
            }
        }

        return names;
    }
}
