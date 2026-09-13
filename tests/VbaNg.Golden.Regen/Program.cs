using System.Globalization;

using VbaNg.Golden.Harness;

namespace VbaNg.Golden.Regen;

/// <summary>
/// <c>dotnet run --project tests/VbaNg.Golden.Regen -- [Area ...]</c>: records every case of the
/// named areas (all areas by default) in a hidden throwaway Excel and rewrites the golden files.
/// Regeneration is deliberate: read the diff before committing (CLAUDE.md R16).
/// </summary>
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        var timeout = TimeSpan.FromSeconds(120);
        (int First, int Last)? range = null;
        var bisect = false;
        var validateOnly = false;
        var names = new List<string>();
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-h" or "--help" or "-?":
                    Console.WriteLine("usage: dotnet run --project tests/VbaNg.Golden.Regen -- [Area ...] [--timeout seconds] [--range first-last | --bisect]");
                    Console.WriteLine("Regenerates tests/VbaNg.Golden/Goldens/<Area>.json from the cases in tests/VbaNg.Golden/Cases.");
                    Console.WriteLine("Starts a hidden throwaway Excel; needs \"Trust access to the VBA project object model\".");
                    Console.WriteLine("--timeout: seconds to wait for an area before killing Excel (default 120).");
                    Console.WriteLine("--range: run only cases first..last of each area and print what happened; writes no golden.");
                    Console.WriteLine("--bisect: find the cases that stop an area (VBE compile errors) by halving; writes no golden.");
                    Console.WriteLine("--validate: check that the cases parse and stop; starts no Excel.");
                    return 2;
                case "--timeout" when i + 1 < args.Length && int.TryParse(args[i + 1], NumberStyles.None, CultureInfo.InvariantCulture, out var seconds):
                    timeout = TimeSpan.FromSeconds(seconds);
                    i++;
                    break;
                case "--range" when i + 1 < args.Length && TryParseRange(args[i + 1], out var parsed):
                    range = parsed;
                    i++;
                    break;
                case "--bisect":
                    bisect = true;
                    break;
                case "--validate":
                    validateOnly = true;
                    break;
                case var option when option.StartsWith('-'):
                    Console.Error.WriteLine($"regen: unknown or incomplete option {option}; see --help.");
                    return 2;
                default:
                    names.Add(args[i]);
                    break;
            }
        }

        var files = names.Count == 0 ? GoldenPaths.CaseFiles() : names.Select(GoldenPaths.CasePath).ToList();
        if (files.Count == 0)
        {
            Console.Error.WriteLine($"regen: no case files under {GoldenPaths.CasesDirectory}");
            return 2;
        }

        var areas = new List<CaseArea>();
        var invalid = false;
        foreach (var file in files)
        {
            if (!File.Exists(file))
            {
                Console.Error.WriteLine($"regen: case file not found: {file}");
                return 2;
            }

            CaseArea area;
            try
            {
                area = CaseFile.Load(file);
            }
            catch (FormatException ex)
            {
                Console.Error.WriteLine("regen: " + ex.Message);
                invalid = true;
                continue;
            }

            var problems = RecorderModule.Validate(area);
            foreach (var problem in problems)
            {
                Console.Error.WriteLine("regen: " + problem);
            }

            invalid |= problems.Count > 0;
            areas.Add(area);
        }

        if (invalid)
        {
            return 1;
        }

        if (validateOnly)
        {
            Console.WriteLine(Invariant($"regen: {areas.Count} area(s) validated; no Excel started."));
            return 0;
        }

        using var runner = new Runner(timeout);
        try
        {
            var failures = 0;
            foreach (var area in areas)
            {
                var ok = bisect ? Bisect(runner, area)
                    : range is { } r ? RunRange(runner, area, r)
                    : Record(runner, area);
                if (!ok)
                {
                    failures++;
                }
            }

            return failures == 0 ? 0 : 1;
        }
        catch (ProjectAccessException ex)
        {
            Console.Error.WriteLine($"regen: Excel refused access to the VBA project: {ex.Message}");
            Console.Error.WriteLine("regen: enable File > Options > Trust Center > Trust Center Settings > Macro Settings > \"Trust access to the VBA project object model\" and rerun.");
            return 3;
        }
        catch (InvalidOperationException ex)
        {
            Console.Error.WriteLine("regen: " + ex.Message);
            return 1;
        }
    }

    /// <summary>Records the whole area and writes its golden.</summary>
    private static bool Record(Runner runner, CaseArea area)
    {
        var golden = runner.Run(area);
        if (golden is null)
        {
            ReportTimeout(runner, area);
            return false;
        }

        golden.Save(GoldenPaths.GoldenPath(area.Name));
        Console.WriteLine(Invariant($"{area.Name}: {golden.Cases.Count} cases, {golden.Cases.Sum(c => c.Results.Count)} values, {golden.Cases.Count(c => c.Error is not null)} ended in a VBA error, written to {GoldenPaths.GoldenPath(area.Name)}"));
        ReportSilentCases(golden);
        return true;
    }

    /// <summary>Runs a slice of the area for triage and prints every case that ended in an error.</summary>
    private static bool RunRange(Runner runner, CaseArea area, (int First, int Last) range)
    {
        var slice = Slice(area, range.First, range.Last);
        var golden = runner.Run(slice);
        if (golden is null)
        {
            ReportTimeout(runner, slice);
            return false;
        }

        Console.WriteLine(Invariant($"{area.Name} cases {range.First}-{range.Last}: {golden.Cases.Count} ran, {golden.Cases.Sum(c => c.Results.Count)} values, {golden.Cases.Count(c => c.Error is not null)} ended in a VBA error (no golden written for a range)"));
        foreach (var failed in golden.Cases.Where(c => c.Error is not null))
        {
            Console.WriteLine(Invariant($"  error {failed.Error!.Number} in '{failed.Name}': {failed.Error.Description}"));
        }

        ReportSilentCases(golden);
        return true;
    }

    /// <summary>
    /// A compile error anywhere in the recorder module stops the whole area behind a VBE dialog,
    /// which this tool only sees as a timeout. Halving the case list until single cases remain
    /// names the culprits at a cost of one Excel start per failed half.
    /// </summary>
    private static bool Bisect(Runner runner, CaseArea area)
    {
        var culprits = new List<int>();
        BisectRange(runner, area, 1, area.Cases.Count, culprits);
        if (culprits.Count == 0)
        {
            Console.WriteLine($"{area.Name}: every case runs to completion.");
            return true;
        }

        foreach (var index in culprits)
        {
            var source = area.Cases[index - 1];
            Console.WriteLine(Invariant($"{area.Name}{CaseFile.Extension}({source.Line}): case '{source.Name}' stops the area (a VBE compile error or a dialog)."));
        }

        return false;
    }

    private static void BisectRange(Runner runner, CaseArea area, int first, int last, List<int> culprits)
    {
        Console.WriteLine(Invariant($"{area.Name}: trying cases {first}-{last}"));
        if (runner.Run(Slice(area, first, last)) is not null)
        {
            return;
        }

        if (first == last)
        {
            culprits.Add(first);
            return;
        }

        var middle = (first + last) / 2;
        BisectRange(runner, area, first, middle, culprits);
        BisectRange(runner, area, middle + 1, last, culprits);
    }

    private static CaseArea Slice(CaseArea area, int first, int last)
    {
        first = Math.Max(1, first);
        last = Math.Min(area.Cases.Count, last);
        return area with { Cases = area.Cases.Skip(first - 1).Take(last - first + 1).ToList() };
    }

    private static void ReportTimeout(Runner runner, CaseArea area) =>
        Console.Error.WriteLine(runner.LastOutcome == RunOutcome.Rejected
            ? Invariant($"regen: {area.Name}: Excel refused to run the recorder, so the project does not compile (a name that does not resolve, a Friend or Private member reached from outside); the module is at {runner.ModulePath(area)}. Run with --bisect to find the case.")
            : Invariant($"regen: {area.Name}: Excel did not return within {runner.Timeout.TotalSeconds:F0} s and was killed. A compile error or a modal dialog is the usual cause; the module is at {runner.ModulePath(area)}. Run with --bisect to find the case."));

    private static void ReportSilentCases(GoldenArea golden)
    {
        foreach (var quiet in golden.Cases.Where(c => c.Results.Count == 0 && c.Error is null))
        {
            Console.WriteLine($"  warning: case '{quiet.Name}' recorded nothing.");
        }
    }

    private static bool TryParseRange(string text, out (int First, int Last) range)
    {
        range = default;
        var parts = text.Split('-');
        if (parts.Length != 2
            || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var first)
            || !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var last)
            || first < 1 || last < first)
        {
            return false;
        }

        range = (first, last);
        return true;
    }

    private static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);
}

/// <summary>Excel refused access to the VBA project, which the recorder needs to inject its module.</summary>
internal sealed class ProjectAccessException(string message) : Exception(message);

/// <summary>
/// Runs recorder modules in one Excel after another: a session lives until an area times out and
/// Excel has to be killed, then the next run starts a fresh one.
/// </summary>
internal sealed class Runner(TimeSpan timeout) : IDisposable
{
    private readonly string scratch = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "vbang-golden")).FullName;
    private readonly string culture = CultureInfo.CurrentCulture.Name;
    private ExcelSession? excel;

    public TimeSpan Timeout { get; } = timeout;

    /// <summary>How the last run ended, so a refusal is reported as the compile error it is.</summary>
    public RunOutcome LastOutcome { get; private set; }

    public string ModulePath(CaseArea area) => Path.Combine(scratch, RecorderModule.ModuleName(area.Name) + ".bas");

    /// <summary>Records the area; null when Excel had to be killed. The module text stays on disk for inspection.</summary>
    public GoldenArea? Run(CaseArea area)
    {
        if (excel is null)
        {
            excel = ExcelSession.Start();
            if (!excel.TryOpenProject(out var reason))
            {
                excel.Dispose();
                excel = null;
                throw new ProjectAccessException(reason);
            }
        }

        var moduleName = RecorderModule.ModuleName(area.Name);
        var resultsPath = Path.Combine(scratch, area.Name + ".jsonl");
        File.Delete(resultsPath);
        var module = RecorderModule.Generate(area, resultsPath);
        File.WriteAllText(ModulePath(area), module.Text);
        // A class whose Implements names a class the project does not hold yet stays broken after the import, so interfaces go in first.
        var classFiles = new List<string>();
        foreach (var source in area.Companions.OrderBy(c => c.Lines.Any(l => l.TrimStart().StartsWith("Implements ", StringComparison.OrdinalIgnoreCase)) ? 1 : 0))
        {
            var companionPath = Path.Combine(scratch, source.FileName);
            File.WriteAllText(companionPath, RecorderModule.CompanionText(source));
            classFiles.Add(companionPath);
        }

        var version = excel.Version;
        LastOutcome = excel.RunModule(moduleName, module.Text, classFiles, Timeout);
        if (LastOutcome != RunOutcome.Completed)
        {
            // A refusal leaves the project in a state the next slice cannot trust either, so both outcomes start over.
            excel.Dispose();
            excel = null;
            return null;
        }

        IReadOnlyList<RecordLine> lines;
        try
        {
            lines = GoldenArea.ReadRecordLines(File.ReadAllText(resultsPath));
        }
        catch (System.Text.Json.JsonException ex)
        {
            // A recorder bug, not a case problem: keep the output for inspection and fail loudly.
            throw new InvalidOperationException($"{area.Name}: the recorder wrote a line that is not valid JSON ({ex.Message}); see {resultsPath}.", ex);
        }

        File.Delete(resultsPath);
        return GoldenArea.Merge(area, lines, version, culture);
    }

    public void Dispose()
    {
        excel?.Dispose();
        excel = null;
    }
}
