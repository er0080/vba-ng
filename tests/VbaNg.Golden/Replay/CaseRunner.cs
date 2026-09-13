using System.Globalization;

using VbaNg.Golden.Harness;
using VbaNg.Runtime;

namespace VbaNg.Golden.Replay;

public enum CaseVerdict
{
    Pass,
    Fail,
    Unsupported,
}

/// <summary>
/// How a replayed case fared; <paramref name="Leaked"/> is how many BSTRs and arrays it left allocated after its frame
/// ended and its results were released, a number M7 drives to zero (ROADMAP.md WP2). <paramref name="WithinUlp"/> marks
/// a failure whose only differences are Double or Single results one ulp from VBA's, which Expected/Ulps.txt may
/// accept (ARCHITECTURE.md D22).
/// </summary>
public sealed record CaseOutcome(string Name, CaseVerdict Verdict, string? Detail, long Leaked = 0, bool WithinUlp = false)
{
    public override string ToString() => Detail is null ? $"{Verdict}: {Name}" : $"{Verdict}: {Name}: {Detail}";
}

/// <summary>Compares what a replayed case recorded with the golden: every value field by field, then the error, if any.</summary>
public static class CaseRunner
{
    public static CaseOutcome Compare(GoldenCase golden, IReadOnlyList<Variant> recorded, VbaException? failure)
    {
        ArgumentNullException.ThrowIfNull(golden);
        ArgumentNullException.ThrowIfNull(recorded);

        var expected = golden.Results;
        var count = Math.Min(expected.Count, recorded.Count);
        string? ulp = null;
        for (var i = 0; i < count; i++)
        {
            var actual = GoldenDescriber.Describe(recorded[i]);
            var difference = GoldenMatch.Compare(expected[i].Value, actual, $"result {i + 1}");
            if (difference is null)
            {
                continue;
            }

            if (OneUlpApart(expected[i].Value, actual))
            {
                ulp ??= difference + " (one ulp)";
                continue;
            }

            return Fail(difference);
        }

        if (expected.Count != recorded.Count)
        {
            return Fail($"recorded {expected.Count} values, replay produced {recorded.Count}" + Describe(failure));
        }

        if (golden.Error is null)
        {
            return failure is null ? Passed() : Fail("unexpected" + Describe(failure));
        }

        if (failure is null)
        {
            return Fail($"expected error {golden.Error.Number} ({golden.Error.Description}), replay succeeded");
        }

        if (failure.Number != golden.Error.Number || failure.Description != golden.Error.Description || (failure.Source ?? string.Empty) != golden.Error.Source)
        {
            return Fail($"expected error {golden.Error.Number} \"{golden.Error.Description}\" from {golden.Error.Source}, got {failure.Number} \"{failure.Description}\" from {failure.Source}");
        }

        return Passed();

        CaseOutcome Passed() => ulp is null ? new(golden.Name, CaseVerdict.Pass, null) : new(golden.Name, CaseVerdict.Fail, ulp, WithinUlp: true);

        CaseOutcome Fail(string detail) => new(golden.Name, CaseVerdict.Fail, detail);
    }

    /// <summary>Two Doubles or two Singles whose bits are neighbours: one ulp apart.</summary>
    public static bool OneUlpApart(GoldenValue expected, GoldenValue actual)
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);
        if (expected.Type != actual.Type || expected.Bits is null || actual.Bits is null)
        {
            return false;
        }

        return expected.Type switch
        {
            "Double" => Int128.Abs((Int128)Hex64(expected.Bits) - Hex64(actual.Bits)) == 1,
            "Single" => Math.Abs((long)Hex32(expected.Bits) - Hex32(actual.Bits)) == 1,
            _ => false,
        };
    }

    private static long Hex64(string hex) => long.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);

    private static int Hex32(string hex) => int.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);

    private static string Describe(VbaException? error) => error is null ? string.Empty : $" error {error.Number} \"{error.Description}\"";
}
