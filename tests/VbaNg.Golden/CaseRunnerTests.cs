using VbaNg.Golden.Harness;
using VbaNg.Golden.Replay;
using VbaNg.Runtime;

using Xunit;

namespace VbaNg.Golden;

/// <summary>
/// The replay's comparison (R2): exact by default. A failure whose only differences are Doubles
/// or Singles one ulp from VBA's is marked, so Expected/Ulps.txt can accept it (ARCHITECTURE.md
/// D22); anything more, or anything else beside it, fails outright.
/// </summary>
public sealed class CaseRunnerTests
{
    [Fact]
    public void Compare_Exact_Passes() =>
        Assert.Equal(CaseVerdict.Pass, CaseRunner.Compare(Case(Recorded(0.1)), [Variant.FromDouble(0.1)], null).Verdict);

    [Fact]
    public void Compare_OneUlpInADouble_FailsWithinOneUlp()
    {
        var outcome = CaseRunner.Compare(Case(Recorded(0.1)), [Variant.FromDouble(Math.BitIncrement(0.1))], null);

        Assert.Equal((CaseVerdict.Fail, true), (outcome.Verdict, outcome.WithinUlp));
    }

    [Fact]
    public void Compare_OneUlpInASingle_FailsWithinOneUlp()
    {
        var recorded = GoldenDescriber.Describe(Variant.FromSingle(0.1f));
        var outcome = CaseRunner.Compare(Case(recorded), [Variant.FromSingle(MathF.BitDecrement(0.1f))], null);

        Assert.Equal((CaseVerdict.Fail, true), (outcome.Verdict, outcome.WithinUlp));
    }

    [Fact]
    public void Compare_TwoUlps_FailsOutright()
    {
        var outcome = CaseRunner.Compare(Case(Recorded(0.1)), [Variant.FromDouble(Math.BitIncrement(Math.BitIncrement(0.1)))], null);

        Assert.Equal((CaseVerdict.Fail, false), (outcome.Verdict, outcome.WithinUlp));
    }

    [Fact]
    public void Compare_OneUlpBesideAnotherDifference_FailsOutright()
    {
        var outcome = CaseRunner.Compare(Case(Recorded(0.1), Recorded(2)), [Variant.FromDouble(Math.BitIncrement(0.1)), Variant.FromDouble(3)], null);

        Assert.Equal((CaseVerdict.Fail, false), (outcome.Verdict, outcome.WithinUlp));
    }

    private static GoldenValue Recorded(double value) => GoldenDescriber.Describe(Variant.FromDouble(value));

    private static GoldenCase Case(params GoldenValue[] values) =>
        new("case", [], values.Select((value, index) => new GoldenResult(index, value)).ToList(), null);
}
