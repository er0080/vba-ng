using Xunit;

using VbaAssert = VbaNg.Runtime.Library.Assert;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The Assert module of test procedures (ARCHITECTURE.md section 8): a vba-ng addition with no
/// VBA counterpart, so unit tests rather than goldens define it. Equality is VBA's <c>=</c>.
/// </summary>
public sealed class AssertModuleTests
{
    [Fact]
    public void AreEqual_UsesVbaEquality()
    {
        VbaAssert.AreEqual(3, "3", Variant.Missing);
        VbaAssert.AreEqual(1.0, 1, Variant.Missing);
        VbaAssert.AreEqual(true, -1, Variant.Missing);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual(3, 1 + 1, Variant.Missing));
        Assert.Equal("Assert.AreEqual failed. Expected:<3>. Actual:<2>.", failure.Message);
    }

    [Fact]
    public void AreEqual_FollowsTheCompareMode()
    {
        VbaAssert.AreEqual("a", "A", Variant.Missing, CompareMode.Text);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual("a", "A", Variant.Missing, CompareMode.Binary));
        Assert.Equal("Assert.AreEqual failed. Expected:<a>. Actual:<A>.", failure.Message);
    }

    [Fact]
    public void AreEqual_TwoNullsAreEqualAndNullNeverEqualsAValue()
    {
        VbaAssert.AreEqual(Variant.Null, Variant.Null, Variant.Missing);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual(Variant.Null, 1, Variant.Missing));
        Assert.Equal("Assert.AreEqual failed. Expected:<Null>. Actual:<1>.", failure.Message);
    }

    [Fact]
    public void AreEqual_ComparesObjectsByReference()
    {
        var items = new Library.Collection();
        VbaAssert.AreEqual(Variant.FromObject(items), Variant.FromObject(items), Variant.Missing);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual(Variant.FromObject(items), Variant.FromObject(new Library.Collection()), Variant.Missing));
        Assert.Equal("Assert.AreEqual failed. Expected:<Collection>. Actual:<Collection>.", failure.Message);
    }

    [Fact]
    public void AreNotEqual_FailsOnEqualValues()
    {
        VbaAssert.AreNotEqual(1, 2, Variant.Missing);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreNotEqual(1, 1, "same"));
        Assert.Equal("Assert.AreNotEqual failed. Expected any value except:<1>. Actual:<1>. same", failure.Message);
    }

    [Fact]
    public void Message_IsAppendedWhenGiven()
    {
        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual(1, 2, "why"));
        Assert.Equal("Assert.AreEqual failed. Expected:<1>. Actual:<2>. why", failure.Message);

        var silent = Assert.Throws<AssertFailedException>(() => VbaAssert.AreEqual(1, 2, string.Empty));
        Assert.Equal("Assert.AreEqual failed. Expected:<1>. Actual:<2>.", silent.Message);
    }

    [Fact]
    public void IsTrueAndIsFalse_TakeTheConditionAsIfWould()
    {
        VbaAssert.IsTrue(-1, Variant.Missing);
        VbaAssert.IsTrue("1", Variant.Missing);
        VbaAssert.IsFalse(0, Variant.Missing);
        VbaAssert.IsFalse(Variant.Null, Variant.Missing);

        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.IsTrue(Variant.Null, Variant.Missing));
        Assert.Equal("Assert.IsTrue failed. Condition:<Null>.", failure.Message);
        Assert.Equal("Assert.IsFalse failed. Condition:<True>.", Assert.Throws<AssertFailedException>(() => VbaAssert.IsFalse(true, Variant.Missing)).Message);
    }

    [Fact]
    public void IsNothingAndIsNotNothing_CheckObjectReferences()
    {
        VbaAssert.IsNothing(Variant.Nothing, Variant.Missing);
        VbaAssert.IsNotNothing(Variant.FromObject(new Library.Collection()), Variant.Missing);

        Assert.Equal("Assert.IsNothing failed. Actual:<1>.", Assert.Throws<AssertFailedException>(() => VbaAssert.IsNothing(1, Variant.Missing)).Message);
        Assert.Equal("Assert.IsNotNothing failed. Actual:<Nothing>.", Assert.Throws<AssertFailedException>(() => VbaAssert.IsNotNothing(Variant.Nothing, Variant.Missing)).Message);
        Assert.Equal("Assert.IsNotNothing failed. Actual:<5>.", Assert.Throws<AssertFailedException>(() => VbaAssert.IsNotNothing(5, Variant.Missing)).Message);
    }

    [Fact]
    public void Fail_AlwaysThrows()
    {
        Assert.Equal("Assert.Fail failed. boom", Assert.Throws<AssertFailedException>(() => VbaAssert.Fail("boom")).Message);
        Assert.Equal("Assert.Fail failed.", Assert.Throws<AssertFailedException>(() => VbaAssert.Fail(Variant.Missing)).Message);
    }

    [Fact]
    public void Failures_AreNotVbaErrors()
    {
        var failure = Assert.Throws<AssertFailedException>(() => VbaAssert.Fail(Variant.Missing));

        Assert.Null(VbaException.From(failure));
    }
}
