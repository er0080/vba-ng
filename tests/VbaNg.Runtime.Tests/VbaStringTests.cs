using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The tests that read <see cref="Bstr.LiveCount"/>, a process-wide number: they run after every
/// other collection and one at a time, so no test allocating strings on another thread moves it.
/// (A Fact's own DisableParallelization flag counts only under xunit's ParallelMode.All.)
/// </summary>
[CollectionDefinition(LiveBstrTests.Name, DisableParallelization = true)]
public sealed class LiveBstrTests
{
    public const string Name = "Live BSTRs";
}

/// <summary>
/// The BSTR behind a String (ARCHITECTURE.md section 5, D20): the byte-length prefix, odd byte
/// counts, the null BSTR, and the ownership rails, which the Memory and Strings goldens pin from
/// VBA's side; these check the runtime's own contract, the live-allocation count included.
/// </summary>
[Collection(LiveBstrTests.Name)]
public sealed class VbaStringTests
{
    [Fact]
    public void Alloc_HoldsTheCodeUnitsBehindABytePrefix()
    {
        var before = Bstr.LiveCount;
        var text = VbaString.Alloc("abc");
        try
        {
            Assert.False(text.IsNull);
            Assert.Equal(6, text.ByteLength);
            Assert.Equal(3, text.Length);
            Assert.Equal("abc", text.ToString());
            Assert.Equal(6, Bstr.ByteLength(text.Pointer));
            Assert.Equal(before + 1, Bstr.LiveCount);
        }
        finally
        {
            text.Free();
        }

        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void AllocBytes_KeepsAnOddByteCount()
    {
        var text = VbaString.AllocBytes([0x68, 0x00, 0x41]);
        try
        {
            Assert.Equal(3, text.ByteLength);
            Assert.Equal(1, text.Length);
            Assert.Equal("h", text.ToString());
            Assert.Equal([0x68, 0x00, 0x41], text.Bytes.ToArray());
        }
        finally
        {
            text.Free();
        }
    }

    [Fact]
    public void Null_ReadsAsEmptyAndOwnsNothing()
    {
        var text = VbaString.Null;
        Assert.True(text.IsNull);
        Assert.Equal(0, text.ByteLength);
        Assert.Equal(string.Empty, text.ToString());
        Assert.True(text.Chars.IsEmpty);
        text.Free();

        var empty = VbaString.Alloc(ReadOnlySpan<char>.Empty);
        try
        {
            Assert.False(empty.IsNull);
            Assert.Equal(0, empty.ByteLength);
            Assert.Equal(text, empty);
        }
        finally
        {
            empty.Free();
        }
    }

    [Fact]
    public void Copy_IsANewAllocationWithTheSameBytes()
    {
        var text = VbaString.Alloc("copy");
        var copy = text.Copy();
        try
        {
            Assert.NotEqual(text.Pointer, copy.Pointer);
            Assert.Equal(text, copy);
            Assert.Equal(text.GetHashCode(), copy.GetHashCode());
        }
        finally
        {
            text.Free();
            copy.Free();
        }
    }

    [Fact]
    public void Free_TwiceIsRefused()
    {
        var text = VbaString.Alloc("once");
        text.Free();
        Assert.Throws<InvalidOperationException>(text.Free);
    }

    [Fact]
    public void Variant_StringIsATemporaryUntilTheStatementEnds()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var value = Variant.FromString("temp");
        Assert.Equal("temp", value.AsString());
        Assert.Equal(before + 1, Bstr.LiveCount);

        // A store copies; the temporary goes with the statement, the slot keeps its own copy until released.
        var slot = Variant.Empty;
        ObjectRefs.Assign(ref slot, value);
        Assert.Equal(before + 2, Bstr.LiveCount);
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 1, Bstr.LiveCount);
        Assert.Equal("temp", slot.AsString());

        // Reassigning the slot to itself survives, and a null BSTR store frees the old one.
        ObjectRefs.Assign(ref slot, slot);
        Assert.Equal("temp", slot.AsString());
        Assert.Equal(before + 1, Bstr.LiveCount);
        ObjectRefs.Assign(ref slot, Variant.EmptyString);
        Assert.Equal(before, Bstr.LiveCount);
        Assert.Equal(string.Empty, slot.AsString());
        ObjectRefs.Release(ref slot);
        Assert.True(slot.IsEmpty);
    }

    [Fact]
    public void Variant_EmptyStringAndVbNullStringCompareEqual()
    {
        var mark = ObjectRefs.Mark();
        try
        {
            Assert.Equal(Variant.EmptyString, Variant.FromString(string.Empty));
            Assert.Equal(Variant.FromString("a"), Variant.FromString("a"));
            Assert.NotEqual(Variant.FromString("a"), Variant.FromString("b"));
        }
        finally
        {
            ObjectRefs.ReleaseTo(mark);
        }
    }
}
