using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// Arrays as SAFEARRAYs (ARCHITECTURE.md section 5, "Arrays"; ROADMAP.md M7 C4): the descriptor
/// is one oleaut32 reads as VBA's (the bounds, the element type, and the feature flags the Memory
/// golden recorded), a Variant holds it under the VARIANT's tag with nothing managed, storage
/// owns it on D18's rails, and a For Each locks it. Each test checks the live counts, so the
/// class runs with the others that do.
/// </summary>
[Collection(LiveBstrTests.Name)]
public sealed partial class VbaArrayTests
{
    /// <summary>A fixed-length string array member lies inline as its characters (ROADMAP.md M7 C10): an element is read and written by its subscripts, padded or cut to its length, a subscript outside the bounds raises 9, and the locals window lists the elements.</summary>
    [Fact]
    public void InlineChars_ReadWriteAndListTheElementsOfAFixedLengthStringArray()
    {
        var storage = default(TestCharsArray);
        VbaArray.SetInlineChars(ref storage, [1], "ab");
        VbaArray.SetInlineChars(ref storage, [0], "wxyz");

        Assert.Equal("wxy", VbaArray.InlineChars(ref storage, [0]));
        Assert.Equal("ab ", VbaArray.InlineChars(ref storage, [1]));
        Assert.Equal(9, Assert.Throws<VbaException>(() => VbaArray.InlineChars(ref storage, [2])).Number);

        var elements = new EmbeddedCharsView<TestCharsArray>(storage).Elements.Cast<VbaArrayElementView>().ToList();
        Assert.Equal(["(0)", "(1)"], elements.Select(e => e.Subscripts));
        Assert.Contains("ab", elements[1].Value.ToString(), StringComparison.Ordinal);
    }

    /// <summary>The locals window lists an array's elements under their VBA subscripts (ARCHITECTURE.md section 9), read from the native storage the debugger cannot see into.</summary>
    [Fact]
    public void DebugView_ListsTheElementsUnderTheirSubscripts()
    {
        var array = VbaArray.Create(VarType.Long, [(1, 2), (5, 6)]);
        try
        {
            array.Set([2, 5], Variant.FromInt32(7));
            var proxyType = Type.GetType(typeof(VbaArray).GetCustomAttribute<DebuggerTypeProxyAttribute>()!.ProxyTypeName)!;
            var proxy = Activator.CreateInstance(proxyType, array)!;
            var elements = (Array)proxyType.GetProperty("Elements")!.GetValue(proxy)!;
            Assert.Equal(
                ["(1, 5) = 0 (Long)", "(2, 5) = 7 (Long)", "(1, 6) = 0 (Long)", "(2, 6) = 0 (Long)"],
                elements.Cast<object>().Select(e => e.GetType().GetProperty("Subscripts")!.GetValue(e) + " = " + e.GetType().GetProperty("Value")!.GetValue(e)));

            // A Variant holding the array expands to the same elements; one holding a scalar to nothing.
            var variantProxyType = Type.GetType(typeof(Variant).GetCustomAttribute<DebuggerTypeProxyAttribute>()!.ProxyTypeName)!;
            var contents = variantProxyType.GetProperty("Contents")!;
            Assert.Equal(4, ((Array)contents.GetValue(Activator.CreateInstance(variantProxyType, Variant.FromArray(array)))!).Length);
            Assert.Empty((Array)contents.GetValue(Activator.CreateInstance(variantProxyType, Variant.FromInt32(1)))!);
        }
        finally
        {
            ObjectRefs.Release(ref array);
        }
    }

    [Fact]
    public void Descriptor_ReadsAsOleAutomationReadsIt()
    {
        var array = VbaArray.Create(VarType.Long, [(1, 2), (5, 7)]);
        Assert.Equal(0, SafeArrayGetLBound(array.Descriptor, 1, out var lower1));
        Assert.Equal(0, SafeArrayGetUBound(array.Descriptor, 1, out var upper1));
        Assert.Equal(0, SafeArrayGetLBound(array.Descriptor, 2, out var lower2));
        Assert.Equal(0, SafeArrayGetUBound(array.Descriptor, 2, out var upper2));
        Assert.Equal((1, 2, 5, 7), (lower1, upper1, lower2, upper2));
        Assert.Equal((lower1, upper1, lower2, upper2), (array.LBound(1), array.UBound(1), array.LBound(2), array.UBound(2)));
        Assert.Equal(0, SafeArrayGetVartype(array.Descriptor, out var vt));
        Assert.Equal((ushort)VarType.Long, vt);
        Assert.Equal(4u, SafeArrayGetElemsize(array.Descriptor));

        // Column-major, as VBA lays a two-dimensional array out: a(2, 5) is the second element.
        array.Set([2, 5], Variant.FromInt32(42));
        var data = Marshal.ReadIntPtr(array.Descriptor, 16);
        Assert.Equal(42, Marshal.ReadInt32(data, 4));
        ObjectRefs.Release(ref array);
    }

    [Theory]
    [InlineData(VarType.Long, true, 0x92)]
    [InlineData(VarType.Long, false, 0x80)]
    [InlineData(VarType.Variant, true, 0x892)]
    [InlineData(VarType.String, true, 0x192)]
    public void FeatureFlags_AreTheOnesVbaSets(VarType elementType, bool fixedSize, int flags)
    {
        var array = VbaArray.Create(elementType, [(0, 2)], fixedSize);
        Assert.Equal(flags, (ushort)Marshal.ReadInt16(array.Descriptor, 2));
        ObjectRefs.Release(ref array);
    }

    [Fact]
    public void Variant_HoldsTheDescriptor_UnderTheVariantsTag()
    {
        Assert.False(RuntimeHelpers.IsReferenceOrContainsReferences<Variant>());
        var array = VbaArray.Create(VarType.Long, [(0, 1)]);
        var view = Variant.FromArray(array);
        Assert.Equal(VarType.Array, view.Type);
        Assert.Equal(VarType.Array | VarType.Long, view.VarTypeValue);
        Assert.Equal(array, view.AsArray());

        var unallocated = Variant.FromArray(VbaArray.Unallocated(VarType.String));
        Assert.True(unallocated.IsArray);
        Assert.False(unallocated.AsArray().IsAllocated);
        Assert.Equal(VarType.String, unallocated.AsArray().ElementType);
        ObjectRefs.Release(ref array);
    }

    [Fact]
    public void Storage_OwnsTheArray_EachStoreACopy()
    {
        var arrays = VbaArray.LiveCount;
        var strings = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var slot = Variant.Empty;
        ObjectRefs.Assign(ref slot, Variant.FromArray(VbaArray.FromValues(VarType.String, [Variant.FromString("a"), Variant.FromString("b")])));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(arrays + 1, VbaArray.LiveCount);
        Assert.Equal(strings + 2, Bstr.LiveCount);

        // A Variant array holding that array holds a copy of it, and a store of the Variant array copies both.
        var nested = Variant.Empty;
        ObjectRefs.Assign(ref nested, Variant.FromArray(VbaArray.FromValues([slot, Variant.FromInt32(1)])));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(arrays + 3, VbaArray.LiveCount);
        Assert.Equal("b", nested.AsArray().Get([0]).AsArray().Get([1]).AsVbaString().ToString());

        ObjectRefs.Release(ref nested);
        ObjectRefs.Release(ref slot);
        Assert.Equal(arrays, VbaArray.LiveCount);
        Assert.Equal(strings, Bstr.LiveCount);
    }

    [Fact]
    public void ReDimPreserve_KeepsThePrefix_AndTheSameBoundsKeepTheArrayInPlace()
    {
        var arrays = VbaArray.LiveCount;
        var array = VbaArray.Create(VarType.Long, [(1, 3)]);
        array.Set([2], Variant.FromInt32(7));
        var descriptor = array.Descriptor;
        VbaArray.ReDimDeclared(ref array, [(1, 3)], preserve: true);
        Assert.Equal(descriptor, array.Descriptor);

        VbaArray.ReDimDeclared(ref array, [(1, 5)], preserve: true);
        Assert.Equal(7, array.Get([2]).AsInt32());
        Assert.Equal(5, array.UBound());
        Assert.Equal(arrays + 1, VbaArray.LiveCount);

        // Preserve cannot move the lower bound, and the array is untouched when it raises.
        var kept = array;
        Assert.Equal(9, Assert.Throws<VbaException>(() => VbaArray.ReDimDeclared(ref kept, [(0, 5)], preserve: true)).Number);
        Assert.Equal(kept, array);

        VbaArray.Erase(ref array);
        Assert.False(array.IsAllocated);
        Assert.Equal(arrays, VbaArray.LiveCount);
    }

    [Fact]
    public void Erase_ResetsAFixedArray_AndDestroysADynamicOne()
    {
        var arrays = VbaArray.LiveCount;
        var strings = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var fixedArray = VbaArray.Create(VarType.String, [(0, 1)], fixedSize: true);
        fixedArray.Set([1], Variant.FromString("x"));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(strings + 1, Bstr.LiveCount);

        VbaArray.Erase(ref fixedArray);
        Assert.True(fixedArray.IsAllocated);
        Assert.True(fixedArray.Get([1]).AsVbaString().IsNull);
        Assert.Equal(strings, Bstr.LiveCount);
        Assert.Equal(10, Assert.Throws<VbaException>(() => VbaArray.ReDimDeclared(ref fixedArray, [(0, 3)], preserve: false)).Number);

        ObjectRefs.Release(ref fixedArray);
        Assert.Equal(arrays, VbaArray.LiveCount);
    }

    [Fact]
    public void ForEach_LocksTheArray_SoReDimAndEraseRaise10()
    {
        var arrays = VbaArray.LiveCount;
        var array = VbaArray.Create(VarType.Long, [(0, 1)]);
        var enumerator = ForEachEnumerator.Create(Variant.FromArray(array));
        Assert.True(enumerator.MoveNext());

        var walked = array;
        Assert.Equal(10, Assert.Throws<VbaException>(() => VbaArray.Erase(ref walked)).Number);
        Assert.Equal(10, Assert.Throws<VbaException>(() => VbaArray.ReDimDeclared(ref walked, [(0, 3)], preserve: false)).Number);
        Assert.True(enumerator.MoveNext());
        Assert.False(enumerator.MoveNext());

        // The loop is over, so the array lets go.
        VbaArray.Erase(ref array);
        Assert.Equal(arrays, VbaArray.LiveCount);
    }

    [LibraryImport("oleaut32.dll")]
    private static partial int SafeArrayGetLBound(nint array, uint dimension, out int bound);

    [LibraryImport("oleaut32.dll")]
    private static partial int SafeArrayGetUBound(nint array, uint dimension, out int bound);

    [LibraryImport("oleaut32.dll")]
    private static partial int SafeArrayGetVartype(nint array, out ushort vt);

    [LibraryImport("oleaut32.dll")]
    private static partial uint SafeArrayGetElemsize(nint array);
}

/// <summary>The inline characters of one String * 3, as the emitter declares them (__Chars3).</summary>
[InlineArray(6)]
internal struct TestChars
{
    public byte First;
}

/// <summary>A String * 3 (0 To 1) member's inline storage, as the emitter declares it (__Inline_String3_0To1).</summary>
[InlineArray(2)]
internal struct TestCharsArray : IEmbeddedArray
{
    public TestChars First;

    public static VarType ElementType => VarType.String;

    public static int[] Bounds => [0, 1];
}
