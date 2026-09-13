using VbaNg.Runtime;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The invoke path and the marshaling, exercised through Scripting.Dictionary and
/// Scripting.FileSystemObject (scrrun.dll, part of Windows): every VBA type round-trips through a
/// Dictionary item, which stores the VARIANT it is given.
/// </summary>
public sealed class ScriptingTests
{
    private static readonly ComProvider Provider = new();

    /// <summary>A Variant holds a COM object as its pointer; reading it back as an object goes through the provider, installed as the hosts install it.</summary>
    public ScriptingTests() => Com.Provider = Provider;

    private static IDispatchObject Dictionary() => (IDispatchObject)Provider.CreateObject("Scripting.Dictionary", null);

    private static Variant Get(IDispatchObject target, string name, params Variant[] arguments) =>
        target.Invoke(target.GetDispId(name), InvokeKind.PropertyGet | InvokeKind.Method, arguments);

    private static void Call(IDispatchObject target, string name, params Variant[] arguments) =>
        target.Invoke(target.GetDispId(name), InvokeKind.Method, arguments);

    [Fact]
    public void Dictionary_MembersRoundTrip()
    {
        var dictionary = Dictionary();

        Call(dictionary, "Add", "a", 1);
        Call(dictionary, "add", "b", 2.5);

        Assert.Equal(Variant.FromInt32(2), Get(dictionary, "Count"));
        Assert.Equal(Variant.FromInt32(1), Get(dictionary, "Item", "a"));
        Assert.Equal(Variant.FromDouble(2.5), Get(dictionary, "Item", "b"));
        Assert.Equal(Variant.True, Get(dictionary, "Exists", "b"));
        Assert.Equal(Variant.False, Get(dictionary, "Exists", "z"));
        Assert.Equal("Dictionary", dictionary.TypeName);

        dictionary.Put(dictionary.GetDispId("Item"), ["a"], 10, asReference: false);
        Assert.Equal(Variant.FromInt32(10), Get(dictionary, "Item", "a"));

        Call(dictionary, "Remove", "a");
        Assert.Equal(Variant.FromInt32(1), Get(dictionary, "Count"));
    }

    [Fact]
    public void Dictionary_DefaultMemberIsItem()
    {
        var dictionary = Dictionary();
        Call(dictionary, "Add", "k", "v");

        Assert.Equal(Variant.FromString("v"), dictionary.Invoke(0, InvokeKind.PropertyGet | InvokeKind.Method, ["k"]));

        dictionary.Put(0, ["k"], "w", asReference: false);
        Assert.Equal(Variant.FromString("w"), Get(dictionary, "Item", "k"));
    }

    [Fact]
    public void Dictionary_ErrorsAreVbaErrors()
    {
        var dictionary = Dictionary();
        Call(dictionary, "Add", "a", 1);

        var duplicate = Assert.Throws<VbaException>(() => Call(dictionary, "Add", "a", 2));
        Assert.Equal(457, duplicate.Number);
        Assert.Equal("This key is already associated with an element of this collection", duplicate.Description);

        Assert.Equal(438, Assert.Throws<VbaException>(() => dictionary.GetDispId("NoSuchMember")).Number);
        Assert.Equal(450, Assert.Throws<VbaException>(() => Get(dictionary, "Exists")).Number);
        Assert.Equal(429, Assert.Throws<VbaException>(() => Provider.CreateObject("No.Such.ProgId", null)).Number);
    }

    [Fact]
    public void Dictionary_EnumeratesKeysAndReturnsThemAsAnArray()
    {
        var dictionary = Dictionary();
        Call(dictionary, "Add", "first", 1);
        Call(dictionary, "Add", "second", 2);

        var keys = new List<Variant>();
        using (var enumerator = dictionary.Enumerate())
        {
            while (enumerator.MoveNext())
            {
                // Current is the enumerator's until the next element; a caller that keeps it takes its own copy, as a loop variable does.
                keys.Add(ObjectRefs.Own(enumerator.Current));
            }
        }

        Assert.Equal([Variant.FromString("first"), Variant.FromString("second")], keys);

        var array = Get(dictionary, "Keys").AsArray();
        Assert.Equal(VarType.Variant, array.ElementType);
        Assert.Equal(0, array.LBound());
        Assert.Equal(1, array.UBound());
        Assert.Equal(Variant.FromString("second"), array.Get([1]));
    }

    public static TheoryData<Variant> ScalarValues() =>
    [
        Variant.Empty,
        Variant.Null,
        Variant.FromInt16(-7),
        Variant.FromInt32(123456),
        Variant.FromInt64(1L << 40),
        Variant.FromSingle(1.5f),
        Variant.FromDouble(Math.PI),
        Variant.FromCurrency(Currency.FromScaled(1234567)),
        Variant.FromDate(VbaDate.FromSerial(45000.75)),
        Variant.FromString("héllo"),
        Variant.FromString(string.Empty),
        Variant.True,
        Variant.False,
        Variant.FromByte(200),
        Variant.FromDecimal(-12345.6789m),
        Variant.FromError(new ErrorValue(unchecked((int)0x800A07D2))),
    ];

    [Theory]
    [MemberData(nameof(ScalarValues))]
    public void Dictionary_StoresEveryScalarTypeUnchanged(Variant value)
    {
        var dictionary = Dictionary();

        dictionary.Put(dictionary.GetDispId("Item"), ["k"], value, asReference: false);

        var stored = Get(dictionary, "Item", "k");
        Assert.Equal(value.Type, stored.Type);
        Assert.Equal(value, stored);
    }

    [Fact]
    public void Dictionary_StoresArraysWithTheirBounds()
    {
        var dictionary = Dictionary();
        var longs = VbaArray.Create(VarType.Long, [(1, 3)]);
        longs.Set([1], 10);
        longs.Set([2], 20);
        longs.Set([3], 30);
        var grid = VbaArray.Create(VarType.Variant, [(0, 1), (5, 6)]);
        grid.Set([0, 5], "a");
        grid.Set([1, 5], 2);
        grid.Set([0, 6], 3.5);
        grid.Set([1, 6], Variant.True);

        dictionary.Put(dictionary.GetDispId("Item"), ["longs"], Variant.FromArray(longs), asReference: false);
        dictionary.Put(dictionary.GetDispId("Item"), ["grid"], Variant.FromArray(grid), asReference: false);

        var storedLongs = Get(dictionary, "Item", "longs").AsArray();
        Assert.Equal(VarType.Long, storedLongs.ElementType);
        Assert.Equal((1, 3), (storedLongs.LBound(), storedLongs.UBound()));
        Assert.Equal(Variant.FromInt32(20), storedLongs.Get([2]));

        var storedGrid = Get(dictionary, "Item", "grid").AsArray();
        Assert.Equal(2, storedGrid.Rank);
        Assert.Equal((0, 1), (storedGrid.LBound(1), storedGrid.UBound(1)));
        Assert.Equal((5, 6), (storedGrid.LBound(2), storedGrid.UBound(2)));
        Assert.Equal(Variant.FromString("a"), storedGrid.Get([0, 5]));
        Assert.Equal(Variant.FromInt32(2), storedGrid.Get([1, 5]));
        Assert.Equal(Variant.FromDouble(3.5), storedGrid.Get([0, 6]));
        Assert.Equal(Variant.True, storedGrid.Get([1, 6]));
    }

    [Fact]
    public void Dictionary_StoresObjectsByReference()
    {
        var dictionary = Dictionary();
        var other = Dictionary();

        dictionary.Put(dictionary.GetDispId("Item"), ["self"], Variant.FromObject(dictionary), asReference: true);
        dictionary.Put(dictionary.GetDispId("Item"), ["other"], Variant.FromObject(other), asReference: true);
        dictionary.Put(dictionary.GetDispId("Item"), ["nothing"], Variant.Nothing, asReference: true);

        var self = Get(dictionary, "Item", "self");
        Assert.True(self.IsObject);
        Assert.NotSame(dictionary, self.AsObject());
        Assert.True(dictionary.IsSameObject(self.AsObject()));
        Assert.False(dictionary.IsSameObject(Get(dictionary, "Item", "other").AsObject()));
        Assert.False(dictionary.IsSameObject(null));
        Assert.True(Get(dictionary, "Item", "nothing").IsNothing);
        Assert.Equal(Variant.True, Operators.Is(Variant.FromObject(dictionary), self));
    }

    [Fact]
    public void FileSystemObject_ReturnsNestedObjectsDatesAndStrings()
    {
        var fso = (IDispatchObject)Provider.CreateObject("Scripting.FileSystemObject", null);

        Assert.Equal("FileSystemObject", fso.TypeName);
        var temp = Get(fso, "GetSpecialFolder", 2);
        Assert.True(temp.IsObject);
        var folder = (IDispatchObject)temp.AsObject()!;
        Assert.Equal("Folder", folder.TypeName);
        Assert.Equal(VarType.String, Get(folder, "Path").Type);
        Assert.Equal(VarType.Date, Get(folder, "DateLastModified").Type);
        Assert.Equal(VarType.String, Get(fso, "GetTempName").Type);
        Assert.Equal(Variant.True, Get(fso, "FolderExists", Get(folder, "Path")));
    }

    [Fact]
    public void LateBound_ReachesComObjectsThroughTheRuntime()
    {
        var dictionary = Variant.FromObject(Dictionary());

        LateBound.Call(dictionary, "Add", ["x", 42]);
        LateBound.Let(dictionary, "Item", ["y"], 43);

        Assert.Equal(Variant.FromInt32(42), LateBound.Get(dictionary, "Item", ["x"]));
        Assert.Equal(Variant.FromInt32(43), LateBound.Index(dictionary, ["y"]));
        Assert.Equal(Variant.FromInt32(2), LateBound.CallByName(dictionary, "Count", 2, []));
        LateBound.CallByName(dictionary, "Item", 4, ["x", 44]);
        Assert.Equal(Variant.FromInt32(44), LateBound.Get(dictionary, "Item", ["x"]));

        var keys = new List<Variant>();
        var enumerator = ForEachEnumerator.Create(dictionary);
        while (enumerator.MoveNext())
        {
            keys.Add(ObjectRefs.Own(enumerator.Current));
        }

        Assert.Equal([Variant.FromString("x"), Variant.FromString("y")], keys);
        Assert.Equal(438, Assert.Throws<VbaException>(() => LateBound.Get(dictionary, "Nope", [])).Number);
    }
}
