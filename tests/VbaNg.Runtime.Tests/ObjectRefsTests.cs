using System.Runtime.CompilerServices;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// Set on a WithEvents variable of a library type (MS-VBAL 5.2.3.1.4): the runtime advises the
/// new object through <see cref="IComEventSource"/>, which the Interop wrapper implements, and
/// drops the advisory on the old one; an object that raises no COM events is assigned plainly.
/// </summary>
[Collection(LiveBstrTests.Name)]
public sealed class ObjectRefsTests
{
    private static readonly ComEventInterface Events = new(Guid.NewGuid(), "DocEvents", [(1545, "Change")]);

    [Fact]
    public void Subscribe_AdvisesTheNewObjectAndUnadvisesTheOld()
    {
        var sink = new Sink();
        var first = new Source();
        var second = new Source();
        Source? slot = null;
        IDisposable? advisory = null;

        ObjectRefs.Subscribe(ref slot, first, sink, "ws", Events, ref advisory);
        Assert.Same(first, slot);
        Assert.Equal([("ws", Events)], first.Advised);
        Assert.False(first.Unadvised);

        ObjectRefs.Subscribe(ref slot, second, sink, "ws", Events, ref advisory);
        Assert.Same(second, slot);
        Assert.True(first.Unadvised);
        Assert.Single(second.Advised);

        ObjectRefs.Subscribe(ref slot, null, sink, "ws", Events, ref advisory);
        Assert.Null(slot);
        Assert.True(second.Unadvised);
        Assert.Null(advisory);
    }

    [Fact]
    public void Subscribe_SameObject_KeepsTheAdvisory()
    {
        var source = new Source();
        Source? slot = null;
        IDisposable? advisory = null;
        ObjectRefs.Subscribe(ref slot, source, new Sink(), "ws", Events, ref advisory);

        ObjectRefs.Subscribe(ref slot, source, new Sink(), "ws", Events, ref advisory);

        Assert.Single(source.Advised);
        Assert.False(source.Unadvised);
    }

    [Fact]
    public void Subscribe_ObjectWithoutComEvents_IsAssignedWithoutAnAdvisory()
    {
        object? slot = null;
        IDisposable? advisory = null;
        var value = new object();

        ObjectRefs.Subscribe(ref slot, value, new Sink(), "ws", Events, ref advisory);

        Assert.Same(value, slot);
        Assert.Null(advisory);
    }

    // The ownership rails for arrays and records (ARCHITECTURE.md D18, D20; ROADMAP.md M7 B5): a
    // store copies, a release frees the copy's elements or references, a result moves to the
    // caller's statement, a For Each keeps a temporary array for the loop, and a late-bound
    // ByRef cell owns a copy. Each test checks the live BSTR count (the Live BSTRs collection).

    [Fact]
    public void Assign_ArrayIntoVariant_CopiesItAndReleasesWhatTheSlotHeld()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var source = VbaArray.FromValues(VarType.String, [Variant.FromString("a"), Variant.FromString("b")]);
        var slot = Variant.Empty;
        ObjectRefs.Assign(ref slot, Variant.FromArray(source));
        Assert.NotEqual(source.Descriptor, slot.AsArray().Descriptor);

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 2, Bstr.LiveCount);
        Assert.Equal("b", slot.AsArray().Get([1]).AsVbaString().ToString());

        ObjectRefs.Assign(ref slot, Variant.FromInt32(1));
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void Transfer_ArrayResult_BecomesATemporaryOfTheCallersStatement()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var result = VbaArray.Create(VarType.String, [(1, 2)]);
        result.Set([1], Variant.FromString("x"));
        result.Set([2], Variant.FromString("y"));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 2, Bstr.LiveCount);

        var value = ObjectRefs.Transfer(ref result);
        Assert.False(result.IsAllocated);
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 2, Bstr.LiveCount);
        Assert.Equal("y", value.Get([2]).AsVbaString().ToString());

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void Slot_OwnsACopyTheStatementReleases()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var argument = ObjectRefs.Own(Variant.FromString("kept"));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 1, Bstr.LiveCount);

        ref var cell = ref ObjectRefs.Slot(argument);
        ObjectRefs.Assign(ref cell, Variant.FromString("changed"));
        Assert.Equal("kept", argument.AsVbaString().ToString());

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 1, Bstr.LiveCount);
        ObjectRefs.Release(ref argument);
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void ForEach_OverATemporaryArray_KeepsItForTheLoop()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var enumerator = ForEachEnumerator.Create(Variant.FromArray(VbaArray.FromValues(VarType.String, [Variant.FromString("p"), Variant.FromString("q")])));
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(before + 2, Bstr.LiveCount);

        Assert.True(enumerator.MoveNext());
        Assert.Equal("p", enumerator.Current.AsVbaString().ToString());
        Assert.True(enumerator.MoveNext());
        Assert.Equal("q", enumerator.Current.AsVbaString().ToString());
        Assert.False(enumerator.MoveNext());
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void ForEach_OverAVariablesArray_WalksItLive()
    {
        var before = Bstr.LiveCount;
        var mark = ObjectRefs.Mark();
        var array = VbaArray.Create(VarType.String, [(0, 1)]);
        array.Set([0], Variant.FromString("p"));
        var enumerator = ForEachEnumerator.Create(Variant.FromArray(array));
        array.Set([1], Variant.FromString("late"));
        ObjectRefs.ReleaseTo(mark);

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.MoveNext());
        Assert.Equal("late", enumerator.Current.AsVbaString().ToString());
        Assert.False(enumerator.MoveNext());
        Assert.Equal(before + 2, Bstr.LiveCount);

        ObjectRefs.Release(ref array);
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void AssignRecord_TakesACopy_AndReleasesWhatTheSlotHeld()
    {
        var before = Bstr.LiveCount;
        var original = new Named { Name = VbaString.Alloc("n") };
        var slot = Named.Fresh();
        ObjectRefs.AssignRecord(ref slot, original.Copy());
        Assert.Equal(before + 2, Bstr.LiveCount);
        Assert.Equal("n", slot.Name.ToString());

        ObjectRefs.AssignRecord(ref slot, Named.Fresh());
        Assert.Equal(before + 1, Bstr.LiveCount);
        Assert.Equal("n", original.Name.ToString());

        original.ReleaseReferences();
        Assert.Equal(before, Bstr.LiveCount);
    }

    [Fact]
    public void RecordArray_KeepsItsElementsTyped_CopiedAndReleasedWithIt()
    {
        var before = Bstr.LiveCount;
        var array = VbaArray.CreateRecords<Named>([(0, 1)]);
        ObjectRefs.AssignRecord(ref array.Record<Named>([1]), new Named { Name = VbaString.Alloc("x") });
        var copy = array.Clone();
        Assert.Equal(before + 2, Bstr.LiveCount);
        Assert.Equal("x", copy.Record<Named>([1]).Name.ToString());

        // ReDim Preserve drops the element it cuts off; a record never comes out as a Variant.
        VbaArray.ReDimRecords<Named>(ref array, [(0, 0)], preserve: true);
        Assert.Equal(before + 1, Bstr.LiveCount);
        var kept = array;
        Assert.Throws<VbaException>(() => kept.Get([0]));

        ObjectRefs.Release(ref copy);
        ObjectRefs.Release(ref array);
        Assert.Equal(before, Bstr.LiveCount);
    }

    private struct Named : IVbaRecord<Named>
    {
        public VbaString Name;

        public Named() => Name = VbaString.Null;

        public static Named Fresh() => new();

        public readonly Named Copy() => this with { Name = ObjectRefs.Own(Name) };

        public void ReleaseReferences() => ObjectRefs.Release(ref Name);
    }

    /// <summary>
    /// A typed object variable is its interface pointer (ROADMAP.md M7 E3): one machine word, a
    /// store takes a reference on the new object and drops the old one, and a pointer written
    /// into the slot raw, as CopyMemory writes it, reaches the object without a reference.
    /// </summary>
    [Fact]
    public void ObjectSlot_IsTheInterfacePointer_OwningOneReference()
    {
        var first = new Library.Collection();
        var second = new Library.Collection();
        var slot = default(ObjectSlot<Library.Collection>);
        Assert.Equal(8, Unsafe.SizeOf<ObjectSlot<Library.Collection>>());
        Assert.Null(slot.Target);

        ObjectRefs.Assign(ref slot, first);
        Assert.Equal(1, first.References);
        Assert.Same(first, slot.Target);
        Assert.Equal(Variant.FromObject(first).InterfacePointer, slot.Pointer);

        ObjectRefs.Assign(ref slot, second);
        Assert.Equal(0, first.References);
        Assert.Equal(1, second.References);

        var copy = default(ObjectSlot<Library.Collection>);
        Unsafe.As<ObjectSlot<Library.Collection>, nint>(ref copy) = slot.Pointer;
        Assert.Same(second, copy.Target);
        Assert.Equal(1, second.References);

        ObjectRefs.Release(ref slot);
        Assert.Equal(0, second.References);
        Assert.Equal(0, slot.Pointer);
    }

    /// <summary>The cell a late-bound call hands a ByRef object parameter holds a reference of its own until the statement ends.</summary>
    [Fact]
    public void ObjectSlotCell_HoldsAReference_TheStatementReleases()
    {
        var item = new Library.Collection();
        var mark = ObjectRefs.Mark();
        ref var cell = ref ObjectRefs.Cell<Library.Collection>(item);
        Assert.Same(item, cell.Target);
        Assert.Equal(1, item.References);

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(0, item.References);
    }

    private sealed class Sink : IVbaEventSink
    {
        public void RaiseVbaEvent(string source, string name, Variant[] arguments)
        {
        }
    }

    private sealed class Source : IComEventSource
    {
        public List<(string Variable, ComEventInterface Events)> Advised { get; } = [];

        public bool Unadvised { get; private set; }

        public IDisposable Advise(ComEventInterface events, IVbaEventSink sink, string variable)
        {
            Advised.Add((variable, events));
            return new Advisory(this);
        }

        private sealed class Advisory(Source owner) : IDisposable
        {
            public void Dispose() => owner.Unadvised = true;
        }
    }
}
