using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The COM identity of the runtime's own objects (ARCHITECTURE.md section 5, "Class modules";
/// ROADMAP.md M7 D2): one reference count for VBA code and COM callers, the object itself handed
/// to COM and recognized when it comes back, IDispatch served through the object's dispid table,
/// and a project reset that destroys an instance a COM caller still holds and cuts its block off.
/// Exercised through the vtable directly and through Scripting.Dictionary, which needs no Excel
/// (R9). The tests of a class run one at a time, and only this class makes Probe instances, so
/// the reset in the last test destroys nothing else.
/// </summary>
public sealed unsafe class RuntimeObjectTests
{
    private const int ENoInterface = unchecked((int)0x80004002);
    private const int RpcEDisconnected = unchecked((int)0x80010108);

    private static readonly Guid IidIUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly ComProvider Provider = new();

    /// <summary>IDispatch::Invoke on a runtime object and reading COM objects back out of Variants go through the provider, installed as the hosts install it.</summary>
    public RuntimeObjectTests() => Com.Provider = Provider;

    [Fact]
    public void Pointer_SharesTheCountWithVbaCode()
    {
        var probe = new Probe();
        var pointer = probe.AddRefPointer();
        Assert.Equal(1, probe.References);
        Assert.Same(probe, RuntimeObject.FromPointer(pointer));

        var iid = IidIUnknown;
        nint unknown;
        Assert.Equal(0, QueryInterface(pointer, &iid, &unknown));
        Assert.Equal(pointer, unknown);
        Assert.Equal(2, probe.References);
        var other = Guid.NewGuid();
        Assert.Equal(ENoInterface, QueryInterface(pointer, &other, &unknown));

        // A VBA reference and two COM references on one count: the object terminates when the last goes, whoever drops it.
        probe.AddRef();
        Release(pointer);
        Release(pointer);
        Assert.Equal(0, probe.Terminated);
        probe.Release();
        Assert.Equal(1, probe.Terminated);
        Assert.True(probe.FieldsReleased);
        Assert.Equal(0, probe.References);
    }

    /// <summary>
    /// For Each over a COM collection: the element the loop is on keeps a reference of the
    /// enumerator's own until the next one, since the header statement's temporaries end before
    /// the loop variable takes its reference (ARCHITECTURE.md D18). Across processes a proxy no
    /// one holds is destroyed at once; Excel crashed on the next use of a multi-area range's area
    /// (VBA-Better-Array's FromExcelRange).
    /// </summary>
    [Fact]
    public void Enumerator_HoldsTheElementItIsOn_AfterTheStatementEnds()
    {
        var mark = ObjectRefs.Mark();
        IDispatchObject? dictionary = null;
        ObjectRefs.Assign(ref dictionary, (IDispatchObject)Provider.CreateObject("Scripting.Dictionary", null));
        var probe = new Probe();
        probe.AddRef();
        dictionary!.Invoke(dictionary.GetDispId("Add"), InvokeKind.Method, [Variant.FromObject(probe), Variant.FromInt32(1)]);
        ObjectRefs.ReleaseTo(mark);
        var held = probe.References;

        var enumerator = dictionary.Enumerate();
        Assert.True(enumerator.MoveNext());
        ObjectRefs.ReleaseTo(mark);
        var during = probe.References;
        enumerator.Dispose();
        var after = probe.References;

        dictionary.Invoke(dictionary.GetDispId("RemoveAll"), InvokeKind.Method, []);
        ObjectRefs.Release(ref dictionary);
        probe.Release();
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(held + 1, during);
        Assert.Equal(held, after);
    }

    [Fact]
    public void Dictionary_KeepsTheObjectItself_AndReleasesItLast()
    {
        var mark = ObjectRefs.Mark();
        IDispatchObject? dictionary = null;
        ObjectRefs.Assign(ref dictionary, (IDispatchObject)Provider.CreateObject("Scripting.Dictionary", null));
        var probe = new Probe();
        var slot = Variant.Empty;
        ObjectRefs.Assign(ref slot, Variant.FromObject(ObjectRefs.Owned(probe)));
        var counts = new List<int> { probe.References };

        Call(dictionary!, "Add", Variant.FromString("k"), slot);
        counts.Add(probe.References);
        var back = Call(dictionary!, "Item", Variant.FromString("k"));
        counts.Add(probe.References);
        Assert.Same(probe, back.AsObject());
        ObjectRefs.ReleaseTo(mark);
        counts.Add(probe.References);

        // VBA's reference goes first; the dictionary holds the rest and drops them from COM, the last one terminating the object once.
        ObjectRefs.Release(ref slot);
        counts.Add(probe.References);
        Assert.True(probe.Terminated == 0, "counts: " + string.Join(",", counts));
        Call(dictionary!, "RemoveAll");
        counts.Add(probe.References);
        Assert.True(probe.Terminated == 1 && probe.References == 0, "counts: " + string.Join(",", counts));
        ObjectRefs.Release(ref dictionary);
    }

    [Fact]
    public void Invoke_FromCom_RunsTheMemberThroughItsDispidTable()
    {
        var probe = new Probe();
        var pointer = probe.AddRefPointer();
        try
        {
            Assert.Equal(1, DispIdOf(pointer, "Echo"));
            Assert.Equal(NativeMethods.DispEUnknownName, GetIDsOfNames(pointer, "Nope", out var missing));
            Assert.Equal(-1, missing);

            var argument = new VARIANT { vt = Vt.Bstr, pointer = VariantMarshal.AllocBstr("hi") };
            var parameters = new DISPPARAMS { rgvarg = (nint)(&argument), cArgs = 1 };
            VARIANT result = default;
            EXCEPINFO info = default;
            Assert.Equal(0, Invoke(pointer, 1, 1, &parameters, &result, &info));
            Assert.Equal(Vt.Bstr, result.vt);
            Assert.Equal("hi", VariantMarshal.ReadBstr(result.pointer));
            VariantMarshal.Clear(&result);
            VariantMarshal.Clear(&argument);

            // A run-time error reaches the caller as DISP_E_EXCEPTION, described the way VBA describes it.
            var none = default(DISPPARAMS);
            Assert.Equal(NativeMethods.DispEException, Invoke(pointer, 2, 1, &none, &result, &info));
            Assert.Equal(unchecked((int)0x800A0005), info.scode);
            Assert.Equal("Invalid procedure call or argument", VariantMarshal.ReadBstr(info.bstrDescription));
            NativeMethods.SysFreeString(info.bstrDescription);
            NativeMethods.SysFreeString(info.bstrSource);
            NativeMethods.SysFreeString(info.bstrHelpFile);
        }
        finally
        {
            Release(pointer);
        }

        Assert.Equal(1, probe.Terminated);
    }

    [Fact]
    public void Reset_DestroysAnInstanceComStillHolds_WithoutClassTerminate()
    {
        var probe = new Probe();
        var pointer = probe.AddRefPointer();

        ProjectReset.Run(typeof(Probe).Assembly);

        Assert.Equal(0, probe.Terminated);
        Assert.True(probe.FieldsReleased);
        Assert.Null(RuntimeObject.FromPointer(pointer));
        var iid = IidIUnknown;
        nint unknown;
        Assert.Equal(RpcEDisconnected, QueryInterface(pointer, &iid, &unknown));
        Assert.Equal(0u, Release(pointer));
    }

    [Fact]
    public void ErrAndTheEnumerator_CrossToComAsThemselves()
    {
        var mark = ObjectRefs.Mark();
        var enumerator = new VbaNg.Runtime.Library.Collection().NewEnum();
        foreach (var value in new object[] { Err.Current, enumerator })
        {
            VARIANT native;
            VariantMarshal.ToNative(Variant.FromObject(value), &native);
            Assert.Equal(Vt.Dispatch, native.vt);
            Assert.Same(value, VariantMarshal.FromNative(&native).AsObject());
            VariantMarshal.Clear(&native);
        }

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(0, enumerator.References);
    }

    /// <summary>A Collection and Err handed to COM answer a caller's member calls through their late-bound members (ROADMAP.md M7 F): the default member at dispid 0, the rest by name.</summary>
    [Fact]
    public void CollectionAndErr_AnswerComCallers_ThroughTheirMembers()
    {
        var collection = new VbaNg.Runtime.Library.Collection();
        var pointer = collection.AddRefPointer();
        try
        {
            Assert.Equal(1, DispIdOf(pointer, "add"));
            var argument = new VARIANT { vt = Vt.Bstr, pointer = VariantMarshal.AllocBstr("x") };
            var parameters = new DISPPARAMS { rgvarg = (nint)(&argument), cArgs = 1 };
            VARIANT result = default;
            EXCEPINFO info = default;
            Assert.Equal(0, Invoke(pointer, 1, 1, &parameters, &result, &info));
            VariantMarshal.Clear(&argument);

            var none = default(DISPPARAMS);
            Assert.Equal(0, Invoke(pointer, DispIdOf(pointer, "Count"), 2, &none, &result, &info));
            Assert.Equal((Vt.I4, 1L), (result.vt, result.data));

            var index = new VARIANT { vt = Vt.I4, data = 1 };
            var byIndex = new DISPPARAMS { rgvarg = (nint)(&index), cArgs = 1 };
            Assert.Equal(0, Invoke(pointer, 0, 2, &byIndex, &result, &info));
            Assert.Equal("x", VariantMarshal.ReadBstr(result.pointer));
            VariantMarshal.Clear(&result);
        }
        finally
        {
            Release(pointer);
        }

        var err = Err.Current;
        err.Number = 11;
        var errPointer = err.AddRefPointer();
        try
        {
            var none = default(DISPPARAMS);
            VARIANT result = default;
            EXCEPINFO info = default;
            Assert.Equal(0, Invoke(errPointer, 0, 2, &none, &result, &info));
            Assert.Equal((Vt.I4, 11L), (result.vt, result.data));
        }
        finally
        {
            Release(errPointer);
            err.Clear();
        }
    }

    /// <summary>A COM caller asks a Collection for DISPID_NEWENUM and walks what it gets through IEnumVARIANT, as For Each over a COM object does (ROADMAP.md M7; the Declares golden holds each method to VBA's).</summary>
    [Fact]
    public void Collection_WalksForAComCaller_ThroughIEnumVariant()
    {
        var collection = new VbaNg.Runtime.Library.Collection();
        collection.Add(Variant.FromInt32(10));
        collection.Add(Variant.FromInt32(20));
        collection.Add(Variant.FromInt32(30));
        var pointer = collection.AddRefPointer();
        var com = ComObject.FromPointer(pointer);
        com.AddRef();
        var walked = new List<int>();
        try
        {
            using var items = com.Enumerate();
            while (items.MoveNext())
            {
                walked.Add(Coerce.ToInt32(items.Current));
            }

            items.Reset();
            Assert.True(items.MoveNext());
            walked.Add(Coerce.ToInt32(items.Current));
        }
        finally
        {
            com.Release();
            Release(pointer);
        }

        Assert.Equal([10, 20, 30, 10], walked);
        Assert.Equal(0, collection.References);
    }

    private static Variant Call(IDispatchObject target, string name, params Variant[] arguments) =>
        target.Invoke(target.GetDispId(name), InvokeKind.Method | InvokeKind.PropertyGet, arguments);

    private static int QueryInterface(nint self, Guid* iid, nint* result) =>
        ((delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int>)NativeMethods.Slot(self, 0))(self, iid, result);

    private static uint Release(nint self) => ((delegate* unmanaged[Stdcall]<nint, uint>)NativeMethods.Slot(self, 2))(self);

    private static int DispIdOf(nint self, string name)
    {
        Assert.Equal(0, GetIDsOfNames(self, name, out var id));
        return id;
    }

    private static int GetIDsOfNames(nint self, string name, out int id)
    {
        int result;
        fixed (char* text = name)
        {
            var names = stackalloc char*[1];
            names[0] = text;
            var iid = Guid.Empty;
            int dispId;
            result = ((delegate* unmanaged[Stdcall]<nint, Guid*, char**, uint, uint, int*, int>)NativeMethods.Slot(self, 5))(self, &iid, names, 1, 0, &dispId);
            id = dispId;
        }

        return result;
    }

    private static int Invoke(nint self, int dispId, ushort flags, DISPPARAMS* parameters, VARIANT* result, EXCEPINFO* info)
    {
        var iid = Guid.Empty;
        uint argumentError;
        return ((delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, DISPPARAMS*, VARIANT*, EXCEPINFO*, uint*, int>)NativeMethods.Slot(self, 6))(self, dispId, &iid, 0, flags, parameters, result, info, &argumentError);
    }

    /// <summary>A class instance as generated code would make one: a dispid table, Class_Terminate, and storage to release.</summary>
    private sealed class Probe : VbaClassObject
    {
        public int Terminated { get; private set; }

        public bool FieldsReleased { get; private set; }

        public override string TypeName => "Probe";

        public override int GetDispId(string name) => name.ToUpperInvariant() switch
        {
            "VALUE" => 0,
            "ECHO" => 1,
            "FAIL" => 2,
            _ => NoMember(),
        };

        public override Variant Invoke(int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments) => dispId switch
        {
            0 => Variant.FromInt32(42),
            1 => arguments[0],
            2 => throw VbaErrors.InvalidProcedureCall(),
            _ => NotSupported(),
        };

        public override void Put(int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference) => NotSupported();

        protected override void ClassTerminate() => Terminated++;

        protected override void ReleaseFields() => FieldsReleased = true;
    }
}
