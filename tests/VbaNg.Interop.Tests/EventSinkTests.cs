using VbaNg.Runtime;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The native IDispatch sink: what a COM source sees when it calls it. The tests play the source
/// themselves, so no object with events is needed; advising on a real source is covered by the
/// Excel-driving E2E tests.
/// </summary>
public sealed unsafe class EventSinkTests
{
    private static readonly Guid SourceIid = new("11111111-2222-3333-4444-555555555555");

    [Fact]
    public void Sink_AnswersTheSourceInterfaceAndCountsReferences()
    {
        using var sink = new EventSink(SourceIid, (_, _) => { });

        var asSource = NativeMethods.QueryInterface(sink.Unknown, in SourceIid);
        var asDispatch = NativeMethods.QueryInterface(sink.Unknown, in NativeMethods.IidIDispatch);
        var asOther = NativeMethods.QueryInterface(sink.Unknown, in NativeMethods.IidIEnumVariant);

        Assert.Equal(sink.Unknown, asSource);
        Assert.Equal(sink.Unknown, asDispatch);
        Assert.Equal(0, asOther);
        NativeMethods.Release(asSource);
        NativeMethods.Release(asDispatch);
        Assert.False(sink.IsAdvised);
    }

    [Fact]
    public void Sink_InvokeForwardsArgumentsAndWritesBackByRef()
    {
        var received = new List<(int DispId, Variant[] Arguments)>();
        using var sink = new EventSink(SourceIid, (dispId, arguments) =>
        {
            received.Add((dispId, (Variant[])arguments.Clone()));
            arguments[1] = Variant.True;
        });

        // Two arguments as a source passes them: a string ByVal and a Boolean ByRef (Cancel), right to left in rgvarg.
        short cancel = 0;
        var natives = stackalloc VARIANT[2];
        natives[1].vt = Vt.Bstr;
        natives[1].pointer = VariantMarshal.AllocBstr("Sheet1");
        natives[0].vt = (ushort)(Vt.ByRef | Vt.Bool);
        natives[0].pointer = (nint)(&cancel);
        var parameters = new DISPPARAMS { rgvarg = (nint)natives, cArgs = 2 };
        var iid = Guid.Empty;
        VARIANT result;
        EXCEPINFO exception = default;
        uint argumentError;

        var hr = ((delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, DISPPARAMS*, VARIANT*, EXCEPINFO*, uint*, int>)NativeMethods.Slot(sink.Unknown, 6))(
            sink.Unknown, 1548, &iid, 0, (ushort)InvokeKind.Method, &parameters, &result, &exception, &argumentError);
        NativeMethods.SysFreeString(natives[1].pointer);

        Assert.Equal(0, hr);
        var call = Assert.Single(received);
        Assert.Equal(1548, call.DispId);
        Assert.Equal([Variant.FromString("Sheet1"), Variant.False], call.Arguments);
        Assert.Equal(-1, cancel);
    }

    [Fact]
    public void Sink_ReportsHandlerErrorsWithoutFailingTheSource()
    {
        Exception? reported = null;
        using var sink = new EventSink(SourceIid, (_, _) => throw new VbaException(11))
        {
            Error = ex => reported = ex,
        };
        var iid = Guid.Empty;
        var parameters = new DISPPARAMS();
        VARIANT result;
        EXCEPINFO exception = default;
        uint argumentError;

        var hr = ((delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort, DISPPARAMS*, VARIANT*, EXCEPINFO*, uint*, int>)NativeMethods.Slot(sink.Unknown, 6))(
            sink.Unknown, 1, &iid, 0, (ushort)InvokeKind.Method, &parameters, &result, &exception, &argumentError);

        Assert.Equal(0, hr);
        Assert.Equal(11, Assert.IsType<VbaException>(reported).Number);
    }

    [Fact]
    public void Advise_OnAnObjectWithoutConnectionPoints_Raises459()
    {
        var provider = new ComProvider();
        var dictionary = (ComObject)provider.CreateObject("Scripting.Dictionary", null);
        using var sink = new EventSink(SourceIid, (_, _) => { });

        var error = Assert.Throws<VbaException>(() => dictionary.Advise(sink));

        Assert.Equal(459, error.Number);
        Assert.Null(EventSource.Describe(dictionary));
    }
}
