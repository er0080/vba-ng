using VbaNg.Runtime;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The lifetime of a COM object VBA code reaches through a call (ARCHITECTURE.md D18): the
/// wrapper a call returns is a temporary of the current statement, released when the statement
/// ends unless a store took a reference of its own, so Excel's objects are released on the
/// thread that made them, never by the finalizer. Exercised through scrrun.dll, which needs no
/// Excel (R9).
/// </summary>
public sealed class ComObjectLifetimeTests
{
    private static readonly ComProvider Provider = new();

    /// <summary>A Variant holds a COM object as its pointer; reading it back as an object goes through the provider, installed as the hosts install it.</summary>
    public ComObjectLifetimeTests() => Com.Provider = Provider;

    private static Variant Get(IDispatchObject target, string name, params Variant[] arguments) =>
        target.Invoke(target.GetDispId(name), InvokeKind.PropertyGet | InvokeKind.Method, arguments);

    [Fact]
    public void CallResult_IsATemporaryOfTheStatement()
    {
        var fso = (IDispatchObject)Provider.CreateObject("Scripting.FileSystemObject", null);
        var path = Variant.FromString(Path.GetTempPath());
        var mark = ObjectRefs.Mark();

        // The result is the object's IDispatch pointer with a reference the statement owns (D20): one temporary, no wrapper kept.
        var folder = Get(fso, "GetFolder", path);
        Assert.NotEqual(0, folder.InterfacePointer);
        Assert.Equal(mark + 1, ObjectRefs.PendingTemporaries);
        Assert.False(string.IsNullOrEmpty(Coerce.ToString(Get((IDispatchObject)folder.AsObject()!, "Name"))));

        // The statement ends and its reference goes with it; the pointer is not touched again (RuntimeObjectTests watch a count go down).
        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(mark, ObjectRefs.PendingTemporaries);
    }

    [Fact]
    public void StoredCallResult_OutlivesTheStatement()
    {
        var fso = (IDispatchObject)Provider.CreateObject("Scripting.FileSystemObject", null);
        var mark = ObjectRefs.Mark();

        IDispatchObject? slot = null;
        ObjectRefs.Assign(ref slot, (IDispatchObject?)Get(fso, "GetFolder", Variant.FromString(Path.GetTempPath())).AsObject());
        ObjectRefs.ReleaseTo(mark);

        Assert.False(string.IsNullOrEmpty(Coerce.ToString(Get(slot!, "Name"))));
        var kept = slot!;
        ObjectRefs.Release(ref slot);
        Assert.Null(slot);
        Assert.Equal(VbaErrors.ObjectVariableNotSetNumber, Assert.Throws<VbaException>(() => Get(kept, "Name")).Number);
    }

    [Fact]
    public void CreatedObject_IsATemporaryOfTheStatement()
    {
        var mark = ObjectRefs.Mark();
        var dictionary = (IDispatchObject)Provider.CreateObject("Scripting.Dictionary", null);
        Assert.Equal(mark + 1, ObjectRefs.PendingTemporaries);

        ObjectRefs.ReleaseTo(mark);
        Assert.Equal(VbaErrors.ObjectVariableNotSetNumber, Assert.Throws<VbaException>(() => Get(dictionary, "Count")).Number);
    }
}
