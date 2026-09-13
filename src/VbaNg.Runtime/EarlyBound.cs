namespace VbaNg.Runtime;

/// <summary>
/// Members bound by dispid at compile time (ARCHITECTURE.md section 6, "Binding modes"):
/// generated code for a variable declared with a COM type calls these with the dispid the type
/// library gave, so no name lookup happens at run time. Nothing raises 91 as VBA does.
/// </summary>
public static class EarlyBound
{
    public static Variant Invoke(IDispatchObject? target, int dispId, InvokeKind kind, ReadOnlySpan<Variant> arguments) =>
        Require(target).Invoke(dispId, kind, arguments);

    public static void Put(IDispatchObject? target, int dispId, ReadOnlySpan<Variant> indices, in Variant value, bool asReference) =>
        Require(target).Put(dispId, indices, value, asReference);

    /// <summary>The value an object stands for in a Let context: its default member with no arguments (MS-VBAL 5.6.9.3 let-coercion of an object).</summary>
    public static Variant DefaultValue(IDispatchObject? target) => Coerce.LetValue(Variant.FromObject(Require(target)));

    public static IDispatchObject Require(IDispatchObject? target) => target ?? throw VbaErrors.ObjectVariableNotSet();
}
