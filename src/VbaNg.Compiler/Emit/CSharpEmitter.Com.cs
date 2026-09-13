using System.Globalization;

using VbaNg.Compiler.Binding;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Emit;

/// <summary>Early-bound COM members (ARCHITECTURE.md section 6): calls by dispid through the runtime's EarlyBound helpers, application objects cached per module.</summary>
public sealed partial class CSharpEmitter
{
    private const string DispatchType = R + "IDispatchObject";

    private readonly SortedDictionary<string, (string LibraryId, string ClassId)> appObjects = new(StringComparer.Ordinal);

    /// <summary>True while a Set assignment is being stored, so a COM property chooses its Set accessor.</summary>
    private bool storeIsSet;

    /// <summary>True while a Let assignment is being stored, so an object value is let-coerced to its default member.</summary>
    private bool storeIsLet;

    /// <summary>The Variant a store receives: under Let an object stands for its default member's value (MS-VBAL 5.6.9.3); every other store keeps the reference.</summary>
    private string StoredVariant(Emitted value) =>
        storeIsLet && (value.Type.IsVariant || value.Type.IsObject)
            ? $"{R}Coerce.LetValue({Convert(value, VbaType.Variant)})"
            : LetVariant(value);

    private Emitted ComCall(BoundComCall call)
    {
        // Excel's Run of a procedure of this project runs in-process (ARCHITECTURE.md D23).
        var code = IsExcelRun(call)
            ? $"{R}Hosting.ApplicationRun.Invoke(typeof(global::{(Symbol.Kind == ModuleKind.Class ? Symbol.ClassEmitName : Symbol.EmitName)}), {ComTarget(call.Target)}, {Dispid(call.Member.DispId)}, {Arguments(call.Arguments)})"
            : $"{R}EarlyBound.Invoke({ComTarget(call.Target)}, {Dispid(call.Member.DispId)}, {InvokeKindCode(call.Kind)}, {Arguments(call.Arguments)})";
        return call.Type.IsVariant
            ? new Emitted(code, VbaType.Variant)
            : new Emitted(Convert(new Emitted(code, VbaType.Variant), call.Type), call.Type);
    }

    /// <summary>Excel's Run, on Application or unqualified through Global (dispid 259 on both).</summary>
    private static bool IsExcelRun(BoundComCall call) =>
        (call.Kind & Runtime.InvokeKind.Method) != 0
        && call.Member.DispId == 259
        && string.Equals(call.Member.Name, "Run", StringComparison.Ordinal)
        && string.Equals(call.Target.Type.Library?.Name, "Excel", StringComparison.Ordinal)
        && call.Target.Type.ComInterface?.Name is "_Application" or "_Global";

    /// <summary>The application object of a library, created on first use and kept in a field of the module class.</summary>
    private Emitted AppObject(BoundAppObject app)
    {
        var field = "__app_" + CSharpNames.Identifier(app.Library.Name).TrimStart('@');
        appObjects[field] = (app.Library.Guid.ToString("D", CultureInfo.InvariantCulture), app.AppObject.Guid.ToString("D", CultureInfo.InvariantCulture));
        return new Emitted($"({field} ??= {R}Com.AppObject({Quote(appObjects[field].LibraryId)}, {Quote(appObjects[field].ClassId)}))", app.Type);
    }

    /// <summary>Stores through a property's put accessor, or through the default member of the object a call returns.</summary>
    private string StoreCom(BoundComCall call, Emitted value, bool isSet)
    {
        var accessor = isSet ? call.PutRefMember ?? call.PutMember : call.PutMember ?? call.PutRefMember;
        if (accessor is not null)
        {
            var converted = call.PutValueType is { } declared && declared.IsScalar && !declared.IsVariant && !value.Type.IsObject && !value.Type.IsVariant
                ? Convert(new Emitted(Convert(value, declared), declared), VbaType.Variant)
                : StoredVariant(value);
            var asReference = accessor.Kind == Runtime.TypeLibraries.ComMemberKind.PropertyPutRef ? "true" : "false";
            return $"{R}EarlyBound.Put({ComTarget(call.Target)}, {Dispid(accessor.DispId)}, {Arguments(call.Arguments)}, {converted}, {asReference})";
        }

        if (call.DefaultPut is not null)
        {
            var read = ComCall(call);
            return $"{R}EarlyBound.Put({read.Code}, {Dispid(call.DefaultPut.DispId)}, [], {StoredVariant(value)}, false)";
        }

        Report(call.Syntax, "This property cannot be assigned to.");
        return "_ = " + value.Code;
    }

    private string ComTarget(BoundExpression target) =>
        target.Type.IsCom
            ? EmitExpression(target).Code
            : $"{R}Coerce.ToObject<{DispatchType}>({Convert(EmitExpression(target), VbaType.Variant)})";

    private static string InvokeKindCode(InvokeKind kind) => kind switch
    {
        InvokeKind.Method => R + "InvokeKind.Method",
        InvokeKind.PropertyPut => R + "InvokeKind.PropertyPut",
        InvokeKind.PropertyPutRef => R + "InvokeKind.PropertyPutRef",
        _ => R + "InvokeKind.PropertyGet",
    };

    private static string Dispid(int dispId) => dispId.ToString(CultureInfo.InvariantCulture);

    private static string ComClassId(VbaType type) =>
        Quote(type.ComType!.Guid.ToString("D", CultureInfo.InvariantCulture));

    private void EmitAppObjectFields()
    {
        foreach (var (field, _) in appObjects)
        {
            writer.HiddenLine($"private static {DispatchType}? {field};");
        }
    }
}
