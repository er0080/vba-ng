using System.Globalization;

using VbaNg.Compiler.Binding;

namespace VbaNg.Compiler.Emit;

/// <summary>
/// Class modules (MS-VBAL 4.2) and the reference counting that gives their instances VBA's
/// lifetime (ARCHITECTURE.md D18). A class becomes a sealed C# class deriving from
/// <c>VbaClassObject</c>, with its module variables as instance fields, its procedures as
/// instance methods, <c>Class_Initialize</c> in the constructor, <c>Class_Terminate</c> and the
/// release of its fields in <c>Terminate</c>, and a dispid table so late binding, CallByName,
/// default members, and For Each reach it the way they reach a COM object. A class another
/// class implements is emitted as a C# interface plus its own implementation.
/// </summary>
public sealed partial class CSharpEmitter
{
    /// <summary>The members of the class, in declaration order, with the dispid each answers to.</summary>
    private List<(Symbol Member, int DispId)> DispatchMembers()
    {
        var members = new List<(Symbol, int)>();
        var next = 1;
        foreach (var member in PublicMembers())
        {
            var dispId = UserMemberId(member) ?? next++;
            members.Add((member, dispId));
        }

        return members;
    }

    /// <summary>The public members late binding can reach: the fields as they are declared, then the procedures and properties.</summary>
    private IEnumerable<Symbol> PublicMembers()
    {
        foreach (var variable in Symbol.Variables.Where(v => v.IsPublic))
        {
            yield return variable;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var procedure in Symbol.Procedures.Where(p => p.IsPublic && p.ImplementsInterface is null && p.EventSource is null))
        {
            var member = Symbol.Members.GetValueOrDefault(procedure.Name);
            if (member is not null && seen.Add(procedure.Name))
            {
                yield return member;
            }
        }
    }

    private static int? UserMemberId(Symbol member) => member switch
    {
        ProcedureSymbol procedure => procedure.UserMemberId,
        PropertySymbol property => property.Get?.UserMemberId ?? property.Let?.UserMemberId ?? property.Set?.UserMemberId,
        _ => null,
    };

    /// <summary>The class module's own declarations: the predeclared instance, the constructor, the terminator, and the dispid table.</summary>
    private void EmitClassMembers()
    {
        writer.HiddenLine($"public override string TypeName => {Quote(Symbol.Name)};");
        writer.Line();
        if (Symbol.Instance is { } instance)
        {
            writer.HiddenLine($"public static {instance.Type.CSharpName} {instance.EmitName};");
            writer.Line();
        }

        var initialize = Symbol.Procedures.FirstOrDefault(p => p.Name.Equals("Class_Initialize", StringComparison.OrdinalIgnoreCase));
        writer.HiddenLine($"public {Symbol.ClassEmitName}(){(HasInstanceBlock ? " : base(new __Fields())" : string.Empty)}");
        writer.Open();
        if (initialize is not null)
        {
            // New runs Class_Initialize; an error in it leaves no object behind (Classes golden). The object counts as
            // referenced meanwhile, as New holds it in VBA, so a reference Class_Initialize takes and drops does not end it.
            writer.HiddenLine("base.BeginConstruction();");
            writer.HiddenLine("try");
            writer.Open();
            writer.HiddenLine($"{initialize.EmitName}();");
            writer.Close();
            writer.HiddenLine("finally");
            writer.Open();
            writer.HiddenLine("base.EndConstruction();");
            writer.Close();
        }

        writer.Close();
        writer.Line();

        // The last reference going runs Class_Terminate, then releases the instance's storage; a project being
        // reset skips Class_Terminate (VbaClassObject.Release; docs/vba-quirks.md).
        var terminate = Symbol.Procedures.FirstOrDefault(p => p.Name.Equals("Class_Terminate", StringComparison.OrdinalIgnoreCase));
        if (terminate is not null)
        {
            writer.HiddenLine($"protected override void ClassTerminate() => {terminate.EmitName}();");
            writer.Line();
        }

        // An instance's Static locals are its storage too (MS-VBAL 5.2.3.1), released with its variables.
        var released = Symbol.Variables.Where(NeedsRelease).Concat(Symbol.StaticLocals.Where(s => OwnsStorage(s.Type))).ToList();
        if (released.Count > 0)
        {
            writer.HiddenLine("protected override void ReleaseFields()");
            writer.Open();
            EmitReleases(released);
            writer.Close();
            writer.Line();
        }

        EmitDispatch();
        EmitEventSink();
        EmitInterfaceForwarders();
    }

    private void EmitReleases(IReadOnlyList<VariableSymbol> variables)
    {
        foreach (var variable in variables)
        {
            writer.HiddenLine(ReleaseCode(variable, variable.EmitName));
        }
    }

    /// <summary>
    /// A class module's variable that lies in the instance's block (ARCHITECTURE.md section 5;
    /// ROADMAP.md M7 E4), so its address is stable: all but a WithEvents variable, which keeps its
    /// subscription beside the managed object, and a fixed-length string, still a managed string.
    /// </summary>
    private bool InInstanceBlock(VariableSymbol variable) => Symbol.Kind == ModuleKind.Class && !variable.IsWithEvents && variable.Type.Kind != TypeKind.FixedString;

    private bool HasInstanceBlock => Symbol.Variables.Concat(Symbol.StaticLocals).Any(InInstanceBlock);

    /// <summary>
    /// The C# access of a variable's storage: public when <paramref name="exposed"/>, unless it holds
    /// a Private Type, which C# cannot expose more widely than the type. VBA allows a Public
    /// variable of a Private Type in a standard module (VBA-JSON's JsonOptions), and a class keeps
    /// its variables of one in its instance block (VBA-Dictionary's PreciseTimer); both are
    /// internal, which the project's other modules still reach.
    /// </summary>
    private static string SlotAccess(VariableSymbol variable, bool exposed) =>
        exposed && variable.Type.Record is not { IsPublic: false } ? "public" : "internal";

    /// <summary>The struct of a class instance's block: its variables' storage, each with its initial value, which the constructor hands the runtime.</summary>
    private void EmitInstanceBlock()
    {
        if (!HasInstanceBlock)
        {
            return;
        }

        writer.HiddenLine("public struct __Fields");
        writer.Open();
        foreach (var variable in Symbol.Variables.Concat(Symbol.StaticLocals).Where(InInstanceBlock))
        {
            writer.HiddenLine($"{SlotAccess(variable, exposed: true)} {StorageType(variable)} {variable.EmitName} = {InitialValue(variable)};");
        }

        writer.Line();
        writer.HiddenLine("public __Fields()");
        writer.Open();
        writer.Close();
        writer.Close();
        writer.Line();
    }

    /// <summary>
    /// What the locals window shows under an instance (ARCHITECTURE.md section 9): the class's
    /// variables by value, as VBA's locals window lists an object's, in place of the runtime's
    /// own members. The debugger does not apply a type's display to a ref-returning property, so
    /// the variables in the instance's block would otherwise show as their type names.
    /// </summary>
    private void EmitDebugView()
    {
        writer.HiddenLine("private sealed class __DebugView");
        writer.Open();
        writer.HiddenLine($"private readonly global::{Symbol.ClassEmitName} __instance;");
        writer.Line();
        writer.HiddenLine($"public __DebugView(global::{Symbol.ClassEmitName} instance) => __instance = instance;");
        foreach (var variable in Symbol.Variables)
        {
            writer.Line();
            writer.HiddenLine($"public {StorageType(variable)} {variable.EmitName} => __instance.{variable.EmitName};");
        }

        writer.Close();
        writer.Line();
    }

    /// <summary>A variable holding object references that go when its scope ends (ARCHITECTURE.md D18).</summary>
    private static bool NeedsRelease(VariableSymbol variable) => variable.Kind != VariableKind.Static && !variable.IsAlias && OwnsStorage(variable.Type);

    /// <summary>A value of the type owns something its storage must release: an object reference, a BSTR, a record's references, or an array, which is native memory since M7 C4 (D18, D20).</summary>
    private static bool OwnsStorage(VbaType type) =>
        type.IsObject || type.IsVariant || type.IsVariableString || type.IsRecord || type.IsArray;

    /// <summary>End (MS-VBAL 5.4.2.15): the project is marked for its reset before the frames unwind, so their locals go without Class_Terminate (docs/vba-quirks.md).</summary>
    private string EndCode() =>
        $"throw {R}Hosting.ProjectReset.Ending(typeof(global::{(Symbol.Kind == ModuleKind.Class ? Symbol.ClassEmitName : Symbol.EmitName)}));";

    /// <summary>
    /// The module's reset (ProjectReset): its storage back to its initial values, as VBA leaves it
    /// after End or when the workbook closes (docs/vba-quirks.md). A class module's variables live
    /// in its instances, so its reset covers only the predeclared instance; a document module's
    /// object belongs to the host and is left alone.
    /// </summary>
    private void EmitReset()
    {
        List<VariableSymbol> storage = Symbol.Kind == ModuleKind.Class
            ? (Symbol.Instance is { } instance ? [instance] : [])
            : [.. Symbol.Variables, .. Symbol.StaticLocals];
        if (storage.Count == 0)
        {
            return;
        }

        writer.HiddenLine($"public static void {VbaNg.Runtime.Hosting.ProjectReset.MethodName}()");
        writer.Open();
        foreach (var variable in storage)
        {
            if (OwnsStorage(variable.Type))
            {
                writer.HiddenLine(ReleaseCode(variable, variable.EmitName));
            }

            // An object, Variant, or String slot is empty once released; anything else takes its initial value again.
            if (!variable.Type.IsObject && !variable.Type.IsVariant && !variable.Type.IsVariableString)
            {
                writer.HiddenLine($"{variable.EmitName} = {DefaultValue(variable)};");
            }
        }

        writer.Close();
        writer.Line();
    }

    private string ReleaseCode(VariableSymbol variable, string name)
    {
        if (variable.IsWithEvents)
        {
            return SubscribeCode(variable, name, "null") + ";";
        }

        return variable.Type.IsRecord ? $"{name}.ReleaseReferences();"
            : variable.Type.IsArray ? $"{R}ObjectRefs.Release(ref {name});"
            : $"{R}ObjectRefs.Release(ref {name});";
    }

    /// <summary>
    /// The dispid table (MS-VBAL 5.6.12): names resolve case-insensitively, the default member
    /// answers to 0 and the enumerator to -4, and every other public member takes the next number
    /// in declaration order.
    /// </summary>
    private void EmitDispatch()
    {
        var members = DispatchMembers();
        writer.HiddenLine("public override int GetDispId(string name)");
        writer.Open();
        writer.HiddenLine("switch (name.ToUpperInvariant())");
        writer.Open();
        foreach (var (member, dispId) in members)
        {
            writer.HiddenLine($"case {Quote(member.Name.ToUpperInvariant())}: return {dispId.ToString(CultureInfo.InvariantCulture)};");
        }

        writer.HiddenLine("default: return NoMember();");
        writer.Close();
        writer.Close();
        writer.Line();

        writer.HiddenLine($"public override {VariantType} Invoke(int dispId, {R}InvokeKind kind, global::System.ReadOnlySpan<{VariantType}> arguments)");
        writer.Open();
        writer.HiddenLine("switch (dispId)");
        writer.Open();
        foreach (var (member, dispId) in members)
        {
            if (InvokeCode(member) is not { } code)
            {
                continue;
            }

            writer.HiddenLine($"case {dispId.ToString(CultureInfo.InvariantCulture)}:");
            writer.Indent++;
            writer.HiddenLine($"if ((kind & ({R}InvokeKind.Method | {R}InvokeKind.PropertyGet)) != 0) {code}");
            writer.HiddenLine("break;");
            writer.Indent--;
        }

        writer.Close();
        writer.HiddenLine("return NotSupported();");
        writer.Close();
        writer.Line();

        writer.HiddenLine($"public override void Put(int dispId, global::System.ReadOnlySpan<{VariantType}> indices, in {VariantType} value, bool asReference)");
        writer.Open();
        writer.HiddenLine("switch (dispId)");
        writer.Open();
        foreach (var (member, dispId) in members)
        {
            if (PutCode(member) is not { } code)
            {
                continue;
            }

            writer.HiddenLine($"case {dispId.ToString(CultureInfo.InvariantCulture)}: {code} return;");
        }

        writer.Close();
        writer.HiddenLine("NotSupported();");
        writer.Close();
        writer.Line();

        if (Symbol.EnumMember is { } enumerator)
        {
            var call = new Emitted($"{enumerator.EmitName}()", enumerator.ReturnType ?? VbaType.Variant);
            writer.HiddenLine($"public override global::System.Collections.Generic.IEnumerator<{VariantType}> Enumerate() => {R}ObjectRefs.Enumerator({Convert(call, VbaType.Variant)});");
            writer.Line();
        }
    }

    /// <summary>
    /// The procedures <c>Application.Run</c> reaches by name in a standard or document module
    /// (ARCHITECTURE.md D23; docs/vba-quirks.md, "Host, Application.Run"): its Subs and Functions,
    /// private ones included. The arguments convert as in a late-bound call, and a conversion error
    /// or more arguments than parameters (450) reaches the caller, while an error the procedure
    /// leaves unhandled escapes the caller's handlers, as behind VBA's run-time error dialog. A
    /// procedure with a user-defined type parameter or result is left out: no Variant carries one.
    /// </summary>
    private void EmitRunTable()
    {
        var runnable = module.Procedures.Select(p => p.Symbol)
            .Where(p => p.Kind is ProcedureKind.Sub or ProcedureKind.Function && !p.Parameters.Any(q => HoldsRecord(q.Type)) && !(p.ReturnType is { } result && HoldsRecord(result)))
            .ToList();
        writer.HiddenLine("public static int __RunId(string __name)");
        writer.Open();
        writer.HiddenLine("switch (__name.ToUpperInvariant())");
        writer.Open();
        for (var i = 0; i < runnable.Count; i++)
        {
            writer.HiddenLine($"case {Quote(runnable[i].Name.ToUpperInvariant())}: return {i.ToString(CultureInfo.InvariantCulture)};");
        }

        writer.HiddenLine("default: return -1;");
        writer.Close();
        writer.Close();
        writer.Line();

        writer.HiddenLine($"public static {VariantType} __Run(int __id, global::System.ReadOnlySpan<{VariantType}> arguments)");
        writer.Open();
        writer.HiddenLine("switch (__id)");
        writer.Open();
        for (var i = 0; i < runnable.Count; i++)
        {
            var procedure = runnable[i];
            writer.HiddenLine($"case {i.ToString(CultureInfo.InvariantCulture)}:");
            writer.Open();
            if (!procedure.Parameters.Any(p => p.IsParamArray))
            {
                writer.HiddenLine($"{R}Hosting.ApplicationRun.Arity(arguments, {procedure.Parameters.Count.ToString(CultureInfo.InvariantCulture)});");
            }

            // The arguments convert before the call, so their errors reach the caller.
            var arguments = new List<string>();
            for (var j = 0; j < procedure.Parameters.Count; j++)
            {
                var parameter = procedure.Parameters[j];
                var name = "__a" + j.ToString(CultureInfo.InvariantCulture);
                writer.HiddenLine($"{(parameter.IsParamArray ? "var" : parameter.Type.CSharpName)} {name} = {LateValue(parameter, j)};");
                arguments.Add(parameter.EmitByRef && !parameter.IsParamArray ? ByRefCell(name, parameter) : name);
            }

            var call = $"{procedure.EmitName}({string.Join(", ", arguments)})";
            var body = procedure.ReturnsValue
                ? $"return {Convert(new Emitted(call, procedure.ReturnType ?? VbaType.Variant), VbaType.Variant)};"
                : $"{call}; return {VariantType}.Empty;";
            writer.HiddenLine($"try {{ {body} }}");
            writer.HiddenLine($"catch (global::System.Exception __ex) when ({R}Hosting.ApplicationRun.Escapes(__ex)) {{ throw {R}Hosting.ApplicationRun.Unhandled(__ex); }}");
            writer.Close();
        }

        writer.Close();
        writer.HiddenLine("throw new global::System.ArgumentOutOfRangeException(nameof(__id));");
        writer.Close();
        writer.Line();
    }

    private static bool HoldsRecord(VbaType type) => type.IsRecord || (type.IsArray && type.ElementType is { IsRecord: true });

    /// <summary>Reading a member through a dispid: the field's value, or the call with its arguments taken from the late-bound list.</summary>
    private string? InvokeCode(Symbol member)
    {
        switch (member)
        {
            case VariableSymbol field:
                return $"return {Convert(new Emitted(IsObjectSlot(field) ? field.EmitName + ".Target" : field.EmitName, field.Type), VbaType.Variant)};";
            case ProcedureSymbol { ReturnsValue: false } sub:
                return $"{{ {sub.EmitName}({LateArguments(sub)}); return {VariantType}.Empty; }}";
            case ProcedureSymbol function:
                return $"return {Convert(new Emitted($"{function.EmitName}({LateArguments(function)})", function.ReturnType ?? VbaType.Variant), VbaType.Variant)};";
            case PropertySymbol { Get: { } get }:
                return $"return {Convert(new Emitted($"{get.EmitName}({LateArguments(get)})", get.ReturnType ?? VbaType.Variant), VbaType.Variant)};";
            default:
                return null;
        }
    }

    /// <summary>Writing a member through a dispid: a public field, or the property's Let or Set with the index arguments before the value.</summary>
    private static string? PutCode(Symbol member)
    {
        switch (member)
        {
            case VariableSymbol field:
                return $"{StoreField(field, Convert(new Emitted("value", VbaType.Variant), field.Type))};";
            case PropertySymbol property when property.Let is not null || property.Set is not null:
                // DISPATCH_PROPERTYPUTREF (a Set) reaches the Set accessor and DISPATCH_PROPERTYPUT (a Let) the Let accessor when the property has both (Classes golden).
                return property is { Let: { } letAccessor, Set: { } setAccessor }
                    ? $"if (asReference) {{ {PutCall(setAccessor)} }} else {{ {PutCall(letAccessor)} }}"
                    : PutCall(property.Let ?? property.Set!);
            default:
                return null;
        }
    }

    /// <summary>A property accessor called through Put: its index arguments from the late-bound list, then the value.</summary>
    private static string PutCall(ProcedureSymbol accessor)
    {
        var arguments = new List<string>();
        for (var i = 0; i < accessor.Parameters.Count - 1; i++)
        {
            var parameter = accessor.Parameters[i];
            var index = Convert(new Emitted($"{R}ObjectRefs.At(indices, {i.ToString(CultureInfo.InvariantCulture)})", VbaType.Variant), parameter.Type);
            // A late-bound Put passes values, so a ByRef index parameter takes a temporary, as LateArguments does.
            arguments.Add(parameter.IsByVal ? index : ByRefCell(index, parameter));
        }

        var last = accessor.Parameters[^1];
        var converted = Convert(new Emitted("value", VbaType.Variant), last.Type);
        arguments.Add(last.IsByVal ? converted : ByRefCell(converted, last));
        return $"{accessor.EmitName}({string.Join(", ", arguments)});";
    }

    private static string StoreField(VariableSymbol field, string value) =>
        field.Type.IsObject || field.Type.IsVariant || field.Type.IsVariableString
            ? $"{R}ObjectRefs.Assign(ref {field.EmitName}, {value})"
            : field.Type.IsArray ? $"{R}ObjectRefs.AssignArray(ref {field.EmitName}, {value})"
            : $"{field.EmitName} = {value}";

    /// <summary>A late-bound call's cell for a ByRef parameter, since it passes values: an object slot's for an object (ROADMAP.md M7 E3), a Variant's, a String's, an array's, or a scalar's otherwise.</summary>
    private static string ByRefCell(string value, ParameterSymbol parameter) =>
        IsObjectSlot(parameter) ? $"ref {R}ObjectRefs.Cell<{ClassName(parameter.Type)}>({value})" : $"ref {R}ObjectRefs.Slot({value})";

    /// <summary>The arguments of a late-bound call: each parameter reads its position, an omitted optional takes its declared default, and a ParamArray takes the rest.</summary>
    private string LateArguments(ProcedureSymbol procedure) =>
        // A late-bound call passes values, so a ByRef parameter takes a temporary (ROADMAP.md backlog).
        string.Join(", ", procedure.Parameters.Select((parameter, i) => parameter.IsParamArray || parameter.IsByVal ? LateValue(parameter, i) : ByRefCell(LateValue(parameter, i), parameter)));

    /// <summary>A parameter's value in a late-bound call: its position in the list, its declared default when it was left out, or the rest of the list for a ParamArray.</summary>
    private string LateValue(ParameterSymbol parameter, int position)
    {
        var index = position.ToString(CultureInfo.InvariantCulture);
        if (parameter.IsParamArray)
        {
            return $"{R}VbaArray.FromValues({R}VarType.Variant, {R}ObjectRefs.Rest(arguments, {index}), 0)";
        }

        var value = Convert(new Emitted($"{R}ObjectRefs.At(arguments, {index})", VbaType.Variant), parameter.Type);
        if (parameter.IsOptional)
        {
            var fallback = parameter.Default is { } constant
                ? Convert(Literal(constant), parameter.Type)
                : parameter.Type.IsVariant ? VariantType + ".Missing" : DefaultValue(parameter);
            value = $"arguments.Length > {index} ? {value} : {fallback}";
        }

        return value;
    }

    /// <summary>
    /// RaiseEvent name(arguments) (MS-VBAL 5.4.2.10): the arguments travel as a Variant array, so a
    /// handler's ByRef parameter can write back into the raising procedure's variable.
    /// </summary>
    private string RaiseEvent(BoundRaiseEvent raise)
    {
        var arguments = raise.Arguments.Select(a => Convert(EmitExpression(a), VbaType.Variant)).ToList();
        if (arguments.Count == 0)
        {
            return $"RaiseVbaEvent({Quote(raise.Raised.Name)}, []);";
        }

        var array = NewTemporary();
        before.Add($"var {array} = new {VariantType}[] {{ {string.Join(", ", arguments)} }};");
        for (var i = 0; i < raise.Arguments.Count; i++)
        {
            if (raise.Arguments[i].IsLValue)
            {
                after.Add(Store(raise.Arguments[i], new Emitted($"{array}[{i.ToString(CultureInfo.InvariantCulture)}]", VbaType.Variant)) + ";");
            }
        }

        return $"RaiseVbaEvent({Quote(raise.Raised.Name)}, {array});";
    }

    /// <summary>TypeOf expression Is type (MS-VBAL 5.6.9.10).</summary>
    private Emitted TypeOf(BoundTypeOf typeOf)
    {
        var value = Convert(EmitExpression(typeOf.Value), VbaType.Variant);
        if (typeOf.Tested.IsGenericObject)
        {
            return new Emitted($"{R}ObjectRefs.IsAnyObject({value})", VbaType.Boolean);
        }

        if (typeOf.Tested.IsCom)
        {
            // A type library's type: the object is asked for the type's interface, a coclass's default one, as VBA's TypeOf asks (Objects golden).
            return new Emitted($"{R}ObjectRefs.IsOfInterface({value}, new global::System.Guid({Quote(typeOf.Tested.ComInterface!.Guid.ToString("D"))}))", VbaType.Boolean);
        }

        return new Emitted($"{R}ObjectRefs.IsOfType<{ClassName(typeOf.Tested)}>({value})", VbaType.Boolean);
    }

    /// <summary>The C# expression that creates an instance of a class (MS-VBAL 5.6.16.2 New).</summary>
    private static string NewInstance(VbaType type) =>
        type.ProjectClass is { } projectClass ? $"new global::{projectClass.ClassEmitName}()" : $"new {ClassName(type)}()";

    /// <summary>The explicit implementations that route an interface's members to the <c>Interface_Member</c> procedures of this class (MS-VBAL 5.2.4.2).</summary>
    private void EmitInterfaceForwarders()
    {
        var forwarded = Symbol.Procedures.Where(p => p.ImplementsInterface is not null).ToList();
        foreach (var procedure in forwarded)
        {
            var owner = procedure.ImplementsInterface!;
            var member = owner.Members.GetValueOrDefault(procedure.ImplementsMember!);
            var target = procedure.Kind switch
            {
                ProcedureKind.PropertyGet => (member as PropertySymbol)?.Get,
                ProcedureKind.PropertyLet => (member as PropertySymbol)?.Let,
                ProcedureKind.PropertySet => (member as PropertySymbol)?.Set,
                _ => member as ProcedureSymbol,
            };
            if (target is null)
            {
                Report(procedure.Syntax, $"'{owner.Name}' has no member '{procedure.ImplementsMember}' for '{procedure.Name}' to implement.");
                continue;
            }

            var parameters = string.Join(", ", procedure.Parameters.Select(ParameterDeclaration));
            var arguments = string.Join(", ", procedure.Parameters.Select(p => (p.IsByVal || p.IsParamArray ? string.Empty : "ref ") + p.EmitName));
            var returnType = procedure.ReturnType?.CSharpName ?? "void";
            writer.HiddenLine($"{returnType} global::{owner.EmitName}.{target.EmitName}({parameters}) => {procedure.EmitName}({arguments});");
        }

        if (forwarded.Count > 0)
        {
            writer.Line();
        }
    }

    /// <summary>The C# interface a class gets when another class implements it; its members are the class's own public procedures.</summary>
    private void EmitClassInterface()
    {
        writer.Line($"public interface {Symbol.EmitName} : {R}IVbaClassInstance");
        writer.Line("{");
        writer.Indent++;
        foreach (var procedure in Symbol.Procedures.Where(p => p.IsPublic && p.ImplementsInterface is null && p.EventSource is null))
        {
            var parameters = string.Join(", ", procedure.Parameters.Select(ParameterDeclaration));
            writer.HiddenLine($"{procedure.ReturnType?.CSharpName ?? "void"} {procedure.EmitName}({parameters});");
        }

        writer.Indent--;
        writer.Line("}");
        writer.Hidden();
        writer.Line();
    }

    /// <summary>The base list of a class module: the runtime base, the interface of its own name when another class implements it, then every interface it implements.</summary>
    private string ClassBases()
    {
        var bases = new List<string> { HasInstanceBlock ? $"{R}VbaClassObject<global::{Symbol.ClassEmitName}.__Fields>" : R + "VbaClassObject" };
        if (Symbol.IsImplemented)
        {
            bases.Add("global::" + Symbol.EmitName);
        }

        foreach (var implemented in Symbol.Implemented)
        {
            bases.Add("global::" + implemented.EmitName);
        }

        if (Symbol.EventSources.Count > 0)
        {
            bases.Add(R + "IVbaEventSink");
        }

        return string.Join(", ", bases);
    }

    /// <summary>
    /// Set on a WithEvents variable (MS-VBAL 5.2.3.1.4): the subscription moves with the value. A
    /// project class keeps its subscribers; a library type is advised through the runtime, the
    /// advisory kept next to the variable. The sink is the class instance, or a document module's
    /// nested stand-in.
    /// </summary>
    private string SubscribeCode(VariableSymbol variable, string slot, string value)
    {
        var sink = Symbol.Kind == ModuleKind.Class ? "this" : "__events";
        return variable.EventInterface is null
            ? $"{R}ObjectRefs.Subscribe(ref {slot}, {value}, {sink}, {Quote(variable.Name)})"
            : $"{R}ObjectRefs.Subscribe(ref {slot}, {value}, {sink}, {Quote(variable.Name)}, {EventsName(variable)}, ref {AdvisoryName(variable)})";
    }

    private static string EventsName(VariableSymbol variable) => "__" + variable.EmitName.TrimStart('@') + "_Events";

    private static string AdvisoryName(VariableSymbol variable) => "__" + variable.EmitName.TrimStart('@') + "_Advisory";

    /// <summary>The handler table of a module with WithEvents variables: an event of a source runs the <c>source_Event</c> procedure, writing ByRef parameters back.</summary>
    private void EmitEventSink()
    {
        if (Symbol.EventSources.Count == 0)
        {
            return;
        }

        writer.HiddenLine($"void {R}IVbaEventSink.RaiseVbaEvent(string source, string name, {VariantType}[] arguments)");
        writer.Open();
        writer.HiddenLine("switch ((source + \".\" + name).ToUpperInvariant())");
        writer.Open();
        foreach (var handler in Symbol.Procedures.Where(p => p.EventVariable is not null))
        {
            var arguments = new List<string>();
            var declarations = new List<string>();
            var copies = new List<string>();
            for (var i = 0; i < handler.Parameters.Count; i++)
            {
                var parameter = handler.Parameters[i];
                var index = i.ToString(CultureInfo.InvariantCulture);
                var value = Convert(new Emitted($"{R}ObjectRefs.At(arguments, {index})", VbaType.Variant), parameter.Type);
                if (parameter.IsByVal)
                {
                    arguments.Add(value);
                    continue;
                }

                if (IsObjectSlot(parameter))
                {
                    // A ByRef object parameter works on a slot with a reference of its own; what the handler left there goes back as a temporary of the raising statement (ROADMAP.md M7 E3).
                    declarations.Add($"var __arg{index} = {R}ObjectRefs.Holding<{ClassName(parameter.Type)}>({value});");
                    copies.Add($"if (arguments.Length > {index}) arguments[{index}] = {R}ObjectRefs.Owned({VariantType}.FromObject(__arg{index}.Target));");
                    copies.Add($"{R}ObjectRefs.Release(ref __arg{index});");
                    arguments.Add("ref __arg" + index);
                    continue;
                }

                declarations.Add($"var __arg{index} = {value};");
                copies.Add($"if (arguments.Length > {index}) arguments[{index}] = {Convert(new Emitted("__arg" + index, parameter.Type), VbaType.Variant)};");
                arguments.Add("ref __arg" + index);
            }

            writer.HiddenLine($"case {Quote((handler.EventSource + "." + handler.EventName).ToUpperInvariant())}:");
            writer.Open();
            foreach (var declaration in declarations)
            {
                writer.HiddenLine(declaration);
            }

            writer.HiddenLine($"{handler.EmitName}({string.Join(", ", arguments)});");
            foreach (var copy in copies)
            {
                writer.HiddenLine(copy);
            }

            writer.HiddenLine("return;");
            writer.Close();
        }

        writer.Close();
        writer.Close();
        writer.Line();
    }
}
