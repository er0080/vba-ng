using System.Globalization;
using System.Text.RegularExpressions;

using VbaNg.Compiler.Binding;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Emit;

/// <summary>A piece of generated C# and the type its value has; <see cref="VbaType.CSharpName"/> of the type is the C# type of the code.</summary>
internal readonly record struct Emitted(string Code, VbaType Type);

/// <summary>Expression emission: every operator goes through the runtime's Variant operators, and values cross type boundaries through Coerce (MS-VBAL 5.5.1, 5.6.9).</summary>
public sealed partial class CSharpEmitter
{
    private static readonly Regex Placeholder = new(@"\{(\w+)(?::([a-z0-9]+))?(?:\|([^}]*))?\}", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private Emitted EmitExpression(BoundExpression expression)
    {
        switch (expression)
        {
            case BoundLiteral literal:
                return Literal(literal.Value);
            case BoundConstant constant:
                return Literal(constant.Constant.Value);
            case BoundVariable variable:
                return Variable(variable.Variable, forStore: false);
            case BoundElement element:
                return Element(element);
            case BoundField { Field: var chars } inline when IsInlineCharsArray(chars):
                // A fixed-length string array member has no descriptor (ROADMAP.md M7 C10): the whole array is a copy of its elements for the statement.
                return new Emitted($"{R}VbaArray.FromInlineChars({FieldOwner(inline)}.{chars.EmitName})", chars.Type);
            case BoundField { Field.IsEmbeddedArray: true } embedded:
                // A fixed-size array member lies inline in its record (ROADMAP.md M7 C6): a view of it for the statement.
                return new Emitted($"{R}VbaArray.Embedded(ref {EmbeddedOwner(embedded)}.{embedded.Field.EmitName}, {Bounds(embedded.Field.Bounds!)})", embedded.Field.Type);
            case BoundField { Field: var member } dynamic when IsDynamicArrayMember(member):
                // A dynamic array member is its descriptor pointer (ROADMAP.md M7 C8): a view of the array, of the member's element type.
                return new Emitted($"{FieldOwner(dynamic)}.{member.EmitName}.View({VarTypeName(member.Type.ElementType!)})", member.Type);
            case BoundField field:
                return IsInlineChars(field.Field)
                    ? new Emitted($"{L}Strings.FromChars({FieldOwner(field)}.{field.Field.EmitName})", field.Field.Type)
                    : new Emitted($"{FieldOwner(field)}.{field.Field.EmitName}{(IsObjectSlot(field.Field) ? ".Target" : string.Empty)}", field.Field.Type);
            case BoundCall call:
                return new Emitted(CallCode(call, null), call.Procedure.ReturnType ?? VbaType.Variant);
            case BoundIntrinsicCall intrinsic:
                return Intrinsic(intrinsic);
            case BoundArrayFunction array:
                return new Emitted($"{L}Interaction.Array({Arguments(array.Items)}, {array.OptionBase.ToString(CultureInfo.InvariantCulture)})", VbaType.Variant);
            case BoundLenOfDeclared { Declared.Kind: TypeKind.FixedString } len:
                // LenB of a fixed-length String variable, two bytes a character, as an Integer (Memory golden); of an element of an array of them, read first (Arrays golden).
                return len.Operand is { } measured
                    ? new Emitted($"{VariantType}.FromInt16({R}Coerce.ToInt16({L}Strings.LenB({Convert(EmitExpression(measured), VbaType.Variant)})))", VbaType.Variant)
                    : new Emitted($"{VariantType}.FromInt16((short){(len.Declared.FixedLength * 2).ToString(CultureInfo.InvariantCulture)})", VbaType.Variant);
            case BoundLenOfDeclared len:
                return new Emitted($"{L}Strings.LenOfDeclared({VarTypeName(len.Declared)})", VbaType.Variant);
            case BoundPointer pointer:
                return new Emitted(PointerCode(pointer), VbaType.LongLong);
            case BoundRecordLength length:
                return new Emitted($"{L}FileSystem.RecordLength{(length.Bytes ? "B" : string.Empty)}({EmitExpression(length.Record).Code})", VbaType.Integer);
            case BoundMember member:
                return Member(member);
            case BoundErrObject:
                return new Emitted(R + "Err.Current", VbaType.ErrObject);
            case BoundErl:
                return new Emitted(R + "Err.Current.Line", VbaType.Long);
            case BoundUnary unary:
                return Unary(unary);
            case BoundBinary binary:
                return Binary(binary);
            case BoundConversion conversion:
                return new Emitted(Convert(EmitExpression(conversion.Operand), conversion.Type), conversion.Type);
            case BoundTypeOf typeOf:
                return TypeOf(typeOf);
            case BoundAddressOf addressOf:
                return new Emitted(AddressOf(addressOf.Procedure), VbaType.LongLong);
            case BoundNewObject created:
                return created.Type.IsCom
                    ? new Emitted($"{R}Com.CreateInstance({ComClassId(created.Type)})", created.Type)
                    : new Emitted(Owned(NewInstance(created.Type)), created.Type);
            case BoundComCall call:
                return ComCall(call);
            case BoundAppObject app:
                return AppObject(app);
            case BoundParenthesized parenthesized:
                {
                    var inner = EmitExpression(parenthesized.Operand);
                    return new Emitted($"({inner.Code})", inner.Type);
                }

            default:
                Report(expression.Syntax, $"{expression.GetType().Name} is not supported yet.");
                return new Emitted(VariantType + ".Empty", VbaType.Variant);
        }
    }

    private Emitted Variable(VariableSymbol variable, bool forStore)
    {
        var name = variable.Kind == VariableKind.Module && !ReferenceEquals(variable.Module, Symbol)
            ? $"global::{variable.Module!.EmitName}.{variable.EmitName}"
            : StorageName(variable);
        if (variable.IsNew && !forStore)
        {
            // As New, and the predeclared instance of a class: the object is created on first use, again after Nothing (MS-VBAL 5.2.3.1.1, Classes golden).
            var created = variable.Type.IsCom ? $"{R}Com.CreateInstance({ComClassId(variable.Type)})" : NewInstance(variable.Type);
            return new Emitted($"{R}ObjectRefs.AutoNew(ref {name}, () => {created})", variable.Type);
        }

        return new Emitted(!forStore && IsObjectSlot(variable) ? name + ".Target" : name, variable.Type);
    }

    private Emitted Element(BoundElement element)
    {
        if (element.Array.Type.IsArray)
        {
            var indices = string.Join(", ", element.Indices.Select(i => Convert(EmitExpression(i), VbaType.Long)));
            if (element.Array is BoundField { Field: var chars } member && IsInlineCharsArray(chars))
            {
                // An element of a fixed-length string array member, read from the record's own characters (ROADMAP.md M7 C10).
                return new Emitted($"{R}VbaArray.InlineChars(ref {EmbeddedOwner(member)}.{chars.EmitName}, [{indices}])", chars.Type.ElementType!);
            }

            var array = EmitExpression(element.Array).Code;
            if (element.Type.IsRecord)
            {
                // An element of an array of records is its storage, reached by ref, so fields written through it change the element (ROADMAP.md M7 C3).
                return new Emitted($"{array}.Record<{element.Type.CSharpName}>([{indices}])", element.Type);
            }

            return new Emitted($"{array}.Get([{indices}])", VbaType.Variant);
        }

        return new Emitted($"{R}LateBound.Index({Convert(EmitExpression(element.Array), VbaType.Variant)}, {Arguments(element.Indices)})", VbaType.Variant);
    }

    /// <summary>The C# that stores a value into an assignable expression (no trailing semicolon).</summary>
    private string Store(BoundExpression target, Emitted value)
    {
        switch (target)
        {
            case BoundVariable variable:
                {
                    var slot = Variable(variable.Variable, forStore: true).Code;
                    var stored = variable.Variable.Type.IsVariant ? StoredVariant(value) : Convert(value, variable.Variable.Type);
                    if (variable.Variable.IsWithEvents)
                    {
                        // Assigning a WithEvents variable moves the subscription with it (MS-VBAL 5.2.3.1.4).
                        return SubscribeCode(variable.Variable, slot, stored);
                    }

                    return Assignment(variable.Variable.Type, slot, stored);
                }
            case BoundElement { Array: BoundField { Field: var chars } member } element when IsInlineCharsArray(chars):
                {
                    // A store into an element of a fixed-length string array member writes over the record's own characters (ROADMAP.md M7 C10).
                    var indices = string.Join(", ", element.Indices.Select(i => Convert(EmitExpression(i), VbaType.Long)));
                    return $"{R}VbaArray.SetInlineChars(ref {EmbeddedOwner(member)}.{chars.EmitName}, [{indices}], {Convert(value, chars.Type.ElementType!)})";
                }

            case BoundElement { Array.Type.IsArray: true, Type.IsRecord: true } element:
                // An element of an array of records is storage of its own: it takes a copy, and what it held goes (MS-VBAL 5.5.1.2.5).
                return $"{R}ObjectRefs.AssignRecord(ref {Element(element).Code}, {Convert(value, element.Type)})";
            case BoundElement { Array.Type.IsArray: true } element:
                {
                    var indices = string.Join(", ", element.Indices.Select(i => Convert(EmitExpression(i), VbaType.Long)));

                    // An element of a fixed-length string array takes the value padded or cut to its length (MS-VBAL 5.4.3.1; Arrays golden).
                    var stored = element.Array.Type.ElementType is { Kind: TypeKind.FixedString } fixedString
                        ? StoredVariant(new Emitted(Convert(value, fixedString), fixedString))
                        : StoredVariant(value);
                    return $"{EmitExpression(element.Array).Code}.Set([{indices}], {stored})";
                }

            case BoundElement element:
                // Set through a late-bound default member hands the object itself (DISPATCH_PROPERTYPUTREF), Let the value as it is, an object
                // included, which the member takes or refuses itself; the runtime let-coerces for an array element (Objects golden).
                return $"{R}LateBound.{(storeIsSet ? "SetIndexReference" : "SetIndex")}({Convert(EmitExpression(element.Array), VbaType.Variant)}, {Arguments(element.Indices)}, {LetVariant(value)})";
            case BoundField field:
                {
                    if (IsInlineChars(field.Field))
                    {
                        // A fixed-length string member lies inline in its type (ROADMAP.md M7 C5): the padded value is written over its characters.
                        return $"{L}Strings.SetChars(ref {FieldOwner(field)}.{field.Field.EmitName}, {Convert(value, field.Field.Type)})";
                    }

                    var slot = $"{FieldOwner(field)}.{field.Field.EmitName}";
                    if (IsDynamicArrayMember(field.Field))
                    {
                        // A dynamic array member takes a copy of its own and destroys the array it held, through its descriptor pointer (ROADMAP.md M7 C8).
                        return $"{slot} = {R}VbaArrayMember.Assign({slot}, {VarTypeName(field.Field.Type.ElementType!)}, {Convert(value, field.Field.Type)}.Clone())";
                    }

                    var stored = field.Field.Type.IsVariant ? StoredVariant(value) : Convert(value, field.Field.Type);
                    return Assignment(field.Field.Type, slot, stored);
                }
            case BoundMember member when ReferenceEquals(member.Target.Type, VbaType.ErrObject):
                return $"{ErrTarget(member.Target)}.{member.Member} = {Convert(value, member.Type)}";
            case BoundMember member when member.Target.Type.IsVariant || member.Target.Type.IsGenericObject:
                // Set through a late-bound member hands the object itself (DISPATCH_PROPERTYPUTREF), Let the value as it is, an object included,
                // which the member takes or refuses itself (Objects golden: d.Key(k) = obj keys the entry by the object, d.Item(k) = obj raises 450).
                return $"{R}LateBound.{(storeIsSet ? "Set" : "Let")}({Convert(EmitExpression(member.Target), VbaType.Variant)}, {Quote(member.Member)}, {Arguments(member.Arguments)}{Names(member)}, {LetVariant(value)})";
            case BoundComCall call:
                return StoreCom(call, value, storeIsSet);
            default:
                Report(target.Syntax, "This expression cannot be assigned to.");
                return "_ = " + value.Code;
        }
    }

    private Emitted Member(BoundMember member)
    {
        var target = member.Target;
        if (ReferenceEquals(target.Type, VbaType.ErrObject))
        {
            var err = ErrTarget(target);
            return member.Member switch
            {
                "Raise" or "Clear" => new Emitted($"{err}.{member.Member}({string.Join(", ", member.Arguments.Select(a => Convert(EmitExpression(a), VbaType.Variant)))})", VbaType.Variant),
                _ => new Emitted($"{err}.{member.Member}", member.Type),
            };
        }

        if (ReferenceEquals(target.Type, VbaType.Collection))
        {
            var collection = $"{R}Coerce.Require({EmitExpression(target).Code})";
            var arguments = string.Join(", ", member.Arguments.Select(a => Convert(EmitExpression(a), VbaType.Variant)));
            return member.Member switch
            {
                "Count" => new Emitted($"{collection}.Count", VbaType.Long),
                "Item" => new Emitted($"{collection}.Item({arguments})", VbaType.Variant),
                "NewEnum" => new Emitted($"{collection}.NewEnum()", VbaType.Object),
                _ => new Emitted($"{collection}.{member.Member}({arguments})", VbaType.Variant),
            };
        }

        var read = member.IsForeign && member.NamedArguments.Count == 0 ? "Foreign" : "Get";
        return new Emitted($"{R}LateBound.{read}({Convert(EmitExpression(target), VbaType.Variant)}, {Quote(member.Member)}, {Arguments(member.Arguments)}{Names(member)})", VbaType.Variant);
    }

    /// <summary>A slot that can hold an object reference, so a store into it counts (ARCHITECTURE.md D18).</summary>
    private static bool Counted(VbaType type) => type.IsObject || type.IsVariant || type.IsVariableString;

    /// <summary>
    /// A variable whose storage is an ObjectSlot, the interface pointer itself (ARCHITECTURE.md
    /// section 6, D20; ROADMAP.md M7 E3): every variable of an object type but a WithEvents
    /// variable, which keeps its subscription beside the object, Me, a class's predeclared
    /// instance, and a Declare's ByVal parameter, which hands the DLL the pointer itself.
    /// </summary>
    private static bool IsObjectSlot(VariableSymbol variable) =>
        variable.Type.IsObject && !variable.IsWithEvents && variable.EmitType is null
        && !ReferenceEquals(variable, variable.Module?.Me) && !ReferenceEquals(variable, variable.Module?.Instance)
        && variable is not ParameterSymbol { IsExternal: true, IsByVal: true };

    private static string SlotType(VbaType type) => $"{R}ObjectSlot<{ClassName(type)}>";

    /// <summary>The C# type of a variable's storage: an object slot, or its type's own C# type.</summary>
    private static string StorageType(VariableSymbol variable) =>
        IsObjectSlot(variable) ? SlotType(variable.Type) : IsInlineChars(variable) ? CharsType(variable.Type) : variable.IsEmbeddedArray || IsInlineCharsArray(variable) ? EmbeddedTypeName(variable) : IsBooleanMember(variable) ? R + "VbaBoolean" : IsDynamicArrayMember(variable) ? R + "VbaArrayMember" : variable.Type.CSharpName;

    /// <summary>A dynamic array member of a record, which is VBA's descriptor pointer (VbaArrayMember) rather than a VbaArray (ROADMAP.md M7 C8; Memory golden).</summary>
    private static bool IsDynamicArrayMember(VariableSymbol variable) => variable.Kind == VariableKind.Field && variable.Type.IsArray && variable.Bounds is null;

    /// <summary>A fixed-size array member of fixed-length strings, which lies inline in its record as its elements' characters, reached by index rather than through a descriptor (ROADMAP.md M7 C10; Memory golden).</summary>
    private static bool IsInlineCharsArray(VariableSymbol variable) => variable.Kind == VariableKind.Field && variable.Bounds is not null && variable.Type.ElementType is { Kind: TypeKind.FixedString };

    /// <summary>A Boolean member of a record, which is VBA's two bytes (VbaBoolean) rather than a bool (ROADMAP.md M7 C; Memory golden).</summary>
    private static bool IsBooleanMember(VariableSymbol variable) => variable.Kind == VariableKind.Field && ReferenceEquals(variable.Type, VbaType.Boolean);

    /// <summary>A fixed-length string member of a record, which lies inline in it as its UTF-16 characters (ROADMAP.md M7 C5; Memory golden).</summary>
    private static bool IsInlineChars(VariableSymbol variable) => variable.Kind == VariableKind.Field && variable.Type.Kind == TypeKind.FixedString;

    /// <summary>The inline array of characters the module declares for a fixed-length string member of its length.</summary>
    private static string CharsType(VbaType type) => "__Chars" + type.FixedLength.ToString(CultureInfo.InvariantCulture);

    /// <summary>The C# name of a variable's storage: a ByVal object parameter arrives as the object and lives in a slot of its own for the call.</summary>
    private static string StorageName(VariableSymbol variable) =>
        variable is ParameterSymbol { IsByVal: true } && IsObjectSlot(variable) ? variable.EmitName.TrimStart('@') + "__slot" : variable.EmitName;

    /// <summary>The value a variable's storage starts with.</summary>
    private static string InitialValue(VariableSymbol variable) => variable.IsEmbeddedArray && variable.Type.ElementType!.IsRecord ? $"{R}VbaArray.FreshEmbedded<{EmbeddedTypeName(variable)}>()" : IsObjectSlot(variable) || IsInlineChars(variable) || variable.IsEmbeddedArray || IsDynamicArrayMember(variable) || IsInlineCharsArray(variable) ? "default" : DefaultValue(variable);

    /// <summary>The inline storage a module declares for a fixed-size array member of its element type and bounds (ROADMAP.md M7 C6).</summary>
    private static string EmbeddedTypeName(VariableSymbol field) =>
        "__Inline_" + EmbeddedElementName(field.Type.ElementType!) + "_" + string.Join("_", field.Bounds!.Select(b => $"{BoundName(b.Lower)}To{BoundName(b.Upper)}"));

    /// <summary>The element type in an inline storage's name: a record's, as C# spells it, a fixed-length string's with its length, or the VarType.</summary>
    private static string EmbeddedElementName(VbaType element) =>
        element.IsRecord ? string.Concat(element.CSharpName.Replace("global::", string.Empty, StringComparison.Ordinal).Select(c => char.IsLetterOrDigit(c) ? c : '_'))
            : element.Kind == TypeKind.FixedString ? "String" + element.FixedLength.ToString(CultureInfo.InvariantCulture) : element.VarType.ToString();

    private static string BoundName(int bound) => bound < 0 ? "m" + (-(long)bound).ToString(CultureInfo.InvariantCulture) : bound.ToString(CultureInfo.InvariantCulture);

    /// <summary>The C# type of one element in an array's storage, as SafeArrayCreate lays it out: a Boolean is two bytes, a String and an object their pointers, a record itself.</summary>
    private static string EmbeddedElementType(VbaType element) => element.VarType switch
    {
        _ when element.IsRecord => element.CSharpName,
        _ when element.Kind == TypeKind.FixedString => CharsType(element),
        VarType.Byte => "byte",
        VarType.Integer or VarType.Boolean => "short",
        VarType.Long => "int",
        VarType.LongLong or VarType.Currency => "long",
        VarType.Single => "float",
        VarType.Double or VarType.Date => "double",
        VarType.Variant => VariantType,
        _ => "nint",
    };

    /// <summary>A fixed-size array member's owner as a C# reference: a record that is not a variable (a function's result) is copied into a temporary for the statement, whose view it then outlives.</summary>
    private string EmbeddedOwner(BoundField field)
    {
        var owner = FieldOwner(field);
        if (field.Record.IsLValue)
        {
            return owner;
        }

        var temporary = NewTemporary();
        before.Add($"var {temporary} = {owner};");
        return temporary;
    }

    /// <summary>A C# reference to an array's storage: a fixed-size array member's view is held in a temporary of the statement (ROADMAP.md M7 C6).</summary>
    private string ArrayReference(BoundExpression target)
    {
        var code = EmitExpression(target).Code;
        if (target is BoundField { Field: var chars } inline && IsInlineCharsArray(chars))
        {
            // A statement that works on a fixed-length string array member as a whole (a ByRef call) works on its copy, stored back after (ROADMAP.md M7 C10).
            var copy = NewTemporary();
            before.Add($"var {copy} = {code};");
            if (!inline.Record.Type.IsRecord || inline.Record.IsLValue)
            {
                after.Add($"{R}VbaArray.ToInlineChars(ref {FieldOwner(inline)}.{chars.EmitName}, {copy});");
            }

            return copy;
        }

        if (target is BoundField { Field: var member } field && IsDynamicArrayMember(member))
        {
            // A statement that can replace a dynamic array member's descriptor works on a view and stores it back (ROADMAP.md M7 C8).
            var array = NewTemporary();
            before.Add($"var {array} = {code};");
            if (!field.Record.Type.IsRecord || field.Record.IsLValue)
            {
                after.Add($"{FieldOwner(field)}.{member.EmitName} = {array};");
            }

            return array;
        }

        if (target is not BoundField { Field.IsEmbeddedArray: true })
        {
            return code;
        }

        var view = NewTemporary();
        before.Add($"var {view} = {code};");
        return view;
    }

    /// <summary>The value a temporary for a ByRef parameter starts with: an object slot holds a reference of its own (M7 E3), a Variant or String a copy.</summary>
    private static string RefValue(string value, ParameterSymbol parameter) =>
        IsObjectSlot(parameter) ? $"{R}ObjectRefs.Holding<{ClassName(parameter.Type)}>({value})" : Own(value, parameter.Type);

    /// <summary>The slot an expression is when it is the storage of a typed object variable of the given type, for a ref to it (M7 E3); null otherwise.</summary>
    private string? SlotReference(BoundExpression expression, VbaType type) => expression switch
    {
        BoundVariable { Variable: var variable } when IsObjectSlot(variable) && variable.Type.CSharpName == type.CSharpName => Variable(variable, forStore: true).Code,
        BoundField { Field: var field } access when IsObjectSlot(field) && field.Type.CSharpName == type.CSharpName => $"{FieldOwner(access)}.{field.EmitName}",
        _ => null,
    };

    /// <summary>A store into a slot: counted slots go through the runtime, which releases what the slot held (D18, D20); a record slot releases the old record's references, an array slot takes a copy of its own and destroys the array it held; anything else is a plain assignment.</summary>
    private static string Assignment(VbaType type, string slot, string stored) =>
        Counted(type) ? $"{R}ObjectRefs.Assign(ref {slot}, {stored})"
        : type.IsRecord ? $"{R}ObjectRefs.AssignRecord(ref {slot}, {stored})"
        : type.IsArray ? $"{R}ObjectRefs.AssignArray(ref {slot}, {stored}.Clone())"
        : $"{slot} = {stored}";

    /// <summary>The receiver of a field access: a record value, or the instance a class member belongs to, which must not be Nothing (MS-VBAL 5.6.12).</summary>
    private string FieldOwner(BoundField field)
    {
        var owner = EmitExpression(field.Record);
        return field.Record.Type.IsRecord ? owner.Code : Instance(owner, field.Record.Type);
    }

    /// <summary>The instance a class member is reached through: an array element or a Variant is converted to the class first, and Nothing raises 91 (MS-VBAL 5.6.12).</summary>
    private static string Instance(Emitted target, VbaType type) => $"{R}Coerce.Require({Convert(target, type)})";

    private string ErrTarget(BoundExpression target) =>
        target is BoundErrObject ? R + "Err.Current" : $"{R}Coerce.Require({EmitExpression(target).Code})";

    private string Arguments(IReadOnlyList<BoundExpression> arguments) =>
        "[" + string.Join(", ", arguments.Select(a => Convert(EmitExpression(a), VbaType.Variant))) + "]";

    /// <summary>The trailing named-argument names of a late-bound member, as one more argument to the LateBound call, or nothing.</summary>
    private static string Names(BoundMember member) =>
        member.NamedArguments.Count == 0 ? string.Empty : ", [" + string.Join(", ", member.NamedArguments.Select(Quote)) + "]";

    /// <summary>A call to a project procedure; <paramref name="trailingArgument"/> is the value of a Property Let or Set.</summary>
    private string CallCode(BoundCall call, string? trailingArgument)
    {
        var arguments = new List<string>();
        foreach (var argument in call.Arguments)
        {
            arguments.Add(Argument(argument, addressOnly: call.Procedure.IsRuntimeVarPtr));
        }

        if (call.Procedure.ParamArray is not null)
        {
            arguments.Add($"{R}VbaArray.FromValues({R}VarType.Variant, {Arguments(call.ParamArrayArguments)}, 0)");
        }

        if (trailingArgument is not null)
        {
            arguments.Add(trailingArgument);
        }

        if (call.Procedure.ReturnType is { } returned && Counted(returned))
        {
            // The callee hands its result's reference to this statement (ARCHITECTURE.md D18).
            OpenScope();
        }

        var callee = call.Receiver is { } receiver
            ? $"{Instance(EmitExpression(receiver), receiver.Type)}.{call.Procedure.EmitName}"
            : ReferenceEquals(call.Procedure.Module, Symbol)
                ? call.Procedure.EmitName
                : $"global::{call.Procedure.Module.EmitName}.{call.Procedure.EmitName}";
        return $"{callee}({string.Join(", ", arguments)})";
    }

    /// <summary>One argument as the callee receives it (MS-VBAL 5.6.13.1): a value, a ref, or a temporary copied back after the call.</summary>
    private string Argument(BoundArgument argument, bool addressOnly)
    {
        var parameter = argument.Parameter;
        if (ReferenceEquals(parameter.Type, VbaType.Any))
        {
            return AnyArgument(argument);
        }

        if (argument.Value is null)
        {
            var value = parameter.Default is { } constant
                ? Convert(Literal(constant), parameter.Type)
                : parameter.Type.IsVariant ? VariantType + ".Missing" : DefaultValue(parameter);
            if (!parameter.EmitByRef)
            {
                return value;
            }

            var omitted = NewTemporary();
            before.Add($"var {omitted} = {RefValue(value, parameter)};");
            ReleaseAfter(omitted, parameter.Type, IsObjectSlot(parameter));
            return "ref " + omitted;
        }

        if (parameter.IsExternal && !parameter.IsByVal && ReferenceEquals(parameter.Type, VbaType.Boolean)
            && !(argument.Mode == ArgumentMode.ByRef && argument.Value is BoundElement { Array.Type.IsArray: true }))
        {
            // A Boolean crosses to a DLL as the two bytes VBA passes (Declares golden): a temporary of them, stored back into a variable after the call.
            var flag = NewTemporary();
            before.Add($"short {flag} = {Convert(EmitExpression(argument.Value), VbaType.Boolean)} ? (short)-1 : (short)0;");
            if (argument.Mode != ArgumentMode.ByVal && argument.Value.IsLValue)
            {
                after.Add(Store(argument.Value, new Emitted($"{flag} != 0", VbaType.Boolean)) + ";");
            }

            return "ref " + flag;
        }

        switch (argument.Mode)
        {
            case ArgumentMode.ByVal:
                {
                    var value = parameter.Type.IsVariant ? LetVariant(EmitExpression(argument.Value)) : Convert(EmitExpression(argument.Value), parameter.Type);
                    if (!parameter.EmitByRef)
                    {
                        return value;
                    }

                    if (parameter.IsExternal && parameter.Type.IsRecord)
                    {
                        // A record that is not a variable (a function's result) reaches a Declare as a temporary, as VBA passes it (MS-VBAL 5.6.13.1; Declares golden).
                        var record = NewTemporary();
                        before.Add($"var {record} = {RefValue(value, parameter)};");
                        ReleaseAfter(record, parameter.Type);
                        return DeclareRecord(record, parameter.Type, argument.Value.Syntax);
                    }

                    // A value for a ByRef parameter (parenthesized, a literal, an expression): a temporary the callee may change freely (MS-VBAL 5.6.13.1).
                    var copy = NewTemporary();
                    before.Add($"var {copy} = {RefValue(value, parameter)};");
                    ReleaseAfter(copy, parameter.Type, IsObjectSlot(parameter));
                    return "ref " + copy;
                }
            case ArgumentMode.ByRef:
                if (IsMe(argument.Value))
                {
                    // Me is not assignable: it reaches a ByRef parameter as a temporary, which nothing is stored back from (Classes golden).
                    return CopyBack(argument.Value, parameter);
                }

                if (parameter.IsExternal && parameter.Type.IsRecord)
                {
                    // A record handed to a Declare: its block, or a copy whose Strings are ANSI (DeclareRecord; Declares golden).
                    return DeclareRecord(argument.Value);
                }

                if (argument.Value is BoundVariable { Variable.IsNew: true } created)
                {
                    // As New creates the object before the callee sees the variable (MS-VBAL 5.2.3.1.1).
                    before.Add($"_ = {Variable(created.Variable, forStore: false).Code};");
                    if (!IsObjectSlot(parameter))
                    {
                        return "ref " + Variable(created.Variable, forStore: true).Code;
                    }
                }

                if (argument.Value is BoundElement element && parameter.IsExternal && (element.Array.Type.IsArray || element.Array.Type.IsVariant))
                {
                    // An array element handed to a Declare, a Variant's array's included: its place in the array's native storage (M7 C4), as VBA passes the element's address (Declares golden).
                    return "ref " + ElementReference(element, ExternType(parameter.Type));
                }

                if (IsObjectSlot(parameter))
                {
                    // A typed object variable travels as its slot, so a Set in the callee reaches it (ROADMAP.md M7 E3); storage that is not a slot takes the copy-back path.
                    return SlotReference(argument.Value, parameter.Type) is { } slot ? "ref " + slot : CopyBack(argument.Value, parameter);
                }

                if (addressOnly && argument.Value is BoundField { Field: var array } arrayMember && IsDynamicArrayMember(array))
                {
                    // VBE7's VarPtr takes only the address: a dynamic array member's own, where its descriptor pointer lies (ROADMAP.md M7 C8; Memory golden).
                    return $"ref global::System.Runtime.CompilerServices.Unsafe.As<{R}VbaArrayMember, {R}VbaArray>(ref {FieldOwner(arrayMember)}.{array.EmitName})";
                }

                if (argument.Value is BoundField { Field: var member } && IsBooleanMember(member))
                {
                    // A Boolean member is VBA's two bytes, not the callee's bool: a temporary, stored back after the call.
                    return CopyBack(argument.Value, parameter);
                }

                return "ref " + ArrayReference(argument.Value);
            default:
                return CopyBack(argument.Value, parameter);
        }
    }

    /// <summary>A variable for a ByRef parameter it cannot be handed as itself (MS-VBAL 5.6.13.1): a temporary of the parameter's type, stored back into the variable after the call.</summary>
    private string CopyBack(BoundExpression variable, ParameterSymbol parameter)
    {
        // Declared as the parameter's type, since an object of another type converts to it without changing its C# type (a class instance to As Object).
        var temporary = NewTemporary();
        var value = Convert(EmitExpression(variable), parameter.Type);
        var slot = IsObjectSlot(parameter);
        before.Add(slot ? $"var {temporary} = {RefValue(value, parameter)};" : $"{parameter.Type.CSharpName} {temporary} = {Own(value, parameter.Type)};");
        if (variable is BoundElement { Array.Type.IsArray: false } late)
        {
            // value(indices) on a Variant or an object: an array element takes the change back, an object's default member gave a temporary (Procedures golden).
            after.Add($"{R}LateBound.StoreBack({Convert(EmitExpression(late.Array), VbaType.Variant)}, {Arguments(late.Indices)}, {LetVariant(new Emitted(slot ? temporary + ".Target" : temporary, parameter.Type))});");
        }
        else if (variable is BoundVariable or BoundElement or BoundField && !IsMe(variable))
        {
            // Storage takes the callee's change back; a property or a call's result was only ever a temporary, as VBA passes it (MS-VBAL 5.6.13.1; Classes golden).
            after.Add(Store(variable, new Emitted(slot ? temporary + ".Target" : temporary, parameter.Type)) + ";");
        }

        ReleaseAfter(temporary, parameter.Type, slot);
        return "ref " + temporary;
    }

    /// <summary>Whether an expression is the module's Me, which no assignment reaches.</summary>
    private bool IsMe(BoundExpression expression) => expression is BoundVariable { Variable: var variable } && ReferenceEquals(variable, Symbol.Me);

    private string NewTemporary() => "__t" + (++temporaries).ToString(CultureInfo.InvariantCulture);

    private Emitted Intrinsic(BoundIntrinsicCall call)
    {
        var intrinsic = call.Intrinsic;
        var arguments = call.Arguments;
        var template = intrinsic.TemplateFor(arguments.Count);
        var code = Placeholder.Replace(template, match =>
        {
            var name = match.Groups[1].Value;
            var format = match.Groups[2].Success ? match.Groups[2].Value : string.Empty;
            var fallback = match.Groups[3].Success ? match.Groups[3].Value : null;
            switch (name)
            {
                case "cmp":
                    return CompareModeName;
                case "base":
                    return Symbol.Options.Base.ToString(CultureInfo.InvariantCulture);
                case "rnd":
                    return L + "RandomGenerator.Shared";
                case "rest":
                    {
                        var start = int.Parse(format, CultureInfo.InvariantCulture);
                        var rest = arguments.Skip(start).Select(a => a is null ? VariantType + ".Missing" : Convert(EmitExpression(a), VbaType.Variant));
                        return "[" + string.Join(", ", rest) + "]";
                    }

                default:
                    {
                        var index = int.Parse(name, CultureInfo.InvariantCulture);
                        var argument = index < arguments.Count ? arguments[index] : null;
                        var target = format switch
                        {
                            "d" => VbaType.Double,
                            "i" => VbaType.Long,
                            _ => VbaType.Variant,
                        };
                        if (argument is not null)
                        {
                            return Convert(EmitExpression(argument), target);
                        }

                        return fallback ?? (target.IsVariant ? VariantType + ".Missing" : "0");
                    }
            }
        });

        var emitted = new Emitted(code, intrinsic.ResultType);
        if (call.DollarForm)
        {
            return new Emitted($"{L}Strings.DollarForm({Convert(emitted, VbaType.Variant)})", VbaType.String);
        }

        return emitted;
    }

    private Emitted Unary(BoundUnary unary) => Floating(unary, () => UnaryCore(unary));

    private Emitted UnaryCore(BoundUnary unary)
    {
        var operand = EmitExpression(unary.Operand);
        return unary.Kind switch
        {
            UnaryKind.Negate => new Emitted($"{R}Operators.Negate({Convert(operand, VbaType.Variant)}, {(unary.Operand.Type.IsVariant ? "true" : "false")})", VbaType.Variant),
            UnaryKind.Not => new Emitted($"{R}Operators.Not({Convert(operand, VbaType.Variant)})", VbaType.Variant),
            _ => operand,
        };
    }

    /// <summary>
    /// A typed Double or Single operation: its division by zero or overflow stores the IEEE result
    /// first and raises afterwards (docs/vba-quirks.md, "Runtime, floating point"), which the
    /// emitter arranges by checking after the store of a Double or Single variable and before any
    /// other use of the outermost such operation.
    /// </summary>
    private static bool IsFloating(BoundExpression expression) => expression switch
    {
        BoundParenthesized parenthesized => IsFloating(parenthesized.Operand),
        BoundBinary { Kind: BinaryKind.Add or BinaryKind.Subtract or BinaryKind.Multiply or BinaryKind.Divide or BinaryKind.Power } binary => IsFloatingType(binary.Type),
        BoundUnary { Kind: UnaryKind.Negate } unary => IsFloatingType(unary.Type),
        _ => false,
    };

    private static bool IsFloatingType(VbaType type) => ReferenceEquals(type, VbaType.Double) || ReferenceEquals(type, VbaType.Single);

    /// <summary>The floating-point check around the outermost typed operation, unless the statement stores it into a Double or Single first (<see cref="deferFloatingCheck"/>).</summary>
    private Emitted Floating(BoundExpression expression, Func<Emitted> emit)
    {
        if (!IsFloating(expression))
        {
            return emit();
        }

        var outermost = floatingDepth == 0;
        floatingDepth++;
        Emitted result;
        try
        {
            result = emit();
        }
        finally
        {
            floatingDepth--;
        }

        if (!outermost)
        {
            return result;
        }

        if (deferFloatingCheck)
        {
            deferFloatingCheck = false;
            floatingCheckDeferred = true;
            return result;
        }

        return result.Type.Kind == TypeKind.Builtin && result.Type.VarType == VarType.Double
            ? new Emitted($"{R}Operators.CheckFloating({result.Code})", VbaType.Double)
            : new Emitted($"{R}Operators.CheckFloating({Convert(result, VbaType.Variant)})", VbaType.Variant);
    }

    private Emitted Binary(BoundBinary binary) => Floating(binary, () => BinaryCore(binary));

    private Emitted BinaryCore(BoundBinary binary)
    {
        if (TypedOperation(binary) is { } typed)
        {
            return typed;
        }

        var left = Convert(EmitExpression(binary.Left), VbaType.Variant);
        var right = Convert(EmitExpression(binary.Right), VbaType.Variant);
        var declared = DeclaredFlags(binary.Left.Type.IsVariant, binary.Right.Type.IsVariant, IsFloating(binary));
        var code = binary.Kind switch
        {
            BinaryKind.Add => $"{R}Operators.Add({left}, {right}, {declared})",
            BinaryKind.Subtract => $"{R}Operators.Subtract({left}, {right}, {declared})",
            BinaryKind.Multiply => $"{R}Operators.Multiply({left}, {right}, {declared})",
            BinaryKind.Divide => $"{R}Operators.Divide({left}, {right}, {declared})",
            BinaryKind.IntegerDivide => $"{R}Operators.IntegerDivide({left}, {right}, {declared})",
            BinaryKind.Modulo => $"{R}Operators.Modulo({left}, {right}, {declared})",
            BinaryKind.Power => $"{R}Operators.Power({left}, {right}, {declared})",
            BinaryKind.Concatenate => $"{R}Operators.Concatenate({left}, {right})",
            BinaryKind.Equal => $"{R}Operators.Equal({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.NotEqual => $"{R}Operators.NotEqual({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.LessThan => $"{R}Operators.LessThan({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.GreaterThan => $"{R}Operators.GreaterThan({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.LessThanOrEqual => $"{R}Operators.LessThanOrEqual({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.GreaterThanOrEqual => $"{R}Operators.GreaterThanOrEqual({left}, {right}, {declared}, {CompareModeName})",
            BinaryKind.Like => $"{R}Operators.Like({left}, {right}, {CompareModeName})",
            BinaryKind.Is => $"{R}Operators.Is({left}, {right})",
            BinaryKind.And => $"{R}Operators.And({left}, {right})",
            BinaryKind.Or => $"{R}Operators.Or({left}, {right})",
            BinaryKind.Xor => $"{R}Operators.Xor({left}, {right})",
            BinaryKind.Eqv => $"{R}Operators.Eqv({left}, {right})",
            _ => $"{R}Operators.Imp({left}, {right})",
        };
        return new Emitted(code, VbaType.Variant);
    }

    /// <summary>
    /// Typed operands through C# (ROADMAP.md D-J, M7 E): +, -, and * whose operands both have
    /// declared numeric types, in the result type MS-VBAL 5.6.9.3 gives them when it is Byte,
    /// Integer, Long, LongLong, or Double, / when that type is Double (D-M: Double arithmetic
    /// rather than VBA's x87 extended precision), and the comparisons of two integers, which are exact
    /// whatever their declared types (MS-VBAL 5.6.9.5). The runtime helpers raise 6 when an
    /// integral result does not fit and leave a Double's overflow pending for the statement's
    /// check, as the Variant operators do for typed operands (Operators and Errors goldens).
    /// </summary>
    private Emitted? TypedOperation(BoundBinary binary)
    {
        if (!IsTypedNumber(binary.Left.Type) || !IsTypedNumber(binary.Right.Type))
        {
            return null;
        }

        if (ComparisonOperator(binary.Kind) is { } comparison)
        {
            // Each operand as its declared type: an element of a typed array reads as a Variant (Arrays golden).
            return IsTypedInteger(binary.Left.Type) && IsTypedInteger(binary.Right.Type)
                ? new Emitted($"(({Convert(EmitExpression(binary.Left), binary.Left.Type)}) {comparison} ({Convert(EmitExpression(binary.Right), binary.Right.Type)}))", VbaType.Boolean)
                : null;
        }

        if (binary.Kind is not (BinaryKind.Add or BinaryKind.Subtract or BinaryKind.Multiply or BinaryKind.Divide))
        {
            return null;
        }

        var suffix = binary.Type.Kind != TypeKind.Builtin ? null : binary.Type.VarType switch
        {
            VarType.Byte => "Byte",
            VarType.Integer => "Int16",
            VarType.Long => "Int32",
            VarType.LongLong => "Int64",
            VarType.Double => "Double",
            _ => null,
        };

        // / only where the quotient is Double. A Single divided by a Single, Byte, or Integer, either way round, is a
        // Single at run time whatever the declared type says (MS-VBAL 5.6.9.3.4, Operators golden), so it keeps the Variant path.
        if (suffix is null || (binary.Kind == BinaryKind.Divide && (suffix != "Double" || IsSingleQuotient(binary.Left.Type, binary.Right.Type))))
        {
            return null;
        }

        var left = Convert(EmitExpression(binary.Left), binary.Type);
        var right = Convert(EmitExpression(binary.Right), binary.Type);
        return new Emitted($"{R}Operators.{binary.Kind}{suffix}({left}, {right})", binary.Type);
    }

    private static string? ComparisonOperator(BinaryKind kind) => kind switch
    {
        BinaryKind.Equal => "==",
        BinaryKind.NotEqual => "!=",
        BinaryKind.LessThan => "<",
        BinaryKind.GreaterThan => ">",
        BinaryKind.LessThanOrEqual => "<=",
        BinaryKind.GreaterThanOrEqual => ">=",
        _ => null,
    };

    private static bool IsSingleQuotient(VbaType left, VbaType right)
    {
        static bool Is(VbaType type, VarType varType) => type.Kind == TypeKind.Builtin && type.VarType == varType;
        static bool Narrow(VbaType type) => Is(type, VarType.Single) || Is(type, VarType.Byte) || Is(type, VarType.Integer);
        return (Is(left, VarType.Single) && Narrow(right)) || (Is(right, VarType.Single) && Narrow(left));
    }

    private static bool IsTypedInteger(VbaType type) =>
        type.Kind == TypeKind.Enum || (type.Kind == TypeKind.Builtin && type.VarType is VarType.Byte or VarType.Integer or VarType.Long or VarType.LongLong);

    private static bool IsTypedNumber(VbaType type) =>
        type.Kind == TypeKind.Enum || (type.Kind == TypeKind.Builtin && type.VarType is VarType.Byte or VarType.Integer or VarType.Long or VarType.LongLong or VarType.Single or VarType.Double);

    /// <summary>A conversion between typed numbers that loses nothing: a narrower integer to a wider one or to a floating type that holds it exactly, and Single to Double.</summary>
    private static bool IsWidening(VbaType from, VbaType to)
    {
        if (from.Kind != TypeKind.Builtin || to.Kind != TypeKind.Builtin)
        {
            return false;
        }

        return (from.VarType, to.VarType) switch
        {
            (VarType.Byte, VarType.Integer or VarType.Long or VarType.LongLong or VarType.Single or VarType.Double) => true,
            (VarType.Integer, VarType.Long or VarType.LongLong or VarType.Single or VarType.Double) => true,
            (VarType.Long, VarType.LongLong or VarType.Double) => true,
            (VarType.Single, VarType.Double) => true,
            _ => false,
        };
    }

    private static string DeclaredFlags(bool leftVariant, bool rightVariant, bool floating = false)
    {
        var flags = R + "DeclaredTypes." + (leftVariant, rightVariant) switch
        {
            (true, true) => "BothVariant",
            (true, false) => "LeftVariant",
            (false, true) => "RightVariant",
            _ => "None",
        };
        return floating ? flags + " | " + R + "DeclaredTypes.Floating" : flags;
    }

    /// <summary>Let-coercion between C# representations (MS-VBAL 5.5.1): through the Variant operators unless the C# types already agree.</summary>
    private static string Convert(Emitted value, VbaType to)
    {
        var from = value.Type;
        var code = value.Code;
        if (to.IsVariant)
        {
            return ToVariant(value);
        }

        if (from.IsVariant)
        {
            return FromVariant(code, to);
        }

        if (to.IsArray)
        {
            if (from.IsArray)
            {
                // A view: whatever keeps the array takes a copy of its own (Assignment, Own), so a conversion never copies (M7 C4).
                return code;
            }

            if (from.IsString)
            {
                return $"{R}VbaArray.FromString({code})";
            }

            return FromVariant(ToVariant(value), to);
        }

        if (from.IsArray)
        {
            // A Byte array assigned to a String keeps every byte (D20); a fixed-length string is still a .NET string.
            if (to.IsVariableString)
            {
                return code + ".ToText()";
            }

            var text = code + ".ToStringValue()";
            return to.Kind == TypeKind.FixedString ? Fixed(text, to) : FromVariant($"{VariantType}.FromString({text})", to);
        }

        if (to.IsRecord)
        {
            return from.IsRecord ? $"{code}.Copy()" : code;
        }

        if (to.Kind == TypeKind.FixedString)
        {
            return Fixed(from.IsString ? code : FromVariant(ToVariant(value), VbaType.String), to);
        }

        if (to.IsVariableString && from.Kind == TypeKind.FixedString)
        {
            return $"{R}VbaString.Temporary({code})";
        }

        if (from.CSharpName == to.CSharpName)
        {
            return code;
        }

        if (IsWidening(from, to))
        {
            // Widening between typed numbers is exact, so the C# conversion is VBA's (ROADMAP.md D-J).
            return $"(({to.CSharpName}){code})";
        }

        if (to.IsObject && from.IsObject)
        {
            return to.IsGenericObject ? code : $"{R}Coerce.ToObject<{ClassName(to)}>({VariantType}.FromObject({code}))";
        }

        return FromVariant(ToVariant(value), to);
    }

    private static string Fixed(string text, VbaType fixedString) =>
        $"{L}Strings.ToFixed({text}, {fixedString.FixedLength.ToString(CultureInfo.InvariantCulture)})";

    private static string ToVariant(Emitted value)
    {
        var (code, type) = value;
        if (type.IsVariant)
        {
            return code;
        }

        if (type.IsArray)
        {
            return $"{VariantType}.FromArray({code})";
        }

        if (type.IsRecord)
        {
            // The binder rejects a user-defined type where a Variant is wanted (MS-VBAL 5.5.1.2.5), and the compiler stops before emitting (ROADMAP.md M7 C3).
            throw new InvalidOperationException("A user-defined type never converts to a Variant.");
        }

        if (type.IsObject)
        {
            return $"{VariantType}.FromObject({code})";
        }

        if (type.IsVariableString)
        {
            // A String variable reaches an operator or a Variant parameter as a view of its BSTR; a store copies (D20).
            return $"{VariantType}.ViewString({code})";
        }

        var factory = type.VarType switch
        {
            VarType.Integer => "FromInt16",
            VarType.Long => "FromInt32",
            VarType.LongLong => "FromInt64",
            VarType.Single => "FromSingle",
            VarType.Double => "FromDouble",
            VarType.Currency => "FromCurrency",
            VarType.Date => "FromDate",
            VarType.String => "FromString",
            VarType.Boolean => "FromBoolean",
            VarType.Byte => "FromByte",
            VarType.Decimal => "FromDecimal",
            _ => "FromObject",
        };
        return $"{VariantType}.{factory}({code})";
    }

    private static string FromVariant(string code, VbaType to)
    {
        if (to.IsArray)
        {
            return $"{R}Coerce.ToArray({code}, {VarTypeName(to.ElementType!)})";
        }

        if (to.IsRecord)
        {
            throw new InvalidOperationException("A Variant never converts to a user-defined type.");
        }

        if (to.IsObject)
        {
            return $"{R}Coerce.ToObject<{ClassName(to)}>({code})";
        }

        if (to.Kind == TypeKind.FixedString)
        {
            return Fixed($"{R}Coerce.ToString({code})", to);
        }

        if (to.IsVariableString)
        {
            return $"{R}Coerce.ToText({code})";
        }

        var method = to.VarType switch
        {
            VarType.Integer => "ToInt16",
            VarType.Long => "ToInt32",
            VarType.LongLong => "ToInt64",
            VarType.Single => "ToSingle",
            VarType.Double => "ToDouble",
            VarType.Currency => "ToCurrency",
            VarType.Date => "ToDate",
            VarType.String => "ToString",
            VarType.Boolean => "ToBoolean",
            VarType.Byte => "ToByte",
            VarType.Decimal => "ToDecimal",
            _ => "ToObject<object>",
        };
        return $"{R}Coerce.{method}({code})";
    }

    /// <summary>A value stored into a Variant by Let (MS-VBAL 5.5.1.2.5): a view the store copies, an array included (ObjectRefs.Own); other values converted.</summary>
    private static string LetVariant(Emitted value) => Convert(value, VbaType.Variant);

    /// <summary>The C# class behind an object type, without the nullable annotation.</summary>
    private static string ClassName(VbaType type) => type.IsGenericObject ? "object" : type.CSharpName.TrimEnd('?');

    /// <summary>A compile-time value as a C# expression of its own type (deterministic text, CLAUDE.md R10).</summary>
    private Emitted Literal(Variant value)
    {
        switch (value.Type)
        {
            case VarType.Integer:
                return new Emitted($"((short){value.AsInt16().ToString(CultureInfo.InvariantCulture)})", VbaType.Integer);
            case VarType.Long:
                return new Emitted(value.AsInt32().ToString(CultureInfo.InvariantCulture), VbaType.Long);
            case VarType.LongLong:
                return new Emitted(value.AsInt64().ToString(CultureInfo.InvariantCulture) + "L", VbaType.LongLong);
            case VarType.Single:
                return new Emitted(SingleLiteral(value.AsSingle()), VbaType.Single);
            case VarType.Double:
                return new Emitted(DoubleLiteral(value.AsDouble()), VbaType.Double);
            case VarType.Currency:
                return new Emitted($"{R}Currency.FromScaled({value.AsCurrency().Scaled.ToString(CultureInfo.InvariantCulture)}L)", VbaType.Currency);
            case VarType.Date:
                return new Emitted($"{R}VbaDate.FromSerial({DoubleLiteral(value.AsDate().Serial)})", VbaType.Date);
            case VarType.String:
                return new Emitted(StringLiteral(value.AsString()), VbaType.String);
            case VarType.Boolean:
                return new Emitted(value.AsBoolean() ? "true" : "false", VbaType.Boolean);
            case VarType.Byte:
                return new Emitted($"((byte){value.AsByte().ToString(CultureInfo.InvariantCulture)})", VbaType.Byte);
            case VarType.Decimal:
                return new Emitted(value.AsDecimal().ToString(CultureInfo.InvariantCulture) + "m", VbaType.Decimal);
            case VarType.Null:
                return new Emitted(VariantType + ".Null", VbaType.Variant);
            case VarType.Error:
                return value.IsMissing
                    ? new Emitted(VariantType + ".Missing", VbaType.Variant)
                    : new Emitted($"{VariantType}.FromError(new {R}ErrorValue({value.AsError().Scode.ToString(CultureInfo.InvariantCulture)}))", VbaType.Variant);
            case VarType.Object:
                return new Emitted("((object?)null)", VbaType.Object);
            default:
                return new Emitted(VariantType + ".Empty", VbaType.Variant);
        }
    }

    /// <summary>
    /// A string literal is a BSTR the module allocates once and keeps, as VBA keeps its literals in
    /// the module's data (D20): every use views it, a store copies it, so evaluating a literal
    /// allocates nothing (R20). Literals are numbered in order of first use, deterministically (R10).
    /// </summary>
    private string StringLiteral(string text)
    {
        if (!literalIndex.TryGetValue(text, out var index))
        {
            index = literals.Count;
            literalIndex.Add(text, index);
            literals.Add(text);
        }

        return "__str_" + index.ToString(CultureInfo.InvariantCulture);
    }

    private void EmitLiterals()
    {
        for (var i = 0; i < literals.Count; i++)
        {
            writer.HiddenLine($"private static readonly {R}VbaString __str_{i.ToString(CultureInfo.InvariantCulture)} = {R}VbaString.Literal({Quote(literals[i])});");
        }
    }

    private static string DoubleLiteral(double value)
    {
        if (double.IsNaN(value))
        {
            return "double.NaN";
        }

        if (double.IsPositiveInfinity(value))
        {
            return "double.PositiveInfinity";
        }

        if (double.IsNegativeInfinity(value))
        {
            return "double.NegativeInfinity";
        }

        return value.ToString("R", CultureInfo.InvariantCulture) + "d";
    }

    private static string SingleLiteral(float value)
    {
        if (float.IsNaN(value))
        {
            return "float.NaN";
        }

        if (float.IsPositiveInfinity(value))
        {
            return "float.PositiveInfinity";
        }

        if (float.IsNegativeInfinity(value))
        {
            return "float.NegativeInfinity";
        }

        return value.ToString("R", CultureInfo.InvariantCulture) + "f";
    }
}
