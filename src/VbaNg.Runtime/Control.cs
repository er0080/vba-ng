using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>Helpers for the control-flow statements generated code compiles to.</summary>
public static class ForLoop
{
    /// <summary>For ... Next (MS-VBAL 5.4.2.3): the loop continues while counter is not past the limit in the step's direction; a Null comparison ends it.</summary>
    public static bool Continues(in Variant counter, in Variant limit, in Variant step)
    {
        var ascending = Coerce.ToDouble(step) >= 0;
        var test = ascending
            ? Operators.LessThanOrEqual(counter, limit, DeclaredTypes.BothVariant)
            : Operators.GreaterThanOrEqual(counter, limit, DeclaredTypes.BothVariant);
        return Coerce.ToCondition(test);
    }

    // A counter of a declared numeric type steps natively (ROADMAP.md D-J, M7 E): the start, limit, and step were
    // Let-coerced to its type before the loop (MS-VBAL 5.4.2.3), so the test compares in that type and the step adds
    // in it, raising 6 when the sum does not fit (ControlFlow golden: an Integer counter to 32767).

    public static bool Continues(byte counter, byte limit, byte step) => step >= 0 ? counter <= limit : counter >= limit;

    public static bool Continues(short counter, short limit, short step) => step >= 0 ? counter <= limit : counter >= limit;

    public static bool Continues(int counter, int limit, int step) => step >= 0 ? counter <= limit : counter >= limit;

    public static bool Continues(long counter, long limit, long step) => step >= 0 ? counter <= limit : counter >= limit;

    public static bool Continues(double counter, double limit, double step) => step >= 0 ? counter <= limit : counter >= limit;

    public static byte Next(byte counter, byte step) => Operators.AddByte(counter, step);

    public static short Next(short counter, short step) => Operators.AddInt16(counter, step);

    public static int Next(int counter, int step) => Operators.AddInt32(counter, step);

    public static long Next(long counter, long step) => Operators.AddInt64(counter, step);

    public static double Next(double counter, double step)
    {
        var sum = counter + step;
        return double.IsInfinity(sum) && double.IsFinite(counter) && double.IsFinite(step) ? throw VbaErrors.Overflow() : sum;
    }
}

/// <summary>
/// The enumerator behind For Each (MS-VBAL 5.4.2.4): an array is walked in storage order,
/// reading each element when its turn comes, so an element the body changes before the loop
/// reaches it is seen changed (ControlFlow golden); a Collection is walked live. An unallocated
/// array raises 92, and anything else raises 13.
/// </summary>
public sealed class ForEachEnumerator
{
    private readonly IEnumerator<Variant> items;
    private object? held;

    // An array the loop walks, locked for the loop's length, and whether the loop owns it (a Split, a function's result).
    private VbaArray array;
    private bool ownsArray;

    private ForEachEnumerator(IEnumerator<Variant> items, bool isArray, object? held)
    {
        this.items = items;
        this.held = held;
        IsArray = isArray;
    }

    /// <summary>The source was an array: after the last element a Variant loop variable becomes Empty (Arrays golden).</summary>
    public bool IsArray { get; }

    public Variant Current => items.Current;

    /// <summary>
    /// The enumerator over the source. An array the statement owns (a Split, a function's
    /// result) lives for the loop instead of dying with the header's statement: the enumerator
    /// takes it over and destroys it when the loop ends (D18, D20); every array is locked while the
    /// loop walks it, so a ReDim or an Erase of it raises error 10. An object is held
    /// by a reference of the enumerator's own for the same span.
    /// </summary>
    public static ForEachEnumerator Create(in Variant source) => Create(source, 0);

    /// <summary>
    /// The enumerator of a For Each header in a procedure whose temporaries start at <paramref name="statementMark"/>
    /// (the procedure's mark): only an array its own header statement made is the loop's to take over, so a
    /// ParamArray, the caller's temporary, survives a loop over it (Procedures golden).
    /// </summary>
    public static ForEachEnumerator Create(in Variant source, int statementMark)
    {
        if (source.IsArray)
        {
            var array = source.AsArray();
            if (!array.IsAllocated)
            {
                throw new VbaException(VbaErrors.ForLoopNotInitialized);
            }

            var owned = ObjectRefs.Detach(array, statementMark);
            array.Lock();
            return new ForEachEnumerator(Live(array), isArray: true, held: null) { array = array, ownsArray = owned };
        }

        if (source.IsObject)
        {
            var value = source.AsObject();
            (value as IReferenceCounted)?.AddRef();
            switch (value)
            {
                case IDispatchObject dispatch:
                    return new ForEachEnumerator(dispatch.Enumerate(), isArray: false, value);
                case IEnumerable<Variant> enumerable:
                    return new ForEachEnumerator(enumerable.GetEnumerator(), isArray: false, value);
                case null:
                    throw VbaErrors.ObjectVariableNotSet();
            }
        }

        throw VbaErrors.TypeMismatch();
    }

    public bool MoveNext()
    {
        if (items.MoveNext())
        {
            return true;
        }

        Finish();
        return false;
    }

    /// <summary>The loop is over, by its last element, an Exit For, or the procedure's end: what the enumerator held for it goes.</summary>
    internal void Finish()
    {
        var walked = array;
        var owned = ownsArray;
        array = default;
        ownsArray = false;
        if (walked.IsAllocated)
        {
            walked.Unlock();
            if (owned)
            {
                walked.Destroy();
            }
        }

        var kept = held;
        held = null;
        (kept as IReferenceCounted)?.Release();
    }

    private static IEnumerator<Variant> Live(VbaArray array)
    {
        for (var i = 0; i < array.Count; i++)
        {
            yield return array.ElementAt(i);
        }
    }
}

public enum ComparisonOperator
{
    Equal,
    NotEqual,
    LessThan,
    GreaterThan,
    LessThanOrEqual,
    GreaterThanOrEqual,
}

/// <summary>
/// Select Case clause tests (MS-VBAL 5.4.2.10), evaluated left to right until one matches. The
/// comparison follows the declared types of the selector and the clause value, as the
/// operators do: a String selector against a numeric clause compares as numbers, a Variant
/// number against a String literal too (ControlFlow golden). A comparison that yields Null
/// raises 94, unlike a Null selector, whose clauses simply never match.
/// </summary>
public static class SelectCase
{
    public static bool Matches(in Variant value, in Variant item, DeclaredTypes declared, CompareMode mode) =>
        Decide(value, Operators.Equal(value, item, declared, mode));

    public static bool InRange(in Variant value, in Variant low, in Variant high, DeclaredTypes declaredLow, DeclaredTypes declaredHigh, CompareMode mode) =>
        Decide(value, Operators.GreaterThanOrEqual(value, low, declaredLow, mode))
        && Decide(value, Operators.LessThanOrEqual(value, high, declaredHigh, mode));

    public static bool Compares(in Variant value, ComparisonOperator comparison, in Variant item, DeclaredTypes declared, CompareMode mode)
    {
        var result = comparison switch
        {
            ComparisonOperator.Equal => Operators.Equal(value, item, declared, mode),
            ComparisonOperator.NotEqual => Operators.NotEqual(value, item, declared, mode),
            ComparisonOperator.LessThan => Operators.LessThan(value, item, declared, mode),
            ComparisonOperator.GreaterThan => Operators.GreaterThan(value, item, declared, mode),
            ComparisonOperator.LessThanOrEqual => Operators.LessThanOrEqual(value, item, declared, mode),
            _ => Operators.GreaterThanOrEqual(value, item, declared, mode),
        };
        return Decide(value, result);
    }

    /// <summary>Case Null against a non-Null selector is error 94; a Null selector matches nothing (ControlFlow golden).</summary>
    private static bool Decide(in Variant selector, in Variant result) =>
        result.IsNull ? (selector.IsNull ? false : throw VbaErrors.InvalidUseOfNull()) : Coerce.ToBoolean(result);
}

/// <summary>
/// Members of objects generated code cannot bind at compile time: a Variant or Object variable
/// holding a Collection or the Err object (MS-VBAL 5.6.12 late binding). COM objects join in M4
/// through the IDispatch path (ARCHITECTURE.md D7). Nothing raises 91, a non-object raises 424,
/// an unknown member raises 438.
/// </summary>
public static class LateBound
{
    /// <summary>DISPID_VALUE: the default member of a COM object.</summary>
    public const int DefaultMember = 0;

    public static Variant Get(in Variant target, string name, ReadOnlySpan<Variant> arguments)
    {
        ArgumentNullException.ThrowIfNull(name);
        switch (RequireObject(target))
        {
            case IDispatchObject dispatch:
                return dispatch.Invoke(dispatch.GetDispId(name), InvokeKind.PropertyGet | InvokeKind.Method, arguments);

            case Collection collection:
                switch (name.ToUpperInvariant())
                {
                    case "COUNT":
                        return Variant.FromInt32(collection.Count);
                    case "ITEM":
                        return arguments.Length == 1 ? collection.Item(arguments[0]) : throw new VbaException(VbaErrors.WrongNumberOfArguments);
                    case "ADD":
                        collection.Add(Argument(arguments, 0), Argument(arguments, 1), Argument(arguments, 2), Argument(arguments, 3));
                        return Variant.Empty;
                    case "REMOVE":
                        collection.Remove(Argument(arguments, 0));
                        return Variant.Empty;
                    default:
                        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
                }

            case ErrObject err:
                switch (name.ToUpperInvariant())
                {
                    case "NUMBER":
                        return Variant.FromInt32(err.Number);
                    case "DESCRIPTION":
                        return Variant.FromString(err.DescriptionText);
                    case "SOURCE":
                        return Variant.FromString(err.SourceText);
                    case "HELPFILE":
                        return Variant.FromString(err.HelpFileText);
                    case "HELPCONTEXT":
                        return Variant.FromInt32(err.HelpContext);
                    case "LASTDLLERROR":
                        return Variant.FromInt32(err.LastDllError);
                    case "RAISE":
                        throw err.Raise(Argument(arguments, 0), Argument(arguments, 1), Argument(arguments, 2), Argument(arguments, 3), Argument(arguments, 4));
                    case "CLEAR":
                        err.Clear();
                        return Variant.Empty;
                    default:
                        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
                }

            default:
                throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }
    }

    /// <summary>
    /// obj.[name] (MS-VBAL 3.3.5.3): the member of that name when the object has one, else the object's
    /// Evaluate of the text, which is how ws.[A1] reaches a Range through a late-bound receiver.
    /// </summary>
    public static Variant Foreign(in Variant target, string name, ReadOnlySpan<Variant> arguments)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (RequireObject(target) is IDispatchObject dispatch)
        {
            int dispId;
            try
            {
                dispId = dispatch.GetDispId(name);
            }
            catch (VbaException ex) when (ex.Number == VbaErrors.ObjectDoesNotSupportMember)
            {
                return dispatch.Invoke(dispatch.GetDispId("Evaluate"), InvokeKind.PropertyGet | InvokeKind.Method, [Variant.FromString(name)]);
            }

            return dispatch.Invoke(dispId, InvokeKind.PropertyGet | InvokeKind.Method, arguments);
        }

        return Get(target, name, arguments);
    }

    public static void Call(in Variant target, string name, ReadOnlySpan<Variant> arguments) => Get(target, name, arguments);

    /// <summary>
    /// A late-bound call with named arguments (MS-VBAL 5.6.13.1): <paramref name="arguments"/> holds
    /// the positional ones and then the named ones, whose names <paramref name="names"/> gives in the
    /// same order. A COM object answers through GetIDsOfNames; the runtime's own Collection and Err
    /// take their documented parameter names; anything else is error 448.
    /// </summary>
    public static Variant Get(in Variant target, string name, ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (names.IsEmpty)
        {
            return Get(target, name, arguments);
        }

        switch (RequireObject(target))
        {
            case IDispatchObject dispatch:
                {
                    var dispIds = dispatch.GetDispIds(name, names);
                    return dispatch.InvokeNamed(dispIds[0], InvokeKind.PropertyGet | InvokeKind.Method, arguments, dispIds.AsSpan(1));
                }

            case Collection when name.Equals("Add", StringComparison.OrdinalIgnoreCase):
                return Get(target, name, Reorder(arguments, names, ["Item", "Key", "Before", "After"]));
            case ErrObject when name.Equals("Raise", StringComparison.OrdinalIgnoreCase):
                return Get(target, name, Reorder(arguments, names, ["Number", "Source", "Description", "HelpFile", "HelpContext"]));
            default:
                throw new VbaException(VbaErrors.NamedArgumentNotFound);
        }
    }

    public static void Call(in Variant target, string name, ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names) => Get(target, name, arguments, names);

    /// <summary>obj.Prop(Index:=i) = value, late-bound: the named index arguments resolve as a call's do.</summary>
    public static void Let(in Variant target, string name, ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names, in Variant value) =>
        PutNamed(target, name, arguments, names, value, asReference: false);

    public static void Set(in Variant target, string name, ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names, in Variant value) =>
        PutNamed(target, name, arguments, names, value, asReference: true);

    private static void PutNamed(in Variant target, string name, ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names, in Variant value, bool asReference)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (names.IsEmpty)
        {
            if (asReference)
            {
                Set(target, name, arguments, value);
            }
            else
            {
                Let(target, name, arguments, value);
            }

            return;
        }

        if (RequireObject(target) is IDispatchObject dispatch)
        {
            var dispIds = dispatch.GetDispIds(name, names);
            dispatch.PutNamed(dispIds[0], arguments, dispIds.AsSpan(1), value, asReference);
            return;
        }

        throw new VbaException(VbaErrors.NamedArgumentNotFound);
    }

    /// <summary>The positional arguments followed by the named ones, placed where their names say among <paramref name="parameters"/>; a name the member does not have is error 448.</summary>
    private static Variant[] Reorder(ReadOnlySpan<Variant> arguments, ReadOnlySpan<string> names, string[] parameters)
    {
        var positional = arguments.Length - names.Length;
        var ordered = new Variant[parameters.Length];
        var present = new bool[parameters.Length];
        for (var i = 0; i < ordered.Length; i++)
        {
            ordered[i] = Variant.Missing;
        }

        for (var i = 0; i < positional && i < ordered.Length; i++)
        {
            ordered[i] = arguments[i];
            present[i] = true;
        }

        for (var i = 0; i < names.Length; i++)
        {
            var wanted = names[i];
            var slot = Array.FindIndex(parameters, p => p.Equals(wanted, StringComparison.OrdinalIgnoreCase));
            if (slot < 0)
            {
                throw new VbaException(VbaErrors.NamedArgumentNotFound);
            }

            ordered[slot] = arguments[positional + i];
            present[slot] = true;
        }

        var last = Array.LastIndexOf(present, true);
        return ordered[..(last + 1)];
    }

    public static void Let(in Variant target, string name, ReadOnlySpan<Variant> arguments, in Variant value)
    {
        ArgumentNullException.ThrowIfNull(name);
        var obj = RequireObject(target);
        if (obj is IDispatchObject dispatch)
        {
            var dispId = dispatch.GetDispId(name);
            try
            {
                dispatch.Put(dispId, arguments, value, asReference: false);
            }
            catch (VbaException ex) when (arguments.Length > 0 && ex.Number == VbaErrors.ObjectDoesNotSupportMember)
            {
                // obj.Range("B1") = x: a property with arguments but no put accessor is read, and the
                // value goes to the default member of what it returns, as VBA does late-bound.
                var result = dispatch.Invoke(dispId, InvokeKind.PropertyGet | InvokeKind.Method, arguments);
                if (result.AsObject() is not IDispatchObject returned)
                {
                    throw;
                }

                returned.Put(DefaultMember, [], value, asReference: false);
            }

            return;
        }

        if (obj is ErrObject err && arguments.Length == 0)
        {
            switch (name.ToUpperInvariant())
            {
                case "NUMBER":
                    err.Number = Coerce.ToInt32(value);
                    return;
                case "DESCRIPTION":
                    err.DescriptionText = Coerce.ToString(value);
                    return;
                case "SOURCE":
                    err.SourceText = Coerce.ToString(value);
                    return;
                case "HELPFILE":
                    err.HelpFileText = Coerce.ToString(value);
                    return;
                case "HELPCONTEXT":
                    err.HelpContext = Coerce.ToInt32(value);
                    return;
            }
        }

        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
    }

    public static void Set(in Variant target, string name, ReadOnlySpan<Variant> arguments, in Variant value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (RequireObject(target) is IDispatchObject dispatch)
        {
            dispatch.Put(dispatch.GetDispId(name), arguments, value, asReference: true);
            return;
        }

        throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
    }

    /// <summary>
    /// A late-bound call passed a class member without parameters (a field, a Property Get, a Function) arguments: they go
    /// to the default member of the object the member returned (MS-VBAL 5.6.12; LateBinding golden). Nothing raises 438, and
    /// any other value 451, an array included, which the arguments do not index.
    /// </summary>
    public static Variant PassOn(in Variant value, ReadOnlySpan<Variant> arguments)
    {
        if (arguments.Length == 0)
        {
            return value;
        }

        if (value.IsNothing)
        {
            throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }

        return value.IsObject ? Index(value, arguments) : throw new VbaException(451);
    }

    /// <summary>value(indices) on a Variant: an array element, the default member of a COM object, or Item of a Collection.</summary>
    public static Variant Index(in Variant target, ReadOnlySpan<Variant> indices)
    {
        if (target.IsArray)
        {
            return target.AsArray().Get(Indices(indices));
        }

        if (target.IsObject)
        {
            return RequireObject(target) switch
            {
                IDispatchObject dispatch => dispatch.Invoke(DefaultMember, InvokeKind.PropertyGet | InvokeKind.Method, indices),
                Collection collection => indices.Length == 1 ? collection.Item(indices[0]) : throw new VbaException(VbaErrors.WrongNumberOfArguments),
                _ => throw new VbaException(VbaErrors.ObjectDoesNotSupportMember),
            };
        }

        throw VbaErrors.TypeMismatch();
    }

    /// <summary>value(indices) = x on a Variant: an array element takes x let-coerced, an object as its default member's value (MS-VBAL 5.6.9.3); a COM object's default member takes x as it is, an object included, and decides itself (Objects golden).</summary>
    public static void SetIndex(in Variant target, ReadOnlySpan<Variant> indices, in Variant value)
    {
        if (target.IsArray)
        {
            target.AsArray().Set(Indices(indices), Coerce.LetValue(value));
            return;
        }

        if (target.IsObject)
        {
            if (RequireObject(target) is IDispatchObject dispatch)
            {
                dispatch.Put(DefaultMember, indices, value, asReference: false);
                return;
            }

            throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }

        throw VbaErrors.TypeMismatch();
    }

    /// <summary>
    /// The copy-back of a ByRef argument written <c>value(indices)</c> on a Variant or an object
    /// (MS-VBAL 5.6.13.1): an element of an array is storage, which takes the callee's change as
    /// the callee left it, while an object's default member gave a call's result, a temporary
    /// nothing is stored back into (Procedures golden).
    /// </summary>
    public static void StoreBack(in Variant target, ReadOnlySpan<Variant> indices, in Variant value)
    {
        if (target.IsArray)
        {
            target.AsArray().Set(Indices(indices), value);
        }
    }

    /// <summary>Set value(indices) = object on a Variant: an array element, or the default member of a COM object, which takes the object itself (DISPATCH_PROPERTYPUTREF), as VBA's Set does (Objects golden: Set d("b") = New Collection on a Scripting.Dictionary).</summary>
    public static void SetIndexReference(in Variant target, ReadOnlySpan<Variant> indices, in Variant value)
    {
        if (target.IsArray)
        {
            target.AsArray().Set(Indices(indices), value);
            return;
        }

        if (target.IsObject)
        {
            if (RequireObject(target) is IDispatchObject dispatch)
            {
                dispatch.Put(DefaultMember, indices, value, asReference: true);
                return;
            }

            throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }

        throw VbaErrors.TypeMismatch();
    }

    /// <summary>CallByName(object, name, calltype, args...): vbMethod 1, vbGet 2, vbLet 4, vbSet 8 (MS-VBAL 6.1.2.6); the value of a Let or Set is the last argument.</summary>
    public static Variant CallByName(in Variant target, string name, int callType, ReadOnlySpan<Variant> arguments)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (!target.IsObject)
        {
            throw new VbaException(VbaErrors.ObjectRequired);
        }

        switch (target.AsObject())
        {
            case null:
                // Nothing is "Invalid procedure call or argument" here, not 91 (Interaction golden).
                throw VbaErrors.InvalidProcedureCall();
            case Collection when callType != 1:
                // VBA's Collection declares Count, Item, Add, and Remove as methods, so only vbMethod reaches them;
                // vbGet, vbLet, and vbSet raise 438 (Interaction golden).
                throw new VbaException(VbaErrors.ObjectDoesNotSupportMember);
        }

        switch (callType)
        {
            case 1:
                return Get(target, name, arguments);
            case 2:
                return Get(target, name, arguments);
            case 4:
            case 8:
                if (arguments.Length == 0)
                {
                    throw new VbaException(VbaErrors.ArgumentNotOptional);
                }

                var indices = arguments[..^1];
                if (callType == 4)
                {
                    Let(target, name, indices, arguments[^1]);
                }
                else
                {
                    Set(target, name, indices, arguments[^1]);
                }

                return Variant.Empty;
            default:
                throw VbaErrors.InvalidProcedureCall();
        }
    }

    private static int[] Indices(ReadOnlySpan<Variant> indices)
    {
        var result = new int[indices.Length];
        for (var i = 0; i < indices.Length; i++)
        {
            result[i] = Coerce.ToInt32(indices[i]);
        }

        return result;
    }

    private static Variant Argument(ReadOnlySpan<Variant> arguments, int index) => index < arguments.Length ? arguments[index] : Variant.Missing;

    private static object RequireObject(in Variant target)
    {
        if (!target.IsObject)
        {
            throw new VbaException(VbaErrors.ObjectRequired);
        }

        return target.AsObject() ?? throw VbaErrors.ObjectVariableNotSet();
    }
}

/// <summary>
/// A user-defined type value (MS-VBAL 5.2.3.3). Generated code emits one struct per Type with a
/// field per member, copied by value and owned on ARCHITECTURE.md D18's rails: a copy duplicates
/// what the record owns, a release drops it (ROADMAP.md M7 C3).
/// </summary>
public interface IVbaRecord
{
    /// <summary>Drops what the fields own (strings, array elements, object references, nested records' too) when the record's storage goes; the fields are left empty.</summary>
    void ReleaseReferences();
}

/// <summary>The fresh value and the copy of one user-defined type, for its assignment and for the arrays that hold it, without reflection.</summary>
public interface IVbaRecord<TSelf> : IVbaRecord
    where TSelf : struct, IVbaRecord<TSelf>
{
    /// <summary>A fresh value (MS-VBAL 5.2.3.3): every field at its initial value, fixed-size array members allocated.</summary>
    static abstract TSelf Fresh();

    /// <summary>A copy for assignment (MS-VBAL 5.4.3.8): its strings and arrays are copies, its object fields hold references of their own (D18, D20).</summary>
    TSelf Copy();
}

/// <summary>On expression GoTo | GoSub (MS-VBAL 5.4.2.13): 0 or a value past the list continues with the next statement; a negative value or one above 255 raises 5.</summary>
public static class ComputedJump
{
    public static int Target(in Variant selector, int next, ReadOnlySpan<int> targets)
    {
        var index = Coerce.ToInt32(selector);
        if (index < 0 || index > 255)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return index >= 1 && index <= targets.Length ? targets[index - 1] : next;
    }
}
