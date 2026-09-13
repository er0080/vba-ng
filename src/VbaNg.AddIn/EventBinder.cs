using System.Globalization;
using System.Reflection;

using VbaNg.Interop;
using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

/// <summary>
/// Connects the event procedures of a document module to the objects that raise them
/// (ARCHITECTURE.md section 6, "Events"): one sink per source object, its events mapped by name
/// to the handlers' dispids, arguments converted to the handlers' parameter types, ByRef
/// parameters written back. Everything advised here is undone by <see cref="Dispose"/>.
/// </summary>
internal sealed class EventBinder : IDisposable
{
    private readonly List<EventSink> sinks = [];
    private readonly Action<string> log;
    private readonly string? projectDir;

    /// <param name="log">Where notes on the binding go.</param>
    /// <param name="projectDir">The project the handlers belong to, whose log gets what they print and the errors they leave unhandled (ROADMAP.md WP4).</param>
    public EventBinder(Action<string> log, string? projectDir = null)
    {
        this.log = log;
        this.projectDir = projectDir;
    }

    public int SinkCount => sinks.Count;

    /// <summary>
    /// Advises the handlers of one source object; a handler whose event the source does not raise
    /// is skipped with a note. The source interface comes from the object's class information
    /// (ActiveX controls) or, failing that, from <paramref name="fallback"/>, the type library's
    /// description of the object's coclass (Excel's own objects).
    /// </summary>
    public void Bind(ComObject source, string sourceName, IReadOnlyList<MethodInfo> handlers, EventSource? fallback = null)
    {
        var description = EventSource.Describe(source) ?? fallback;
        if (description is null)
        {
            log($"{sourceName}: no event source interface; {handlers.Count} handler(s) not connected.");
            return;
        }

        var byDispId = new Dictionary<int, MethodInfo>();
        foreach (var handler in handlers)
        {
            var eventName = handler.GetCustomAttribute<VbaEventHandlerAttribute>()!.EventName;
            if (description.Events.TryGetValue(eventName, out var comEvent))
            {
                byDispId[comEvent.DispId] = handler;
            }
            else
            {
                log($"{sourceName}: {description.InterfaceName} has no event '{eventName}'; {handler.Name} not connected.");
            }
        }

        if (byDispId.Count == 0)
        {
            return;
        }

        var reorder = description.DeclaredOrder;
        var sink = new EventSink(description.InterfaceId, (dispId, arguments) =>
        {
            if (!byDispId.TryGetValue(dispId, out var handler))
            {
                return;
            }

            // The sink decodes rgvarg as the convention says; a declared-order source needs the reverse, and the
            // ByRef write-back happens by position, so the array is put back afterwards.
            if (reorder && arguments.Length > 1)
            {
                Array.Reverse(arguments);
            }

            using var running = ExcelHostServices.Instance.Running(projectDir);
            try
            {
                Invoke(handler, arguments);
            }
            finally
            {
                if (reorder && arguments.Length > 1)
                {
                    Array.Reverse(arguments);
                }
            }
        })
        {
            Error = ex =>
            {
                // End in a handler stops it and resets the project (docs/vba-quirks.md); anything else is reported.
                if (ex is EndStatementException end)
                {
                    VbaNg.Runtime.Hosting.ProjectReset.AfterEnd(end);
                }
                else
                {
                    using var running = ExcelHostServices.Instance.Running(projectDir);
                    ExcelHostServices.Instance.ReportUnhandled($"{sourceName} event", ex is TargetInvocationException { InnerException: { } inner } ? inner : ex, dialog: true);
                }
            },
        };
        source.Advise(sink);
        sinks.Add(sink);
    }

    public void Dispose()
    {
        foreach (var sink in sinks)
        {
            sink.Dispose();
        }

        sinks.Clear();
    }

    /// <summary>The arguments of an event in declared order, whatever order the source used.</summary>
    internal static Variant[] InDeclaredOrder(EventSource source, Variant[] arguments)
    {
        if (!source.DeclaredOrder || arguments.Length < 2)
        {
            return arguments;
        }

        var ordered = (Variant[])arguments.Clone();
        Array.Reverse(ordered);
        return ordered;
    }

    /// <summary>Calls a function with its arguments converted to its parameter types and returns the result as a Variant; a left-out optional parameter takes its declared default.</summary>
    internal static Variant InvokeFunction(MethodInfo function, Variant[] arguments)
    {
        var parameters = function.GetParameters();
        var values = new object?[parameters.Length];
        for (var i = 0; i < parameters.Length; i++)
        {
            var type = parameters[i].ParameterType;
            var target = type.IsByRef ? type.GetElementType()! : type;
            var missing = i >= arguments.Length || arguments[i].IsMissing;
            values[i] = missing && parameters[i].IsOptional ? Type.Missing
                : missing ? Default(target)
                : ToClr(arguments[i], target, owned: type.IsByRef);
        }

        try
        {
            return FromClr(function.Invoke(null, values));
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
        finally
        {
            // What a ByRef parameter left behind is the call's to release: a UDF's caller does not read it.
            for (var i = 0; i < parameters.Length; i++)
            {
                if (parameters[i].ParameterType.IsByRef)
                {
                    Adopt(values[i]);
                }
            }
        }
    }

    /// <summary>Calls a handler with the event's arguments converted to its parameter types, then copies ByRef parameters back.</summary>
    internal static void Invoke(MethodInfo handler, Variant[] arguments)
    {
        var parameters = handler.GetParameters();
        var values = new object?[parameters.Length];
        for (var i = 0; i < parameters.Length; i++)
        {
            var type = parameters[i].ParameterType;
            var target = type.IsByRef ? type.GetElementType()! : type;
            values[i] = i < arguments.Length ? ToClr(arguments[i], target, owned: type.IsByRef) : Default(target);
        }

        try
        {
            handler.Invoke(null, values);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
        }
        finally
        {
            for (var i = 0; i < parameters.Length; i++)
            {
                if (parameters[i].ParameterType.IsByRef)
                {
                    var left = Adopt(values[i]);
                    if (i < arguments.Length)
                    {
                        arguments[i] = left;
                    }
                }
            }
        }
    }

    /// <summary>
    /// The argument as the parameter's CLR type. A ByRef Variant or String slot is the callee's
    /// to reassign, which frees what it held, so such a slot starts with a value of its own
    /// (<paramref name="owned"/>, ARCHITECTURE.md D20); a ByVal one is a view the callee copies.
    /// </summary>
    private static object? ToClr(in Variant value, Type target, bool owned = false)
    {
        if (target.IsGenericType && target.GetGenericTypeDefinition() == typeof(ObjectSlot<>))
        {
            // A ByRef object parameter is a slot holding a reference of its own for the call (ROADMAP.md M7 E3).
            var slot = (IObjectSlot)Activator.CreateInstance(target)!;
            slot.Assign(ToClr(value, target.GetGenericArguments()[0]));
            return slot;
        }

        if (target == typeof(Variant))
        {
            return owned ? ObjectRefs.Own(value) : value;
        }

        if (target == typeof(VbaString))
        {
            var text = Coerce.ToText(value);
            return owned ? ObjectRefs.Own(text) : text;
        }

        if (target == typeof(IDispatchObject))
        {
            return Coerce.ToObject<IDispatchObject>(value);
        }

        if (target == typeof(object))
        {
            return value.IsObject ? value.AsObject() : value;
        }

        if (target == typeof(int))
        {
            return Coerce.ToInt32(value);
        }

        if (target == typeof(short))
        {
            return Coerce.ToInt16(value);
        }

        if (target == typeof(long))
        {
            return Coerce.ToInt64(value);
        }

        if (target == typeof(double))
        {
            return Coerce.ToDouble(value);
        }

        if (target == typeof(float))
        {
            return Coerce.ToSingle(value);
        }

        if (target == typeof(string))
        {
            return Coerce.ToString(value);
        }

        if (target == typeof(bool))
        {
            return Coerce.ToBoolean(value);
        }

        if (target == typeof(byte))
        {
            return Coerce.ToByte(value);
        }

        if (target == typeof(Currency))
        {
            return Coerce.ToCurrency(value);
        }

        if (target == typeof(VbaDate))
        {
            return Coerce.ToDate(value);
        }

        if (target == typeof(VbaArray))
        {
            return value.IsArray ? value.AsArray() : throw VbaErrors.TypeMismatch();
        }

        throw new InvalidOperationException("Event handlers cannot take a parameter of type " + target.FullName + ".");
    }

    private static object? Default(Type target) =>
        target == typeof(Variant) ? Variant.Empty : target.IsValueType ? Activator.CreateInstance(target) : null;

    /// <summary>What a ByRef slot the call owned holds when it returns, handed to the host's frame as a temporary (<see cref="ObjectRefs.Frame"/>).</summary>
    private static Variant Adopt(object? value)
    {
        switch (value)
        {
            case VbaString text:
                return Variant.FromVbaString(text);
            case Variant variant:
                return ObjectRefs.Adopt(ref variant);
            case IObjectSlot slot:
                return ObjectRefs.OwnedInterface(slot.Pointer);
            default:
                return FromClr(value);
        }
    }

    /// <summary>A value a procedure returned as a Variant: a String result is already a temporary of the frame, transferred by the callee (D20).</summary>
    private static Variant FromClr(object? value) => value switch
    {
        null => Variant.Nothing,
        Variant variant => variant,
        VbaString text => Variant.ViewString(text),
        bool b => Variant.FromBoolean(b),
        int i => Variant.FromInt32(i),
        short s => Variant.FromInt16(s),
        long l => Variant.FromInt64(l),
        double d => Variant.FromDouble(d),
        float f => Variant.FromSingle(f),
        string text => Variant.FromString(text),
        byte b => Variant.FromByte(b),
        Currency c => Variant.FromCurrency(c),
        VbaDate d => Variant.FromDate(d),
        VbaArray a => Variant.FromArray(a),
        _ => Variant.FromObject(value),
    };

    internal static string Describe(Exception ex) =>
        (ex is UnhandledErrorException { Error: { } unhandled } ? unhandled : ex as VbaException) is { } error
            ? string.Create(CultureInfo.InvariantCulture, $"'{error.Number}': {error.Description}")
            : ex.GetType().Name + ": " + ex.Message;
}
