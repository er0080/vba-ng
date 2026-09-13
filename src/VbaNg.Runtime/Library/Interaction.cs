namespace VbaNg.Runtime.Library;

/// <summary>The Interaction module of the VBA standard library (MS-VBAL 6.1.2.6), the parts that need no host.</summary>
public static partial class Interaction
{
    /// <summary>Array(...): a Variant array of the arguments, based on the module's Option Base.</summary>
    public static Variant Array(ReadOnlySpan<Variant> items, int optionBase = 0) =>
        Variant.FromArray(VbaArray.FromValues(items, optionBase));

    /// <summary>Choose(Index, ...): the 1-based choice, or Null when the index is out of range; every choice is evaluated by the caller.</summary>
    public static Variant Choose(in Variant index, ReadOnlySpan<Variant> choices)
    {
        var position = (int)Math.Truncate(Coerce.ToDouble(index));
        return position >= 1 && position <= choices.Length ? choices[position - 1] : Variant.Null;
    }

    /// <summary>IIf(Expression, TruePart, FalsePart); the caller has evaluated both parts already.</summary>
    public static Variant IIf(in Variant condition, in Variant truePart, in Variant falsePart) =>
        !condition.IsNull && Coerce.ToBoolean(condition) ? truePart : falsePart;

    /// <summary>Switch(expr1, value1, ...): the value paired with the first True expression, Null when none is; an odd count raises 5.</summary>
    public static Variant Switch(ReadOnlySpan<Variant> arguments)
    {
        if (arguments.Length % 2 != 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        for (var i = 0; i < arguments.Length; i += 2)
        {
            if (!arguments[i].IsNull && Coerce.ToBoolean(arguments[i]))
            {
                return arguments[i + 1];
            }
        }

        return Variant.Null;
    }

    /// <summary>CreateObject(class, [servername]): a new COM object by ProgID through the host's provider (ARCHITECTURE.md section 6).</summary>
    public static Variant CreateObject(in Variant progId, in Variant serverName) =>
        Variant.FromObject(Com.Provider.CreateObject(Coerce.ToString(progId), serverName.IsMissing ? null : Coerce.ToString(serverName)));

    /// <summary>GetObject([pathname], [class]): a running object or one bound from a file moniker.</summary>
    public static Variant GetObject(in Variant pathName, in Variant className) =>
        Variant.FromObject(Com.Provider.GetObject(pathName.IsMissing ? null : Coerce.ToString(pathName), className.IsMissing ? null : Coerce.ToString(className)));

    /// <summary>CallByName(object, procname, calltype, args...).</summary>
    public static Variant CallByName(in Variant target, in Variant procName, in Variant callType, ReadOnlySpan<Variant> arguments) =>
        LateBound.CallByName(target, Coerce.ToString(procName), Coerce.ToInt32(callType), arguments);

    /// <summary>DoEvents returns 0 (an Integer) outside a form host.</summary>
    public static short DoEvents() => 0;

    /// <summary>
    /// Partition(Number, Start, Stop, Interval): "lower:upper" of the range holding Number, both
    /// sides padded to the width of Stop + 1; below Start the lower side is blank, above Stop the
    /// upper side is. A zero interval, a negative Start, or Start above Stop raises 5.
    /// </summary>
    public static Variant Partition(in Variant number, in Variant start, in Variant stop, in Variant interval)
    {
        if (number.IsNull || start.IsNull || stop.IsNull || interval.IsNull)
        {
            return Variant.Null;
        }

        var value = Coerce.ToInt64(number);
        var first = Coerce.ToInt64(start);
        var last = Coerce.ToInt64(stop);
        var step = Coerce.ToInt64(interval);
        if (step < 1 || first < 0 || last <= first)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var width = (last + 1).ToString(System.Globalization.CultureInfo.InvariantCulture).Length;
        string lower;
        string upper;
        if (value < first)
        {
            lower = string.Empty;
            upper = (first - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        else if (value > last)
        {
            lower = (last + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            upper = string.Empty;
        }
        else
        {
            var index = (value - first) / step;
            var low = first + (index * step);
            var high = Math.Min(low + step - 1, last);
            lower = low.ToString(System.Globalization.CultureInfo.InvariantCulture);
            upper = high.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return Variant.FromString(lower.PadLeft(width) + ":" + upper.PadLeft(width));
    }

    /// <summary>Environ(Name | Index): a variable's value, "" for an unknown name; an empty name or an index out of range raises 5.</summary>
    public static VbaString Environ(in Variant argument) => VbaString.Temporary(EnvironText(argument));

    private static string EnvironText(in Variant argument)
    {
        if (argument.IsNull)
        {
            throw VbaErrors.InvalidUseOfNull();
        }

        if (argument.IsString)
        {
            var name = argument.AsString();
            if (name.Length == 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            return Environment.GetEnvironmentVariable(name) ?? string.Empty;
        }

        var index = Coerce.ToInt32(argument);
        var variables = Environment.GetEnvironmentVariables().Keys.Cast<string>().Order(StringComparer.OrdinalIgnoreCase).ToList();
        if (index < 1 || index > variables.Count)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var key = variables[index - 1];
        return key + "=" + Environment.GetEnvironmentVariable(key);
    }
}
