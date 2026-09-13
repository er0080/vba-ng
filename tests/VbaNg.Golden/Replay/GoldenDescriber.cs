using System.Globalization;
using System.Text;

using VbaNg.Golden.Harness;
using VbaNg.Runtime;
using VbaNg.Runtime.Library;

namespace VbaNg.Golden.Replay;

/// <summary>
/// Describes a runtime value exactly as the VBA recorder module does (RecorderModule.Recorder),
/// so a replayed value and a recorded one compare field by field.
/// </summary>
public static class GoldenDescriber
{
    public static GoldenValue Describe(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Object:
                return value.IsNothing ? new GoldenValue("Nothing") : new GoldenValue("Object", Class: Information.ObjectTypeName(value.AsObject()!));
            case VarType.Array:
                return DescribeArray(value.AsArray());
            case VarType.Empty:
            case VarType.Null:
                return new GoldenValue(value.Type.TypeName());
            case VarType.Boolean:
                return new GoldenValue("Boolean", value.AsBoolean() ? "True" : "False");
            case VarType.Integer:
            case VarType.Long:
            case VarType.LongLong:
            case VarType.Byte:
                return new GoldenValue(value.Type.TypeName(), Coerce.ToInt64(value).ToString(CultureInfo.InvariantCulture));
            case VarType.Single:
                return new GoldenValue("Single", StrText(NumberText.FormatSingle(value.AsSingle())), Bits: BitConverter.SingleToInt32Bits(value.AsSingle()).ToString("X8", CultureInfo.InvariantCulture));
            case VarType.Double:
                return new GoldenValue("Double", StrText(NumberText.FormatDouble(value.AsDouble())), Bits: BitConverter.DoubleToInt64Bits(value.AsDouble()).ToString("X16", CultureInfo.InvariantCulture));
            case VarType.Date:
                {
                    var date = value.AsDate();
                    return new GoldenValue("Date", date.ToDateTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), Bits: BitConverter.DoubleToInt64Bits(date.Serial).ToString("X16", CultureInfo.InvariantCulture));
                }

            case VarType.Currency:
                {
                    var currency = value.AsCurrency();
                    return new GoldenValue("Currency", StrText(currency.ToString()), Bits: currency.Scaled.ToString("X16", CultureInfo.InvariantCulture));
                }

            case VarType.Decimal:
                {
                    var number = value.AsDecimal();
                    return new GoldenValue("Decimal", StrText(DecimalText.Normalize(number).ToString(CultureInfo.InvariantCulture)), Bytes: DecimalBytes(number));
                }

            case VarType.Error:
                return new GoldenValue("Error", value.AsError().Scode.ToString(CultureInfo.InvariantCulture));
            case VarType.String:
                return DescribeString(value.AsString());
            default:
                return new GoldenValue(value.Type.TypeName(), ((int)value.Type).ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>Str$ drops the zero before a decimal point: 0.5 prints as ".5" and -0.5 as "-.5".</summary>
    private static string StrText(string text)
    {
        if (text.StartsWith("0.", StringComparison.Ordinal))
        {
            return text[1..];
        }

        if (text.StartsWith("-0.", StringComparison.Ordinal))
        {
            return "-" + text[2..];
        }

        return text;
    }

    private static GoldenValue DescribeString(string text)
    {
        var lone = false;
        var shown = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                shown.Append(c).Append(text[i + 1]);
                i++;
            }
            else if (char.IsSurrogate(c))
            {
                lone = true;
                shown.Append('�');
            }
            else
            {
                shown.Append(c);
            }
        }

        if (!lone)
        {
            return new GoldenValue("String", text);
        }

        var units = new StringBuilder(text.Length * 4);
        foreach (var c in text)
        {
            units.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
        }

        return new GoldenValue("String", shown.ToString(), Units: units.ToString());
    }

    private static GoldenValue DescribeArray(VbaArray array)
    {
        var type = array.ElementType.TypeName() + "()";
        if (array.Rank == 1)
        {
            var items = new List<GoldenValue>(array.Count);
            for (var i = array.LBound(1); i <= array.UBound(1); i++)
            {
                items.Add(Describe(array[[i]]));
            }

            return new GoldenValue(type, Bounds: array.BoundsText, Items: items);
        }

        if (array.Rank == 2)
        {
            // The recorder walks the first dimension in the outer loop.
            var items = new List<GoldenValue>(array.Count);
            for (var i = array.LBound(1); i <= array.UBound(1); i++)
            {
                for (var j = array.LBound(2); j <= array.UBound(2); j++)
                {
                    items.Add(Describe(array[[i, j]]));
                }
            }

            return new GoldenValue(type, Bounds: array.BoundsText, Items: items);
        }

        return new GoldenValue(type, Bounds: array.BoundsText);
    }

    /// <summary>The 16 VARIANT bytes of a Decimal: vt 14, scale, sign, Hi32, Lo64, little-endian, as CopyMemory reads them.</summary>
    private static string DecimalBytes(decimal value)
    {
        var parts = decimal.GetBits(value);
        var lo = (uint)parts[0];
        var mid = (uint)parts[1];
        var hi = (uint)parts[2];
        var scale = (byte)((parts[3] >> 16) & 0xFF);
        var sign = (byte)(parts[3] < 0 ? 0x80 : 0);
        var bytes = new byte[16];
        bytes[0] = 0x0E;
        bytes[1] = 0x00;
        bytes[2] = scale;
        bytes[3] = sign;
        BitConverter.TryWriteBytes(bytes.AsSpan(4), hi);
        BitConverter.TryWriteBytes(bytes.AsSpan(8), lo);
        BitConverter.TryWriteBytes(bytes.AsSpan(12), mid);
        return Convert.ToHexString(bytes);
    }
}

/// <summary>Compares a recorded value with a replayed one, field by field, and names the first difference.</summary>
public static class GoldenMatch
{
    /// <summary>Null when the values agree, otherwise what differs.</summary>
    public static string? Compare(GoldenValue expected, GoldenValue actual, string path = "value")
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);
        if (expected.Type != actual.Type)
        {
            return $"{path}: expected {expected.Type} {Show(expected)}, got {actual.Type} {Show(actual)}";
        }

        switch (expected.Type)
        {
            case "Single":
            case "Double":
            case "Date":
            case "Currency":
                if (expected.Bits != actual.Bits)
                {
                    return $"{path}: expected {expected.Type} {expected.Value} ({expected.Bits}), got {actual.Value} ({actual.Bits})";
                }

                return null;
            case "Decimal":
                if (expected.Bytes != actual.Bytes)
                {
                    return $"{path}: expected Decimal {expected.Value} ({expected.Bytes}), got {actual.Value} ({actual.Bytes})";
                }

                return null;
            case "Object":
                return expected.Class == actual.Class ? null : $"{path}: expected object {expected.Class}, got {actual.Class}";
            case "Empty":
            case "Null":
            case "Nothing":
                return null;
        }

        if (expected.Type.EndsWith("()", StringComparison.Ordinal))
        {
            if (expected.Bounds != actual.Bounds)
            {
                return $"{path}: expected bounds \"{expected.Bounds}\", got \"{actual.Bounds}\"";
            }

            var expectedItems = expected.Items ?? [];
            var actualItems = actual.Items ?? [];
            if (expectedItems.Count != actualItems.Count)
            {
                return $"{path}: expected {expectedItems.Count} items, got {actualItems.Count}";
            }

            for (var i = 0; i < expectedItems.Count; i++)
            {
                var difference = Compare(expectedItems[i], actualItems[i], $"{path}[{i}]");
                if (difference is not null)
                {
                    return difference;
                }
            }

            return null;
        }

        if (expected.Value != actual.Value || expected.Units != actual.Units)
        {
            return $"{path}: expected {expected.Type} {Show(expected)}, got {Show(actual)}";
        }

        return null;
    }

    private static string Show(GoldenValue value) => value.Value is null ? string.Empty : "\"" + value.Value.Replace("\r", "\\r", StringComparison.Ordinal).Replace("\n", "\\n", StringComparison.Ordinal) + "\"";
}
