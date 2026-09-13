using System.Globalization;

namespace VbaNg.Runtime.Library;

using VbaType = VbaNg.Runtime.VarType;

/// <summary>The Information module of the VBA standard library (MS-VBAL 6.1.2.5).</summary>
public static class Information
{
    public static bool IsArray(in Variant value) => value.IsArray;

    public static bool IsDate(in Variant value)
    {
        if (value.IsDate)
        {
            return true;
        }

        if (!value.IsString)
        {
            return false;
        }

        // A string that is only a number is not a date to IsDate, although CDate accepts it (Information golden).
        return DateText.TryParse(value.AsVbaString().Chars, Coerce.Culture, out _, allowNumber: false);
    }

    public static bool IsEmpty(in Variant value) => value.IsEmpty;

    public static bool IsError(in Variant value) => value.IsError;

    public static bool IsMissing(in Variant value) => value.IsMissing;

    public static bool IsNull(in Variant value) => value.IsNull;

    /// <summary>Numbers, Booleans, Empty, and strings that read as numbers; Dates are not numeric.</summary>
    public static bool IsNumeric(in Variant value)
    {
        switch (value.Type)
        {
            case VbaType.Empty:
            case VbaType.Boolean:
                return true;
            case VbaType.String:
                return NumberText.TryParse(value.AsVbaString().Chars, Coerce.Culture, out _, out _, out _);
            case VbaType.Object:
                return false;
            default:
                return value.IsNumeric;
        }
    }

    public static bool IsObject(in Variant value) => value.IsObject;

    /// <summary>LBound of a 1-based dimension (error 9 when the array is unallocated or the dimension is out of range; error 13 for a non-array).</summary>
    public static int LBound(in Variant array, int dimension = 1) => RequireArray(array).LBound(dimension);

    public static int UBound(in Variant array, int dimension = 1) => RequireArray(array).UBound(dimension);

    /// <summary>MS-VBAL 6.1.2.5 TypeName, as a String the statement owns (D20).</summary>
    public static VbaString TypeName(in Variant value) => VbaString.Temporary(TypeNameText(value));

    /// <summary>MS-VBAL 6.1.2.5 TypeName, as the .NET string the runtime's own messages use.</summary>
    public static string TypeNameText(in Variant value)
    {
        switch (value.Type)
        {
            case VbaType.Array:
                return value.AsArray().ElementType.TypeName() + "()";
            case VbaType.Object:
                return value.IsNothing ? "Nothing" : ObjectTypeName(value.AsObject()!);
            default:
                return value.Type.TypeName();
        }
    }

    /// <summary>
    /// MS-VBAL 6.1.2.5 VarType: the VbVarType constant, with vbArray added for arrays. An object
    /// reports the type of its default member's value, and vbObject when it has none, when the
    /// member needs arguments, or when it is Nothing (Classes and Objects goldens).
    /// </summary>
    public static int VarType(in Variant value)
    {
        if (value.IsObject && !value.IsNothing)
        {
            try
            {
                var inner = Coerce.ObjectValue(value);
                return inner.IsObject ? (int)Runtime.VarType.Object : (int)inner.VarTypeValue;
            }
            catch (VbaException)
            {
                return (int)Runtime.VarType.Object;
            }
        }

        return (int)value.VarTypeValue;
    }

    /// <summary>The name VBA reports for a runtime object: the class name for library classes, the COM type name for COM objects.</summary>
    public static string ObjectTypeName(object value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value switch
        {
            IVbaObject vba => vba.TypeName,
            _ => value.GetType().Name,
        };
    }

    /// <summary>
    /// RGB(red, green, blue) (MS-VBAL 6.1.2.8.1.13): the color as red + green * 256 + blue * 65536,
    /// each component an Integer, above 255 taken as 255, below 0 error 5 (Information golden).
    /// </summary>
    public static int RGB(in Variant red, in Variant green, in Variant blue) =>
        ColorComponent(red) | (ColorComponent(green) << 8) | (ColorComponent(blue) << 16);

    /// <summary>QBColor(color) (MS-VBAL 6.1.2.8.1.12): the RGB value of one of the sixteen QuickBasic colors, error 5 outside 0 to 15.</summary>
    public static int QBColor(in Variant color)
    {
        var index = Coerce.ToInt16(color);
        return index is >= 0 and <= 15 ? QuickBasicColors[index] : throw VbaErrors.InvalidProcedureCall();
    }

    private static readonly int[] QuickBasicColors =
    [
        0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0,
        0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF,
    ];

    private static int ColorComponent(in Variant value)
    {
        var component = Coerce.ToInt16(value);
        return component < 0 ? throw VbaErrors.InvalidProcedureCall() : Math.Min(component, (short)255);
    }

    private static VbaArray RequireArray(in Variant value) => value.IsArray ? value.AsArray() : throw VbaErrors.TypeMismatch();

    internal static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);
}

/// <summary>A runtime-provided object VBA code can hold in an Object variable, such as Collection or Err.</summary>
public interface IVbaObject
{
    /// <summary>The name <c>TypeName</c> reports, for example "Collection".</summary>
    string TypeName { get; }
}
