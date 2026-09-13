namespace VbaNg.Runtime;

/// <summary>
/// VBA value types, numbered as the <c>VbVarType</c> constants so <c>VarType(x)</c> returns the
/// enum value directly (MS-VBAL 2.1 data values and value types; 6.1.1 VbVarType).
/// </summary>
public enum VarType : short
{
    Empty = 0,
    Null = 1,
    Integer = 2,
    Long = 3,
    Single = 4,
    Double = 5,
    Currency = 6,
    Date = 7,
    String = 8,
    Object = 9,
    Error = 10,
    Boolean = 11,
    Variant = 12,
    DataObject = 13,
    Decimal = 14,
    Byte = 17,
    LongLong = 20,
    UserDefinedType = 36,

    /// <summary>Combined with the element type for arrays: <c>VarType(Array())</c> is 8204.</summary>
    Array = 8192,
}

public static class VarTypeExtensions
{
    /// <summary>Integer, Long, LongLong, or Byte (MS-VBAL 5.5.1.2.1 integral value types).</summary>
    public static bool IsIntegral(this VarType type) => type is VarType.Integer or VarType.Long or VarType.LongLong or VarType.Byte;

    /// <summary>Single or Double.</summary>
    public static bool IsFloatingPoint(this VarType type) => type is VarType.Single or VarType.Double;

    /// <summary>Currency or Decimal.</summary>
    public static bool IsFixedPoint(this VarType type) => type is VarType.Currency or VarType.Decimal;

    /// <summary>Any numeric value type: integral, floating-point, or fixed-point.</summary>
    public static bool IsNumeric(this VarType type) => type.IsIntegral() || type.IsFloatingPoint() || type.IsFixedPoint();

    /// <summary>The name <c>TypeName</c> reports for a scalar of this type (MS-VBAL 6.1.2.8 TypeName).</summary>
    public static string TypeName(this VarType type) => type switch
    {
        VarType.Empty => "Empty",
        VarType.Null => "Null",
        VarType.Integer => "Integer",
        VarType.Long => "Long",
        VarType.Single => "Single",
        VarType.Double => "Double",
        VarType.Currency => "Currency",
        VarType.Date => "Date",
        VarType.String => "String",
        VarType.Object => "Object",
        VarType.Error => "Error",
        VarType.Boolean => "Boolean",
        VarType.Variant => "Variant",
        VarType.DataObject => "DataObject",
        VarType.Decimal => "Decimal",
        VarType.Byte => "Byte",
        VarType.LongLong => "LongLong",
        VarType.UserDefinedType => "UserDefinedType",
        _ => "Unknown",
    };
}
