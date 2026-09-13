using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

public enum TypeKind
{
    /// <summary>One of the reserved type identifiers, Variant, or the generic Object (MS-VBAL 5.6.16.7).</summary>
    Builtin,

    /// <summary>String * n (MS-VBAL 5.2.3.1.1).</summary>
    FixedString,

    /// <summary>A class: Collection and ErrObject from the runtime; COM and project classes in later milestones.</summary>
    Class,

    /// <summary>A user-defined type (MS-VBAL 5.2.3.3).</summary>
    Record,

    /// <summary>An Enum (MS-VBAL 5.2.3.4); its values are Long.</summary>
    Enum,

    /// <summary>An array of an element type (MS-VBAL 5.2.3.1.3); the bounds are not part of the type.</summary>
    Array,
}

/// <summary>
/// A declared type (MS-VBAL 2.2 declared types): what a variable, parameter, procedure result,
/// or expression is statically known to be. The binder types every expression with one of
/// these; the emitter maps each to the C# type generated code uses.
/// </summary>
public sealed class VbaType
{
    private readonly string csharpName;

    private VbaType(TypeKind kind, VarType varType, string name, string csharpName)
    {
        Kind = kind;
        VarType = varType;
        Name = name;
        this.csharpName = csharpName;
    }

    public static readonly VbaType Integer = new(TypeKind.Builtin, VarType.Integer, "Integer", "short");
    public static readonly VbaType Long = new(TypeKind.Builtin, VarType.Long, "Long", "int");
    public static readonly VbaType LongLong = new(TypeKind.Builtin, VarType.LongLong, "LongLong", "long");
    public static readonly VbaType Single = new(TypeKind.Builtin, VarType.Single, "Single", "float");
    public static readonly VbaType Double = new(TypeKind.Builtin, VarType.Double, "Double", "double");
    public static readonly VbaType Currency = new(TypeKind.Builtin, VarType.Currency, "Currency", "global::VbaNg.Runtime.Currency");
    public static readonly VbaType Date = new(TypeKind.Builtin, VarType.Date, "Date", "global::VbaNg.Runtime.VbaDate");
    public static readonly VbaType String = new(TypeKind.Builtin, VarType.String, "String", "global::VbaNg.Runtime.VbaString");
    public static readonly VbaType Boolean = new(TypeKind.Builtin, VarType.Boolean, "Boolean", "bool");
    public static readonly VbaType Byte = new(TypeKind.Builtin, VarType.Byte, "Byte", "byte");
    public static readonly VbaType Variant = new(TypeKind.Builtin, VarType.Variant, "Variant", "global::VbaNg.Runtime.Variant");

    /// <summary>The generic Object type: any object reference, bound late (MS-VBAL 2.2 Object).</summary>
    public static readonly VbaType Object = new(TypeKind.Builtin, VarType.Object, "Object", "object?");

    /// <summary>Decimal values exist only inside Variants (MS-VBAL 2.1 Decimal); this type describes expressions such as CDec(x), never declarations.</summary>
    public static readonly VbaType Decimal = new(TypeKind.Builtin, VarType.Decimal, "Decimal", "decimal");

    /// <summary>As Any, valid only for the parameter of a Declare (MS-VBAL 5.2.3.5): the variable's own storage travels to the DLL.</summary>
    public static readonly VbaType Any = new(TypeKind.Builtin, VarType.Empty, "Any", "byte");

    public static readonly VbaType Collection = new(TypeKind.Class, VarType.Object, "Collection", "global::VbaNg.Runtime.Library.Collection?");
    public static readonly VbaType ErrObject = new(TypeKind.Class, VarType.Object, "ErrObject", "global::VbaNg.Runtime.Library.ErrObject?");

    public TypeKind Kind { get; }

    /// <summary>The runtime tag values of this type carry: the scalar type, Object for classes, Long for enums, String for fixed strings, Array for arrays, UserDefinedType for records.</summary>
    public VarType VarType { get; }

    /// <summary>The VBA name, as it appears in messages: Long, String * 10, Collection, TPoint, Long().</summary>
    public string Name { get; }

    /// <summary>The C# type generated code uses for a variable of this type, fully qualified.</summary>
    public string CSharpName => csharpName;

    /// <summary>Length of a fixed-length string; 0 otherwise.</summary>
    public int FixedLength { get; private init; }

    /// <summary>Element type of an array.</summary>
    public VbaType? ElementType { get; private init; }

    public RecordSymbol? Record { get; private init; }

    /// <summary>For a project class type: the class module it names (MS-VBAL 5.2.4.1.3).</summary>
    public ModuleSymbol? ProjectClass { get; private init; }

    public EnumSymbol? Enum { get; private init; }

    /// <summary>For a COM type: the library it comes from.</summary>
    public Runtime.TypeLibraries.ComLibrary? Library { get; private init; }

    /// <summary>For a COM type: the coclass or interface the declaration named.</summary>
    public Runtime.TypeLibraries.ComType? ComType { get; private init; }

    /// <summary>For a COM type: the interface whose members and dispids the binder uses (a coclass's default interface).</summary>
    public Runtime.TypeLibraries.ComType? ComInterface { get; private init; }

    /// <summary>A class from a referenced type library, bound by dispid (ARCHITECTURE.md section 6).</summary>
    public bool IsCom => ComInterface is not null;

    internal static VbaType ForCom(Runtime.TypeLibraries.ComLibrary library, Runtime.TypeLibraries.ComType type, Runtime.TypeLibraries.ComType members) =>
        new(TypeKind.Class, VarType.Object, type.Name, "global::VbaNg.Runtime.IDispatchObject?") { Library = library, ComType = type, ComInterface = members };

    public bool IsVariant => Kind == TypeKind.Builtin && VarType == VarType.Variant;

    public bool IsArray => Kind == TypeKind.Array;

    public bool IsRecord => Kind == TypeKind.Record;

    public bool IsEnum => Kind == TypeKind.Enum;

    /// <summary>A class or the generic Object.</summary>
    public bool IsObject => Kind == TypeKind.Class || (Kind == TypeKind.Builtin && VarType == VarType.Object);

    public bool IsGenericObject => Kind == TypeKind.Builtin && VarType == VarType.Object;

    public bool IsString => Kind == TypeKind.FixedString || (Kind == TypeKind.Builtin && VarType == VarType.String);

    /// <summary>A variable-length String: a VbaString slot that owns a BSTR in generated code (ARCHITECTURE.md D20), unlike a fixed-length string, which is still a .NET string.</summary>
    public bool IsVariableString => Kind == TypeKind.Builtin && VarType == VarType.String;

    public bool IsNumeric => (Kind == TypeKind.Builtin && VarType.IsNumeric()) || Kind == TypeKind.Enum;

    /// <summary>A value that a Variant can hold and that Coerce can produce: scalars, strings, and Variant itself.</summary>
    public bool IsScalar => Kind is TypeKind.Builtin or TypeKind.FixedString or TypeKind.Enum && !IsObject;

    /// <summary>The element type an array or Variant indexes to; Variant for anything bound late.</summary>
    public VbaType IndexedType => ElementType ?? Variant;

    public static VbaType FixedString(int length) =>
        new(TypeKind.FixedString, VarType.String, "String * " + length.ToString(System.Globalization.CultureInfo.InvariantCulture), "string") { FixedLength = length };

    public static VbaType ArrayOf(VbaType element)
    {
        ArgumentNullException.ThrowIfNull(element);
        return new VbaType(TypeKind.Array, VarType.Array, element.Name + "()", "global::VbaNg.Runtime.VbaArray") { ElementType = element };
    }

    internal static VbaType ForRecord(RecordSymbol record, string csharpName) =>
        new(TypeKind.Record, VarType.UserDefinedType, record.Name, csharpName) { Record = record };

    /// <summary>
    /// A project class (MS-VBAL 5.2.4.1.3). Variables of the type hold the class's C# name, which
    /// is the generated interface when another class implements it and the class itself otherwise;
    /// either way the name is the module's, so the type does not change when an Implements appears.
    /// </summary>
    internal static VbaType ForClass(ModuleSymbol module) =>
        new(TypeKind.Class, VarType.Object, module.Name, "global::" + module.EmitName + "?") { ProjectClass = module };

    internal static VbaType ForEnum(EnumSymbol enumeration) =>
        new(TypeKind.Enum, VarType.Long, enumeration.Name, "int") { Enum = enumeration };

    /// <summary>The builtin type a reserved type identifier names (MS-VBAL 5.6.16.7), or null.</summary>
    public static VbaType? FromKeyword(SyntaxKind kind) => kind switch
    {
        SyntaxKind.IntegerKeyword => Integer,
        SyntaxKind.LongKeyword => Long,
        SyntaxKind.LongLongKeyword or SyntaxKind.LongPtrKeyword => LongLong,
        SyntaxKind.SingleKeyword => Single,
        SyntaxKind.DoubleKeyword => Double,
        SyntaxKind.CurrencyKeyword => Currency,
        SyntaxKind.DateKeyword => Date,
        SyntaxKind.StringKeyword => String,
        SyntaxKind.BooleanKeyword => Boolean,
        SyntaxKind.ByteKeyword => Byte,
        SyntaxKind.VariantKeyword => Variant,
        _ => null,
    };

    /// <summary>The builtin type a scalar runtime value has.</summary>
    public static VbaType FromVarType(VarType varType) => varType switch
    {
        VarType.Integer => Integer,
        VarType.Long => Long,
        VarType.LongLong => LongLong,
        VarType.Single => Single,
        VarType.Double => Double,
        VarType.Currency => Currency,
        VarType.Date => Date,
        VarType.String => String,
        VarType.Boolean => Boolean,
        VarType.Byte => Byte,
        VarType.Decimal => Decimal,
        VarType.Object => Object,
        _ => Variant,
    };

    /// <summary>The type suffix character of a typed name (MS-VBAL 3.3.5.3): % & ^ ! # @ $.</summary>
    public static VbaType? FromSuffix(char suffix) => suffix switch
    {
        '%' => Integer,
        '&' => Long,
        '^' => LongLong,
        '!' => Single,
        '#' => Double,
        '@' => Currency,
        '$' => String,
        _ => null,
    };

    /// <summary>Two types are the same when they are the same builtin, the same class, record, or enum, or arrays of the same element type.</summary>
    public bool SameAs(VbaType other)
    {
        ArgumentNullException.ThrowIfNull(other);
        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Kind == other.Kind && Kind switch
        {
            TypeKind.Builtin => VarType == other.VarType,
            TypeKind.FixedString => FixedLength == other.FixedLength,
            TypeKind.Class => ProjectClass is not null || other.ProjectClass is not null
                ? ReferenceEquals(ProjectClass, other.ProjectClass)
                : Name.Equals(other.Name, StringComparison.OrdinalIgnoreCase),
            TypeKind.Record => ReferenceEquals(Record, other.Record),
            TypeKind.Enum => ReferenceEquals(Enum, other.Enum),
            TypeKind.Array => ElementType!.SameAs(other.ElementType!),
            _ => false,
        };
    }

    public override string ToString() => Name;
}
