using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaNg.Runtime.TypeLibraries;

/// <summary>The kind of a type library type (TYPEKIND).</summary>
public enum ComTypeKind
{
    Enum,
    Record,
    Module,
    Interface,
    Dispatch,
    CoClass,
    Alias,
    Union,
}

/// <summary>What a member is: a method, one of the three property accessors, an enum or module constant, or a record field.</summary>
public enum ComMemberKind
{
    Method,
    PropertyGet,
    PropertyPut,
    PropertyPutRef,
    Constant,
    Field,
}

/// <summary>
/// A type reference inside a type library: the VARENUM code, the qualified name of a
/// user-defined type (<c>Library.Type</c>) for VT_USERDEFINED, and whether the reference is a
/// SAFEARRAY or reached through a pointer (an object reference for interfaces, ByRef for values).
/// </summary>
public sealed record ComTypeRef(int VarType, string? TypeName, bool IsArray, bool IsPointer)
{
    public const int UserDefined = 29;

    public bool IsUserDefined => VarType == UserDefined;
}

/// <summary>A constant value in the library, kept as its VARENUM code and invariant text so the model stays plain JSON.</summary>
public sealed record ComConstant(int VarType, string Text);

public sealed record ComParameter(
    string Name,
    ComTypeRef Type,
    bool IsOptional,
    bool IsByRef,
    bool IsOut,
    bool IsParamArray,
    ComConstant? Default);

/// <summary>A member of a type: <see cref="Type"/> is the return or property type, null for a Sub-like method.</summary>
public sealed record ComMember(
    string Name,
    int DispId,
    ComMemberKind Kind,
    ComTypeRef? Type,
    IReadOnlyList<ComParameter> Parameters,
    bool IsHidden,
    bool IsRestricted,
    ComConstant? Value)
{
    /// <summary>DISPID_VALUE: the member <c>obj(args)</c> reaches.</summary>
    public bool IsDefault => DispId == 0;

    /// <summary>DISPID_NEWENUM: the enumerator behind For Each.</summary>
    public bool IsNewEnum => DispId == -4;
}

/// <summary>An interface a coclass implements, with the IMPLTYPEFLAGS that mark the default and the event source.</summary>
public sealed record ComImplemented(string TypeName, bool IsDefault, bool IsSource, bool IsRestricted);

/// <summary>
/// A type of a type library (ARCHITECTURE.md section 6, "Type library reader"). A coclass lists
/// what it implements, an interface lists its members and names the interface it inherits in
/// <see cref="Implements"/>, an enum or module carries constants, an alias points at its target.
/// </summary>
public sealed record ComType(
    string Name,
    ComTypeKind Kind,
    Guid Guid,
    IReadOnlyList<ComMember> Members,
    IReadOnlyList<ComImplemented> Implements,
    ComTypeRef? AliasOf,
    bool IsHidden,
    bool IsRestricted,
    bool IsAppObject,
    bool IsDual,
    bool IsExtensible = false)
{
    /// <summary>The default interface of a coclass, by qualified name; null for other kinds.</summary>
    public string? DefaultInterface => Implements.FirstOrDefault(i => i.IsDefault && !i.IsSource)?.TypeName;

    /// <summary>The default source (event) interface of a coclass, by qualified name.</summary>
    public string? DefaultSource => Implements.FirstOrDefault(i => i.IsDefault && i.IsSource)?.TypeName;

    public ComMember? FindMember(string name, ComMemberKind kind) =>
        Members.FirstOrDefault(m => m.Kind == kind && m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    public IEnumerable<ComMember> FindMembers(string name) =>
        Members.Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// The compact symbol model of one type library, as the binder, the code generator, and the
/// cache see it (ARCHITECTURE.md section 6). The Interop project builds it from
/// ITypeLib/ITypeInfo; hosts keep it as JSON per library version under the user profile so a
/// build never reads the registry twice for the same library.
/// </summary>
public sealed record ComLibrary(string Name, Guid Guid, int MajorVersion, int MinorVersion, IReadOnlyList<ComType> Types)
{
    /// <summary>The shape of the model this code writes; a cached file of an older shape is read again from the type library (<see cref="Format"/>).</summary>
    public const int CurrentFormat = 2;

    /// <summary>The shape the file was written in: <see cref="CurrentFormat"/> for a model this code produced, 0 for a file that predates the field.</summary>
    public int Format { get; init; }

    private static readonly JsonSerializerOptions Options = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    /// <summary>A type by name, case-insensitively, as VBA resolves <c>Library.Type</c>.</summary>
    public ComType? FindType(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return Types.FirstOrDefault(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The folder hosts cache libraries in: <c>%LOCALAPPDATA%\vbang\typelibs</c>.</summary>
    public static string CacheDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbang", "typelibs");

    /// <summary>The cache file of a library version: <c>{guid}-{major}.{minor}.json</c>.</summary>
    public static string CachePath(Guid guid, int majorVersion, int minorVersion) =>
        Path.Combine(CacheDirectory, string.Create(System.Globalization.CultureInfo.InvariantCulture, $"{guid:D}-{majorVersion}.{minorVersion}.json"));

    public string ToJson() => JsonSerializer.Serialize(this, Options);

    public static ComLibrary FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        return JsonSerializer.Deserialize<ComLibrary>(json, Options) ?? throw new JsonException("The type library model is empty.");
    }

    public void Save(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, ToJson());
    }

    public static ComLibrary Load(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return FromJson(File.ReadAllText(path));
    }
}
