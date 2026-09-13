using System.Globalization;

using VbaNg.Runtime;
using VbaNg.Runtime.TypeLibraries;

namespace VbaNg.Compiler.Binding;

/// <summary>A name a referenced library puts in scope unqualified: an enum or module constant, or a member of an application object.</summary>
public abstract class LibraryGlobal(string name) : Symbol(name);

/// <summary>A constant of a referenced library: an enum member such as xlUp, or a module constant.</summary>
public sealed class ComConstantSymbol(string name, Variant value, VbaType type) : LibraryGlobal(name)
{
    public Variant Value { get; } = value;

    public VbaType Type { get; } = type;
}

/// <summary>A member of a library's application object (a coclass marked appobject): Excel's ActiveSheet, Range, Worksheets.</summary>
public sealed class ComGlobalSymbol(string name, ComLibrary library, ComType appObject, ComType members) : LibraryGlobal(name)
{
    public ComLibrary Library { get; } = library;

    /// <summary>The appobject coclass, whose instance the host supplies.</summary>
    public ComType AppObject { get; } = appObject;

    /// <summary>The interface the member is looked up on, with the dispids the instance answers to.</summary>
    public ComType Members { get; } = members;
}

/// <summary>A COM type used as an expression: the operand of New or TypeOf.</summary>
public sealed class ComTypeSymbol(string name, VbaType type) : LibraryGlobal(name)
{
    public VbaType Type { get; } = type;
}

/// <summary>A member of a document module's own object, used unqualified: Range or Name inside a sheet module means Me.Range, Me.Name.</summary>
public sealed class DocumentMemberSymbol(string name, VariableSymbol me) : Symbol(name)
{
    public VariableSymbol Me { get; } = me;
}

/// <summary>
/// The type libraries a project references, in manifest order (ARCHITECTURE.md section 6):
/// resolves type names, member types, and the names each library puts in scope. Library types
/// map to <see cref="VbaType"/> instances that are shared per qualified name, so the binder can
/// compare them by reference as it does its own types.
/// </summary>
public sealed class ReferencedLibraries
{
    private readonly IReadOnlyList<ComLibrary> libraries;
    private readonly Dictionary<string, VbaType> types = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, LibraryGlobal>? globals;

    public ReferencedLibraries(IReadOnlyList<ComLibrary> libraries)
    {
        ArgumentNullException.ThrowIfNull(libraries);
        this.libraries = libraries;
    }

    public static ReferencedLibraries None { get; } = new([]);

    public IReadOnlyList<ComLibrary> Libraries => libraries;

    public ComLibrary? FindLibrary(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return libraries.FirstOrDefault(l => l.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>A type by optional library qualifier and name, as a declaration names it; null when no library has it.</summary>
    public VbaType? ResolveType(string? qualifier, string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        foreach (var library in libraries)
        {
            if (qualifier is not null && !library.Name.Equals(qualifier, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var type = library.FindType(name);
            if (type is not null && TypeFor(library, type) is { } resolved)
            {
                return resolved;
            }
        }

        return null;
    }

    /// <summary>The names in scope unqualified from every referenced library, first library first (MS-VBAL 5.6.10, the referenced projects and libraries step).</summary>
    public LibraryGlobal? FindGlobal(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        globals ??= CollectGlobals();
        return globals.GetValueOrDefault(name);
    }

    /// <summary>The members of an interface and the interfaces it inherits, most derived first.</summary>
    public IEnumerable<(ComType Owner, ComMember Member)> Members(ComType type, string name)
    {
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(name);
        var visited = new HashSet<ComType>();
        var pending = new Stack<ComType>();
        pending.Push(type);
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            if (!visited.Add(current))
            {
                continue;
            }

            foreach (var member in current.FindMembers(name))
            {
                yield return (current, member);
            }

            if (current.Kind is ComTypeKind.Interface or ComTypeKind.Dispatch)
            {
                foreach (var implemented in current.Implements)
                {
                    if (FindQualified(implemented.TypeName) is { } parent && parent.Kind is ComTypeKind.Interface or ComTypeKind.Dispatch)
                    {
                        pending.Push(parent);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Whether an interface may carry members the library does not list (no TYPEFLAG_FNONEXTENSIBLE
    /// on it or on one it inherits): VBA binds an unknown member on such a type at run time, as it
    /// does for <c>MSForms.Control</c>, and reports one on Excel's interfaces at compile time
    /// (docs/vba-quirks.md, verified 2026-09-09).
    /// </summary>
    public bool IsExtensible(ComType type)
    {
        ArgumentNullException.ThrowIfNull(type);
        var visited = new HashSet<ComType>();
        var pending = new Stack<ComType>();
        pending.Push(type);
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            if (!visited.Add(current))
            {
                continue;
            }

            if (current.IsExtensible)
            {
                return true;
            }

            if (current.Kind is ComTypeKind.Interface or ComTypeKind.Dispatch)
            {
                foreach (var implemented in current.Implements)
                {
                    if (FindQualified(implemented.TypeName) is { } parent && parent.Kind is ComTypeKind.Interface or ComTypeKind.Dispatch)
                    {
                        pending.Push(parent);
                    }
                }
            }
        }

        return false;
    }

    /// <summary>The default member (DISPID_VALUE) of an interface, if any.</summary>
    public static ComMember? DefaultMember(ComType type)
    {
        ArgumentNullException.ThrowIfNull(type);
        return type.Members.FirstOrDefault(m => m.IsDefault && m.Kind is ComMemberKind.PropertyGet or ComMemberKind.Method);
    }

    /// <summary>The VbaType a member's type reference denotes: a scalar, a COM type, Long for enums, Variant when unknown.</summary>
    public VbaType TypeOf(ComTypeRef? reference)
    {
        if (reference is null)
        {
            return VbaType.Variant;
        }

        if (reference.IsArray)
        {
            return VbaType.Variant;
        }

        if (reference.IsUserDefined)
        {
            var target = FindQualified(reference.TypeName);
            if (target is null)
            {
                return VbaType.Object;
            }

            return target.Kind switch
            {
                ComTypeKind.Enum => VbaType.Long,
                ComTypeKind.Alias => TypeOf(target.AliasOf),
                ComTypeKind.Interface or ComTypeKind.Dispatch or ComTypeKind.CoClass => TypeFor(LibraryOf(reference.TypeName!), target) ?? VbaType.Object,
                _ => VbaType.Variant,
            };
        }

        return (VarType)reference.VarType switch
        {
            VarType.Integer => VbaType.Integer,
            VarType.Long => VbaType.Long,
            VarType.LongLong => VbaType.LongLong,
            VarType.Single => VbaType.Single,
            VarType.Double => VbaType.Double,
            VarType.Currency => VbaType.Currency,
            VarType.Date => VbaType.Date,
            VarType.String => VbaType.String,
            VarType.Boolean => VbaType.Boolean,
            VarType.Byte => VbaType.Byte,
            VarType.Object => VbaType.Object,
            VarType.DataObject => VbaType.Object,
            (VarType)22 => VbaType.Long,
            (VarType)23 => VbaType.Long,
            (VarType)16 => VbaType.Integer,
            (VarType)18 => VbaType.Long,
            (VarType)19 => VbaType.Long,
            _ => VbaType.Variant,
        };
    }

    /// <summary>The interface of a coclass or interface type: the coclass's default interface, the type itself otherwise.</summary>
    public ComType? InterfaceOf(ComType type)
    {
        ArgumentNullException.ThrowIfNull(type);
        return type.Kind == ComTypeKind.CoClass
            ? type.DefaultInterface is { } name ? FindQualified(name) : null
            : type;
    }

    public ComType? FindQualified(string? qualifiedName)
    {
        if (qualifiedName is null)
        {
            return null;
        }

        var dot = qualifiedName.IndexOf('.', StringComparison.Ordinal);
        if (dot < 0)
        {
            return libraries.Select(l => l.FindType(qualifiedName)).FirstOrDefault(t => t is not null);
        }

        return FindLibrary(qualifiedName[..dot])?.FindType(qualifiedName[(dot + 1)..]);
    }

    /// <summary>The Variant a library constant holds, from its VARENUM code and text.</summary>
    public static Variant ConstantValue(ComConstant constant)
    {
        ArgumentNullException.ThrowIfNull(constant);
        var culture = CultureInfo.InvariantCulture;
        return (VarType)constant.VarType switch
        {
            VarType.Integer => Variant.FromInt16(short.Parse(constant.Text, culture)),
            VarType.Long => Variant.FromInt32(int.Parse(constant.Text, culture)),
            VarType.LongLong => Variant.FromInt64(long.Parse(constant.Text, culture)),
            VarType.Single => Variant.FromSingle(float.Parse(constant.Text, culture)),
            VarType.Double => Variant.FromDouble(double.Parse(constant.Text, culture)),
            VarType.Currency => Variant.FromCurrency(Currency.FromDecimal(decimal.Parse(constant.Text, culture))),
            VarType.Boolean => Variant.FromBoolean(constant.Text == "True"),
            VarType.Byte => Variant.FromByte(byte.Parse(constant.Text, culture)),
            VarType.String => Variant.FromString(constant.Text),
            VarType.Null => Variant.Null,
            _ => Variant.Empty,
        };
    }

    private ComLibrary LibraryOf(string qualifiedName)
    {
        var dot = qualifiedName.IndexOf('.', StringComparison.Ordinal);
        return (dot < 0 ? null : FindLibrary(qualifiedName[..dot])) ?? libraries[0];
    }

    /// <summary>The shared VbaType of a coclass, interface, or alias; null for kinds that are not object types.</summary>
    private VbaType? TypeFor(ComLibrary library, ComType type)
    {
        switch (type.Kind)
        {
            case ComTypeKind.Enum:
                return VbaType.Long;
            case ComTypeKind.Alias:
                return TypeOf(type.AliasOf);
            case ComTypeKind.CoClass:
            case ComTypeKind.Interface:
            case ComTypeKind.Dispatch:
                {
                    var key = library.Name + "." + type.Name;
                    if (!types.TryGetValue(key, out var existing))
                    {
                        var members = InterfaceOf(type) ?? type;
                        existing = VbaType.ForCom(library, type, members);
                        types[key] = existing;
                    }

                    return existing;
                }

            default:
                return null;
        }
    }

    private Dictionary<string, LibraryGlobal> CollectGlobals()
    {
        var result = new Dictionary<string, LibraryGlobal>(StringComparer.OrdinalIgnoreCase);
        foreach (var library in libraries)
        {
            foreach (var type in library.Types)
            {
                // Hidden types still put their names in scope: Excel's Global coclass and Constants module are both hidden.
                if (type.IsRestricted)
                {
                    continue;
                }

                switch (type.Kind)
                {
                    case ComTypeKind.Enum:
                    case ComTypeKind.Module:
                        foreach (var member in type.Members)
                        {
                            if (member.Kind == ComMemberKind.Constant && member.Value is not null && !member.IsHidden)
                            {
                                var value = ConstantValue(member.Value);
                                result.TryAdd(member.Name, new ComConstantSymbol(member.Name, value, type.Kind == ComTypeKind.Enum ? VbaType.Long : TypeOf(member.Type)));
                            }
                        }

                        break;

                    case ComTypeKind.CoClass:
                        if (type.IsAppObject && InterfaceOf(type) is { } appInterface)
                        {
                            foreach (var (owner, member) in AllMembers(appInterface))
                            {
                                if (!member.IsHidden && !member.IsRestricted && !member.IsNewEnum)
                                {
                                    result.TryAdd(member.Name, new ComGlobalSymbol(member.Name, library, type, owner));
                                }
                            }
                        }

                        break;
                }
            }
        }

        return result;
    }

    private IEnumerable<(ComType Owner, ComMember Member)> AllMembers(ComType type)
    {
        var visited = new HashSet<ComType>();
        var pending = new Stack<ComType>();
        pending.Push(type);
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            if (!visited.Add(current))
            {
                continue;
            }

            foreach (var member in current.Members)
            {
                yield return (current, member);
            }

            foreach (var implemented in current.Implements)
            {
                if (FindQualified(implemented.TypeName) is { } parent && parent.Kind is ComTypeKind.Interface or ComTypeKind.Dispatch)
                {
                    pending.Push(parent);
                }
            }
        }
    }
}
