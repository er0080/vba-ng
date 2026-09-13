using System.Reflection;

namespace VbaNg.Runtime;

/// <summary>
/// The VBA name of a generated member whose C# name had to differ from it: a Sub named like its
/// module, which C# forbids as a member name (CS0542) while VBA allows it. Hosts resolve
/// procedures by VBA name through <see cref="VbaNames.Of(MethodInfo)"/>.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class VbaNameAttribute(string name) : Attribute
{
    public string Name { get; } = name ?? throw new ArgumentNullException(nameof(name));
}

/// <summary>Resolves the VBA name of generated members.</summary>
public static class VbaNames
{
    /// <summary>The procedure's VBA name: the attribute's when the C# name had to differ, otherwise the method name.</summary>
    public static string Of(MethodInfo method)
    {
        ArgumentNullException.ThrowIfNull(method);
        return method.GetCustomAttribute<VbaNameAttribute>()?.Name ?? method.Name;
    }
}
