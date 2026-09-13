using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace VbaNg.Golden.Regen;

/// <summary>Late-bound COM calls through <c>IDispatch</c>, with the reflection wrapper unwrapped.</summary>
internal static class Com
{
    public static object? Get(object target, string name) => Invoke(target, name, BindingFlags.GetProperty, null);

    public static void Set(object target, string name, object? value) => Invoke(target, name, BindingFlags.SetProperty, [value]);

    public static object? Call(object target, string name, params object?[] arguments) => Invoke(target, name, BindingFlags.InvokeMethod, arguments);

    public static void Release(object? target)
    {
        if (target is not null && Marshal.IsComObject(target))
        {
            Marshal.FinalReleaseComObject(target);
        }
    }

    private static object? Invoke(object target, string name, BindingFlags flags, object?[]? arguments)
    {
        try
        {
            return target.GetType().InvokeMember(name, flags, binder: null, target, arguments, CultureInfo.InvariantCulture);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }
}
