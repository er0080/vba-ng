using VbaNg.Runtime.Library;

namespace VbaNg.Runtime;

/// <summary>
/// The Err object of the running code (MS-VBAL 6.1.3.1). VBA keeps one per thread of execution;
/// generated code reads and raises through <see cref="Current"/>, and compiled error handlers set
/// it from the <see cref="VbaException"/> they catch.
/// </summary>
public static class Err
{
    [ThreadStatic]
    private static ErrObject? current;

    public static ErrObject Current => current ??= new ErrObject();
}
