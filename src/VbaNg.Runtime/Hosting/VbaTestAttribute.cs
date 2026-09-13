namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Marks the compiled form of a test procedure: a Public Sub without parameters carrying a
/// <c>'@Test</c> comment annotation (ARCHITECTURE.md section 8). <see cref="TestRunner"/> runs
/// every method that carries it.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class VbaTestAttribute : Attribute
{
}
