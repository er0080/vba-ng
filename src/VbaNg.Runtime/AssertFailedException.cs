namespace VbaNg.Runtime;

/// <summary>
/// A failed <c>Assert</c> call in a test procedure (ARCHITECTURE.md section 8). Deliberately not a
/// <see cref="VbaException"/>: <see cref="VbaException.From"/> does not map it, so no
/// <c>On Error</c> handler in the test can swallow it, and a test that enables
/// <c>On Error Resume Next</c> to check <c>Err.Number</c> still fails when its assertion fails.
/// </summary>
public sealed class AssertFailedException : Exception
{
    public AssertFailedException()
    {
    }

    public AssertFailedException(string message)
        : base(message)
    {
    }

    public AssertFailedException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
