namespace VbaNg.Runtime;

/// <summary>
/// The host services currently in effect for compiled code. Generated code reaches the host only
/// through this class, so every host (add-in, tests) shares one runtime instance per process
/// (see <see cref="Hosting.ProjectLoadContext"/>).
/// </summary>
public static class Host
{
    private static IHostServices current = NullHostServices.Instance;

    /// <summary>The active host services. Never null; defaults to a host that discards output.</summary>
    public static IHostServices Current
    {
        get => current;
        set => current = value ?? throw new ArgumentNullException(nameof(value));
    }

    private sealed class NullHostServices : IHostServices
    {
        public static readonly NullHostServices Instance = new();

        public void Print(string text)
        {
        }
    }
}
