namespace VbaNg.Runtime.Hosting;

/// <summary>Host services that collect <c>Debug.Print</c> output in memory. Used by the add-in's
/// host commands and by tests.</summary>
public sealed class CapturingHostServices : IHostServices
{
    private readonly List<string> lines = [];

    public IReadOnlyList<string> Lines => lines;

    /// <summary>All captured lines joined with newlines.</summary>
    public string Text => string.Join(Environment.NewLine, lines);

    public void Print(string text) => lines.Add(text);
}
