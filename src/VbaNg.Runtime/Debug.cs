namespace VbaNg.Runtime;

/// <summary>The VBA <c>Debug</c> object. Generated code calls <see cref="Print(PrintList, bool)"/> for <c>Debug.Print</c>.</summary>
public static class Debug
{
    [ThreadStatic]
    private static string? pending;

    /// <summary>Prints one line to the host's debug output (the Immediate window in VBA).</summary>
    public static void Print(string text) => Host.Current.Print(text ?? string.Empty);

    /// <summary>
    /// Prints an output list. A list ending in ";" or "," leaves the line open, so the next
    /// Print continues it (MS-VBAL 5.4.5.6: the line terminator is written only when the list
    /// does not end in a separator).
    /// </summary>
    public static void Print(PrintList list, bool endLine)
    {
        ArgumentNullException.ThrowIfNull(list);
        var text = (pending ?? string.Empty) + list;
        if (endLine)
        {
            pending = null;
            Host.Current.Print(text);
        }
        else
        {
            pending = text;
        }
    }

    /// <summary>
    /// Ends the output of a run: a line a Print left open is written, as the Immediate window
    /// shows it, since a host's captured output has no open line to continue. Hosts call it when a
    /// procedure or a test they ran returns.
    /// </summary>
    internal static void Flush()
    {
        if (pending is { } text)
        {
            pending = null;
            Host.Current.Print(text);
        }
    }

    /// <summary>Debug.Assert: breaks into an attached debugger when the condition is False; no effect otherwise.</summary>
    public static void Assert(in Variant condition)
    {
        if (!Coerce.ToCondition(condition) && System.Diagnostics.Debugger.IsAttached)
        {
            System.Diagnostics.Debugger.Break();
        }
    }

    /// <summary>The Stop statement (MS-VBAL 5.4.2.16): a breakpoint when a debugger is attached; otherwise nothing.</summary>
    public static void Stop()
    {
        if (System.Diagnostics.Debugger.IsAttached)
        {
            System.Diagnostics.Debugger.Break();
        }
    }
}
