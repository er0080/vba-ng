using System.Globalization;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// An error a macro leaves unhandled, as the host reports it (ROADMAP.md WP4; ARCHITECTURE.md
/// section 6): VBA's run-time error dialog when the run is interactive, and the same text in the
/// project's log always.
/// </summary>
public static class UnhandledError
{
    /// <summary>The title of VBA's run-time error dialog.</summary>
    public const string DialogTitle = "Microsoft Visual Basic";

    /// <summary>
    /// The VBA error behind an exception that reached the host: the callee's own for one that
    /// escaped the procedure Application.Run called (ARCHITECTURE.md D23); null for a defect of
    /// the host rather than an error of the macro.
    /// </summary>
    public static VbaException? Of(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        return exception is UnhandledErrorException { Error: { } error } ? error : VbaException.From(exception);
    }

    /// <summary>The text of VBA's run-time error dialog: "Run-time error '5':", a blank line, and the description (Excel probe 2026-09-11, docs/vba-quirks.md).</summary>
    public static string DialogText(VbaException error)
    {
        ArgumentNullException.ThrowIfNull(error);
        return string.Create(CultureInfo.InvariantCulture, $"Run-time error '{error.Number}':\n\n{error.Description}");
    }
}
