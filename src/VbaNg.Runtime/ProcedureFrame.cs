namespace VbaNg.Runtime;

/// <summary>
/// The error-handling state of one procedure invocation (MS-VBAL 5.4.4): which handler is
/// enabled, whether it is active, and where Resume goes back to. Generated code creates a frame
/// for every procedure that uses On Error, Resume, GoTo, GoSub, or line labels, and runs such a
/// procedure as a dispatch loop over a program counter (ARCHITECTURE.md section 4, "Error
/// handling"). Procedures without those statements have no frame.
/// </summary>
public sealed class ProcedureFrame
{
    private const int NoHandler = -1;
    private const int ResumeNextHandler = -2;

    private int handler = NoHandler;
    private int retryPc = -1;
    private int nextPc = -1;
    private bool active;
    private Stack<int>? returns;

    /// <summary>True between an error reaching the handler and the Resume, Exit, or End that leaves it.</summary>
    public bool InHandler => active;

    /// <summary>On Error GoTo label: enables the handler at the given statement (MS-VBAL 5.4.4.1). Every On Error form resets Err (Errors golden).</summary>
    public void OnErrorGoTo(int handlerPc)
    {
        Err.Current.Clear();
        handler = handlerPc;
        active = false;
    }

    public void OnErrorResumeNext()
    {
        Err.Current.Clear();
        handler = ResumeNextHandler;
        active = false;
    }

    /// <summary>On Error GoTo 0: disables the handler.</summary>
    public void OnErrorGoToZero()
    {
        Err.Current.Clear();
        handler = NoHandler;
        active = false;
    }

    /// <summary>On Error GoTo -1: dismisses the active error; the handler stays enabled.</summary>
    public void OnErrorGoToMinusOne()
    {
        Err.Current.Clear();
        active = false;
    }

    /// <summary>
    /// Called from the dispatch loop's exception filter. Returns true, with the statement to run
    /// next, when this frame has an enabled handler that is not already active; the error then
    /// fills Err, and <paramref name="line"/> becomes Erl. Otherwise the exception propagates to
    /// the caller, as VBA passes an unhandled error up the call chain (MS-VBAL 5.4.4). The line is
    /// the compiler's: VBA resolves Erl from a per-procedure table of line numbers and code
    /// offsets, not from the numbered lines that ran, so a jump back to an unnumbered statement
    /// reports 0 after a numbered line ran (Errors golden; docs/vba-quirks.md, Erl).
    /// </summary>
    public bool TryHandle(Exception exception, int pc, int resumeNextPc, int line, out int next)
    {
        next = pc;
        var error = VbaException.From(exception);
        if (error is null || active || handler == NoHandler)
        {
            return false;
        }

        Err.Current.Set(error);
        Err.Current.Line = line;
        if (handler == ResumeNextHandler)
        {
            next = resumeNextPc;
            return true;
        }

        retryPc = pc;
        nextPc = resumeNextPc;
        active = true;
        next = handler;
        return true;
    }

    /// <summary>Resume: retries the statement that raised the error (MS-VBAL 5.4.4.2).</summary>
    public int Resume() => Leave(retryPc);

    /// <summary>Resume Next: continues after the statement that raised the error.</summary>
    public int ResumeNext() => Leave(nextPc);

    /// <summary>Resume label.</summary>
    public int ResumeAt(int pc) => Leave(pc);

    private int Leave(int pc)
    {
        if (!active)
        {
            throw new VbaException(VbaErrors.ResumeWithoutError);
        }

        active = false;
        Err.Current.Clear();
        return pc;
    }

    /// <summary>GoSub: remembers where Return continues (MS-VBAL 5.4.2.12).</summary>
    public void GoSub(int returnPc) => (returns ??= new Stack<int>()).Push(returnPc);

    /// <summary>Return: error 3 when no GoSub is pending.</summary>
    public int Return() => returns is { Count: > 0 } ? returns.Pop() : throw new VbaException(VbaErrors.ReturnWithoutGoSub);

    /// <summary>Leaving the procedure from inside its handler resets Err; an error skipped by Resume Next stays visible to the caller (Errors golden).</summary>
    public void Exit()
    {
        if (active)
        {
            Err.Current.Clear();
            active = false;
        }
    }
}

/// <summary>
/// Raised by the End statement (MS-VBAL 5.4.2.15): the host stops the running code and resets the
/// project (<see cref="Hosting.ProjectReset.AfterEnd"/>). Generated code raises it through
/// <see cref="Hosting.ProjectReset.Ending"/>, which names the project.
/// </summary>
public sealed class EndStatementException : Exception
{
    public EndStatementException()
        : base("End statement executed.")
    {
    }

    /// <summary>End in a module of <paramref name="project"/>.</summary>
    public EndStatementException(System.Reflection.Assembly project)
        : base("End statement executed.")
    {
        Project = project;
    }

    /// <summary>The project whose code ran End, which the host resets; null when raised without one.</summary>
    public System.Reflection.Assembly? Project { get; }

    public EndStatementException(string message)
        : base(message)
    {
    }

    public EndStatementException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// A run-time error no handler traps: one a procedure left unhandled where VBA reports it to the
/// user rather than to the caller, as in a procedure <c>Application.Run</c> called (docs/vba-quirks.md,
/// "Host, Application.Run"). <see cref="VbaException.From"/> does not map it, so it passes every
/// <c>On Error</c> handler and reaches the host as an unhandled error.
/// </summary>
public sealed class UnhandledErrorException : Exception
{
    public UnhandledErrorException()
        : base("Unhandled run-time error.")
    {
    }

    /// <summary>The run-time error <paramref name="error"/>, raised as <paramref name="innerException"/>.</summary>
    public UnhandledErrorException(VbaException error, Exception innerException)
        : base(Describe(error ?? throw new ArgumentNullException(nameof(error))), innerException)
    {
        Error = error;
    }

    public UnhandledErrorException(string message)
        : base(message)
    {
    }

    public UnhandledErrorException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>The error the procedure left unhandled.</summary>
    public VbaException? Error { get; }

    private static string Describe(VbaException error) =>
        string.Create(System.Globalization.CultureInfo.InvariantCulture, $"Run-time error '{error.Number}': {error.Description}");
}
