using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// What the host shows for an error a macro leaves unhandled (ROADMAP.md WP4): VBA's run-time
/// error dialog reads "Run-time error '5':", a blank line, and the description (Excel probe
/// 2026-09-11, docs/vba-quirks.md), and an error that escaped the procedure Application.Run
/// called (ARCHITECTURE.md D23) is the callee's error.
/// </summary>
public sealed class UnhandledErrorTests
{
    [Fact]
    public void DialogText_IsVbasRunTimeErrorDialog()
    {
        Assert.Equal("Run-time error '5':\n\ncustom", UnhandledError.DialogText(new VbaException(5, "custom")));
    }

    [Fact]
    public void Of_AnErrorEscapingApplicationRun_IsTheCalleesError()
    {
        var error = new VbaException(1004, "inner");

        Assert.Same(error, UnhandledError.Of(new UnhandledErrorException(error, error)));
        Assert.Same(error, UnhandledError.Of(error));
        Assert.Null(UnhandledError.Of(new InvalidOperationException("a defect of the host")));
    }
}
