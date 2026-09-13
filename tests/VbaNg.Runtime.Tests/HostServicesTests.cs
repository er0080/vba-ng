using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The non-interactive host services (ARCHITECTURE.md D8): MsgBox prints its text and answers
/// vbOK, InputBox answers its default, and the operating-system members raise the numbers VBA
/// does (Interaction golden). Tests in this class share <see cref="Host.Current"/> and run one
/// at a time, as xunit runs the tests of one class.
/// </summary>
public sealed class HostServicesTests
{
    [Fact]
    public void MsgBox_NonInteractive_PrintsAndAnswersOk()
    {
        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = capture;
        try
        {
            var plain = Library.Interaction.MsgBox("Hello", Variant.Missing, Variant.Missing);
            var titled = Library.Interaction.MsgBox("Sure?", 4 + 32, "Ask");

            Assert.Equal(Variant.FromInt32(1), plain);
            Assert.Equal(Variant.FromInt32(1), titled);
            Assert.Equal(["MsgBox: Hello", "MsgBox (Ask): Sure?"], capture.Lines);
        }
        finally
        {
            Host.Current = previous;
        }
    }

    [Fact]
    public void InputBox_NonInteractive_AnswersTheDefault()
    {
        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = capture;
        try
        {
            Assert.Equal("42", Library.Interaction.InputBox("Value?", Variant.Missing, "42").ToString());
            Assert.Equal(string.Empty, Library.Interaction.InputBox("Value?", "T", Variant.Missing).ToString());
            Assert.Equal(["InputBox: Value?", "InputBox (T): Value?"], capture.Lines);
        }
        finally
        {
            Host.Current = previous;
        }
    }

    [Fact]
    public void BeepAndSendKeys_NonInteractive_LeaveALine()
    {
        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = capture;
        try
        {
            Library.Interaction.Beep();
            Library.Interaction.SendKeys("{ENTER}", Variant.Missing);

            Assert.Equal(["Beep", "SendKeys: {ENTER}"], capture.Lines);
        }
        finally
        {
            Host.Current = previous;
        }
    }

    [Fact]
    public void Shell_RaisesVbaErrors()
    {
        Assert.Equal(5, Assert.Throws<VbaException>(() => Library.Interaction.Shell(string.Empty, Variant.Missing)).Number);
        Assert.Equal(53, Assert.Throws<VbaException>(() => Library.Interaction.Shell("vbang-tests-does-not-exist.exe", Variant.Missing)).Number);
    }

    [Fact]
    public void AppActivate_UnknownTitle_Raises5()
    {
        Assert.Equal(5, Assert.Throws<VbaException>(() => Library.Interaction.AppActivate("VbaNg Tests Window That Does Not Exist", Variant.Missing)).Number);
    }
}
