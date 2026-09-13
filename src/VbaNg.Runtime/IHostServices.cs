namespace VbaNg.Runtime;

/// <summary>
/// Services the host provides to compiled VBA code (ARCHITECTURE.md section 5, "Host services").
/// The Excel add-in supplies the interactive implementation; the CLI and test harnesses supply
/// non-interactive ones (ARCHITECTURE.md D8). The defaults here are the non-interactive
/// behavior: <c>MsgBox</c> prints its text and answers vbOK, <c>InputBox</c> answers its
/// default, <c>Beep</c> and <c>SendKeys</c> do nothing but leave a line.
/// </summary>
public interface IHostServices
{
    /// <summary>Receives one line of <c>Debug.Print</c> output.</summary>
    void Print(string text);

    /// <summary>MsgBox(prompt, buttons, title): the button the user chose as a VbMsgBoxResult (vbOK is 1).</summary>
    int MsgBox(string prompt, int buttons, string? title)
    {
        ArgumentNullException.ThrowIfNull(prompt);
        Print(string.IsNullOrEmpty(title) ? "MsgBox: " + prompt : "MsgBox (" + title + "): " + prompt);
        return 1;
    }

    /// <summary>InputBox(prompt, title, default): what the user typed, or an empty string when they cancelled.</summary>
    string InputBox(string prompt, string? title, string defaultText)
    {
        ArgumentNullException.ThrowIfNull(prompt);
        ArgumentNullException.ThrowIfNull(defaultText);
        Print(string.IsNullOrEmpty(title) ? "InputBox: " + prompt : "InputBox (" + title + "): " + prompt);
        return defaultText;
    }

    void Beep() => Print("Beep");

    void SendKeys(string keys, bool wait)
    {
        ArgumentNullException.ThrowIfNull(keys);
        Print("SendKeys: " + keys);
    }
}
