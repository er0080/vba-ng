using System.Diagnostics;
using System.Windows.Forms;

using ExcelDna.Integration;

using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

/// <summary>
/// Host services inside Excel. <c>Debug.Print</c> goes to the capture of the current host
/// command, if one is running, and to the attached debugger's console so an F5 session in
/// VS Code sees the output (ARCHITECTURE.md section 9). Dialogs are real when code runs from
/// inside Excel (buttons, events) and, per D8, non-interactive during a host command unless the
/// command asked for the user interface.
/// </summary>
internal sealed class ExcelHostServices : IHostServices
{
    private const int RecentLines = 50;
    private const string DefaultTitle = "Microsoft Excel";

    private readonly Queue<string> recent = new();
    private CapturingHostServices? capture;
    private bool interactiveCapture;
    private string? runningProject;

    private ExcelHostServices()
    {
    }

    public static ExcelHostServices Instance { get; } = new();

    /// <summary>The last lines printed, for <c>vbang status</c>: what the lifecycle and user code said most recently.</summary>
    public IReadOnlyCollection<string> Recent => recent;

    /// <summary>True when dialogs may open: outside a host command, or inside one that asked for them.</summary>
    public bool Interactive => capture is null || interactiveCapture;

    /// <summary>Sends what the project prints, and the errors it leaves unhandled, to its <c>out/output.log</c> until the returned scope is disposed (ROADMAP.md WP4).</summary>
    public IDisposable Running(string? projectDir)
    {
        var previous = runningProject;
        runningProject = string.IsNullOrWhiteSpace(projectDir) ? previous : projectDir;
        return new RunningScope(this, previous);
    }

    public void Print(string text)
    {
        capture?.Print(text);
        if (runningProject is not null)
        {
            ProjectLog.Append(runningProject, text);
        }

        recent.Enqueue(text);
        while (recent.Count > RecentLines)
        {
            recent.Dequeue();
        }

        if (Debugger.IsLogging())
        {
            Debugger.Log(0, "vbang", text + Environment.NewLine);
        }
    }

    public int MsgBox(string prompt, int buttons, string? title)
    {
        if (!Interactive)
        {
            Print(string.IsNullOrEmpty(title) ? "MsgBox: " + prompt : "MsgBox (" + title + "): " + prompt);
            return 1;
        }

        var result = MessageBox.Show(
            new ExcelWindow(),
            prompt,
            string.IsNullOrEmpty(title) ? DefaultTitle : title,
            (buttons & 0xF) switch
            {
                1 => MessageBoxButtons.OKCancel,
                2 => MessageBoxButtons.AbortRetryIgnore,
                3 => MessageBoxButtons.YesNoCancel,
                4 => MessageBoxButtons.YesNo,
                5 => MessageBoxButtons.RetryCancel,
                _ => MessageBoxButtons.OK,
            },
            (buttons & 0xF0) switch
            {
                16 => MessageBoxIcon.Error,
                32 => MessageBoxIcon.Question,
                48 => MessageBoxIcon.Warning,
                64 => MessageBoxIcon.Information,
                _ => MessageBoxIcon.None,
            },
            (buttons & 0xF00) switch
            {
                256 => MessageBoxDefaultButton.Button2,
                512 => MessageBoxDefaultButton.Button3,
                _ => MessageBoxDefaultButton.Button1,
            });
        return result switch
        {
            DialogResult.OK => 1,
            DialogResult.Cancel => 2,
            DialogResult.Abort => 3,
            DialogResult.Retry => 4,
            DialogResult.Ignore => 5,
            DialogResult.Yes => 6,
            DialogResult.No => 7,
            _ => 1,
        };
    }

    public string InputBox(string prompt, string? title, string defaultText)
    {
        if (!Interactive)
        {
            Print(string.IsNullOrEmpty(title) ? "InputBox: " + prompt : "InputBox (" + title + "): " + prompt);
            return defaultText;
        }

        using var dialog = new InputBoxForm(prompt, string.IsNullOrEmpty(title) ? DefaultTitle : title, defaultText);
        return dialog.ShowDialog(new ExcelWindow()) == DialogResult.OK ? dialog.Value : string.Empty;
    }

    public void Beep()
    {
        if (Interactive)
        {
            System.Media.SystemSounds.Beep.Play();
        }
        else
        {
            Print("Beep");
        }
    }

    public void SendKeys(string keys, bool wait)
    {
        if (!Interactive)
        {
            Print("SendKeys: " + keys);
            return;
        }

        if (wait)
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
        }
        else
        {
            System.Windows.Forms.SendKeys.Send(keys);
        }
    }

    /// <summary>
    /// An error a macro left unhandled (ROADMAP.md WP4): the line reaches the output and the project's log, and a
    /// command or an event Excel ran shows VBA's run-time error dialog when dialogs may open. Its End button ends
    /// the macro, which has already stopped. A UDF shows #VALUE! instead, as in VBA, so it passes no dialog.
    /// </summary>
    public void ReportUnhandled(string where, Exception exception, bool dialog)
    {
        ArgumentNullException.ThrowIfNull(exception);
        var error = UnhandledError.Of(exception);
        Print(error is null
            ? $"vba-ng could not run {where}: {exception.GetType().Name}: {exception.Message}"
            : string.Create(System.Globalization.CultureInfo.InvariantCulture, $"Run-time error in {where}: '{error.Number}': {error.Description}"));
        if (!dialog || !Interactive)
        {
            return;
        }

        var text = error is null ? $"vba-ng could not run {where}:\n\n{exception.GetType().Name}: {exception.Message}" : UnhandledError.DialogText(error);
        using var form = new RunTimeErrorForm(text);
        form.ShowDialog(new ExcelWindow());
    }

    /// <summary>A message from vba-ng itself: printed, and shown in a dialog when dialogs may open.</summary>
    public void Notify(string message)
    {
        Print(message);
        if (Interactive)
        {
            MessageBox.Show(new ExcelWindow(), message, "vba-ng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    /// <summary>Routes output to <paramref name="target"/> until the returned scope is disposed; dialogs are suppressed unless <paramref name="interactive"/>.</summary>
    public IDisposable Capture(CapturingHostServices target, bool interactive = false)
    {
        ArgumentNullException.ThrowIfNull(target);
        var previous = (capture, interactiveCapture);
        capture = target;
        interactiveCapture = interactive;
        return new CaptureScope(this, previous.capture, previous.interactiveCapture);
    }

    private sealed class CaptureScope(ExcelHostServices owner, CapturingHostServices? previous, bool previousInteractive) : IDisposable
    {
        public void Dispose()
        {
            owner.capture = previous;
            owner.interactiveCapture = previousInteractive;
        }
    }

    private sealed class RunningScope(ExcelHostServices owner, string? previous) : IDisposable
    {
        public void Dispose() => owner.runningProject = previous;
    }

    /// <summary>VBA's run-time error dialog: its title, the number and description, and its four buttons, of which End is the one that works (Excel probe 2026-09-11).</summary>
    private sealed class RunTimeErrorForm : Form
    {
        private static readonly string[] Captions = ["&Continue", "&End", "&Debug", "&Help"];

        public RunTimeErrorForm(string text)
        {
            Text = UnhandledError.DialogTitle;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new System.Drawing.Size(420, 170);

            var label = new Label { Text = text, Left = 16, Top = 16, Width = 388, Height = 100, UseMnemonic = false };
            var buttons = Captions.Select((caption, i) => new Button { Text = caption, Left = 16 + (i * 98), Top = 128, Width = 90, Enabled = caption == "&End" }).ToArray();
            buttons[1].DialogResult = DialogResult.OK;
            Controls.Add(label);
            Controls.AddRange(buttons);
            AcceptButton = buttons[1];
            CancelButton = buttons[1];
        }
    }

    /// <summary>Excel's main window as the owner of dialogs, so they stay in front of it.</summary>
    private sealed class ExcelWindow : IWin32Window
    {
        public nint Handle => ExcelDnaUtil.WindowHandle;
    }

    /// <summary>The InputBox dialog: a prompt, a text box with the default, OK and Cancel.</summary>
    private sealed class InputBoxForm : Form
    {
        private readonly TextBox input;

        public InputBoxForm(string prompt, string title, string defaultText)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new System.Drawing.Size(400, 130);

            var label = new Label { Text = prompt, Left = 12, Top = 12, Width = 280, Height = 70 };
            input = new TextBox { Text = defaultText, Left = 12, Top = 90, Width = 376 };
            var ok = new Button { Text = "OK", Left = 300, Top = 12, Width = 88, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 300, Top = 44, Width = 88, DialogResult = DialogResult.Cancel };
            Controls.AddRange([label, input, ok, cancel]);
            AcceptButton = ok;
            CancelButton = cancel;
            input.SelectAll();
        }

        public string Value => input.Text;
    }
}
