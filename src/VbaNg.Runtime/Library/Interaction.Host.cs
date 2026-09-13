using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The host-bound and OS-bound members of the Interaction module (MS-VBAL 6.1.2.6): dialogs go
/// through <see cref="Host.Current"/> (ARCHITECTURE.md D8), <c>Shell</c> and <c>AppActivate</c>
/// work on the operating system directly with the error numbers VBA raises (Interaction golden).
/// </summary>
public static partial class Interaction
{
    private const int MaxWindowTitle = 512;

    /// <summary>MsgBox(prompt, [buttons], [title], [helpfile], [context]): vbOK (1) unless the host asks the user.</summary>
    public static Variant MsgBox(in Variant prompt, in Variant buttons, in Variant title) =>
        Variant.FromInt32(Host.Current.MsgBox(Coerce.ToString(prompt), buttons.IsMissing ? 0 : Coerce.ToInt32(buttons), title.IsMissing ? null : Coerce.ToString(title)));

    /// <summary>InputBox(prompt, [title], [default], [xpos], [ypos], [helpfile], [context]): the default unless the host asks the user.</summary>
    public static VbaString InputBox(in Variant prompt, in Variant title, in Variant defaultText) => VbaString.Temporary(
        Host.Current.InputBox(Coerce.ToString(prompt), title.IsMissing ? null : Coerce.ToString(title), defaultText.IsMissing ? string.Empty : Coerce.ToString(defaultText)));

    public static void Beep() => Host.Current.Beep();

    /// <summary>SendKeys(keys, [wait]).</summary>
    public static void SendKeys(in Variant keys, in Variant wait) =>
        Host.Current.SendKeys(Coerce.ToString(keys), !wait.IsMissing && Coerce.ToBoolean(wait));

    /// <summary>
    /// Shell(pathname, [windowstyle]): starts a program and returns its task id (the process id)
    /// as a Double; an empty command raises 5 and one that cannot be started raises 53, as VBA
    /// does (Interaction golden). The window style follows VbAppWinStyle: vbHide 0,
    /// vbNormalFocus 1, vbMinimizedFocus 2 (the default), vbMaximizedFocus 3, vbNormalNoFocus 4,
    /// vbMinimizedNoFocus 6.
    /// </summary>
    public static double Shell(in Variant pathName, in Variant windowStyle)
    {
        var command = Coerce.ToString(pathName).Trim();
        if (command.Length == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var style = windowStyle.IsMissing ? 2 : Coerce.ToInt32(windowStyle);
        var (file, arguments) = SplitCommand(command);
        var info = new ProcessStartInfo(file, arguments)
        {
            UseShellExecute = true,
            WindowStyle = style switch
            {
                0 => ProcessWindowStyle.Hidden,
                2 or 6 => ProcessWindowStyle.Minimized,
                3 => ProcessWindowStyle.Maximized,
                _ => ProcessWindowStyle.Normal,
            },
        };

        try
        {
            using var process = Process.Start(info) ?? throw new VbaException(53);
            return process.Id;
        }
        catch (Win32Exception)
        {
            throw new VbaException(53);
        }
        catch (InvalidOperationException)
        {
            throw new VbaException(53);
        }
    }

    /// <summary>
    /// AppActivate(title, [wait]): brings to the front the top-level window whose title equals
    /// <paramref name="title"/>, else the first whose title starts with it, else the main window of
    /// the task id a Shell call returned; none raises 5 (Interaction golden).
    /// </summary>
    public static void AppActivate(in Variant title, in Variant wait)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        nint window = 0;
        if (title.IsNumeric || (title.IsString && double.TryParse(title.AsString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _)))
        {
            var processId = (int)Coerce.ToDouble(title);
            try
            {
                using var process = Process.GetProcessById(processId);
                window = process.MainWindowHandle;
            }
            catch (ArgumentException)
            {
                throw VbaErrors.InvalidProcedureCall();
            }
        }
        else
        {
            window = FindWindowByTitle(Coerce.ToString(title));
        }

        if (window == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (IsIconic(window))
        {
            ShowWindow(window, 9); // SW_RESTORE
        }

        SetForegroundWindow(window);
    }

    private static (string File, string Arguments) SplitCommand(string command)
    {
        if (command[0] == '"')
        {
            var close = command.IndexOf('"', 1);
            if (close > 0)
            {
                return (command[1..close], command[(close + 1)..].Trim());
            }
        }

        var space = command.IndexOf(' ', StringComparison.Ordinal);
        return space < 0 ? (command, string.Empty) : (command[..space], command[(space + 1)..].Trim());
    }

    private static unsafe nint FindWindowByTitle(string title)
    {
        nint exact = 0;
        nint prefix = 0;
        var buffer = new char[MaxWindowTitle];
        EnumWindows((window, _) =>
        {
            if (!IsWindowVisible(window))
            {
                return true;
            }

            int length;
            fixed (char* text = buffer)
            {
                length = GetWindowTextW(window, text, buffer.Length);
            }

            if (length <= 0)
            {
                return true;
            }

            var current = new string(buffer, 0, length);
            if (current.Equals(title, StringComparison.OrdinalIgnoreCase))
            {
                exact = window;
                return false;
            }

            if (prefix == 0 && current.StartsWith(title, StringComparison.OrdinalIgnoreCase))
            {
                prefix = window;
            }

            return true;
        }, 0);
        return exact != 0 ? exact : prefix;
    }

    private delegate bool EnumWindowsCallback(nint window, nint parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsCallback callback, nint parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern unsafe int GetWindowTextW(nint window, char* text, int maxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint window);

    [DllImport("user32.dll")]
    private static extern bool IsIconic(nint window);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(nint window, int command);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(nint window);
}
