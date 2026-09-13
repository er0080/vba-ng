using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using ExcelDna.Integration;

namespace VbaNg.AddIn;

/// <summary>
/// Runs the startup procedure once Excel is actually ready: main window visible, splash screen
/// gone, <c>Application.Ready</c>, and a workbook open. Excel opens its command-line files,
/// including the add-in, while the splash screen (a top-level window of class MsoSplash) still
/// covers the main window, and <c>Application.Visible</c> turns true before the splash closes,
/// so a macro queued straight from AutoOpen would run behind the splash. Polls on the main
/// thread through QueueAsMacro and keeps a trace of the observed states for <c>vbang status</c>.
/// </summary>
internal static partial class StartupRunner
{
    private const int PollIntervalMilliseconds = 250;
    private const int MaxPolls = 240;
    private const string SplashWindowClass = "MsoSplash";

    [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", StringMarshalling = StringMarshalling.Utf16)]
    private static partial nint FindWindowEx(nint parent, nint childAfter, string? className, string? windowName);

    [LibraryImport("user32.dll")]
    private static partial uint GetWindowThreadProcessId(nint hwnd, out uint processId);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool IsWindowVisible(nint hwnd);

    /// <summary>True while this process shows a splash window.</summary>
    private static bool SplashIsShowing()
    {
        var ownPid = (uint)Environment.ProcessId;
        var splash = nint.Zero;
        while ((splash = FindWindowEx(nint.Zero, splash, SplashWindowClass, null)) != nint.Zero)
        {
            GetWindowThreadProcessId(splash, out var pid);
            if (pid == ownPid && IsWindowVisible(splash))
            {
                return true;
            }
        }

        return false;
    }

    public static void Start(string projectDir, string procedure)
    {
        var trace = new StateTrace();
        ExcelAsyncUtil.QueueAsMacro(() => Poll(projectDir, procedure, 1, Stopwatch.StartNew(), trace));
    }

    private static void Poll(string projectDir, string procedure, int attempt, Stopwatch elapsed, StateTrace trace)
    {
        var state = ExcelState.Probe();
        trace.Record(attempt, state);

        if (state.IsReady)
        {
            var summary = string.Create(
                CultureInfo.InvariantCulture,
                $"ready after {attempt} polls, {elapsed.ElapsedMilliseconds} ms; states: {trace}");
            HostCommands.RunFromStartup(projectDir, procedure, summary);
            return;
        }

        if (attempt >= MaxPolls)
        {
            var reason = string.Create(
                CultureInfo.InvariantCulture,
                $"Excel not ready after {elapsed.ElapsedMilliseconds} ms; states: {trace}");
            HostCommands.RecordStartupFailure(reason);
            return;
        }

        Task.Delay(PollIntervalMilliseconds).ContinueWith(
            _ => ExcelAsyncUtil.QueueAsMacro(() => Poll(projectDir, procedure, attempt + 1, elapsed, trace)),
            TaskScheduler.Default);
    }

    private readonly record struct ExcelState(bool Visible, bool Splash, bool Ready, int Workbooks, string? Error)
    {
        public bool IsReady => Visible && !Splash && Ready && Workbooks > 0;

        public static ExcelState Probe()
        {
            var splash = SplashIsShowing();
            try
            {
                var application = ExcelDnaUtil.Application;
                var visible = (bool)GetProperty(application, "Visible")!;
                var ready = (bool)GetProperty(application, "Ready")!;
                var workbooks = GetProperty(application, "Workbooks")!;
                try
                {
                    var count = (int)GetProperty(workbooks, "Count")!;
                    return new ExcelState(visible, splash, ready, count, null);
                }
                finally
                {
                    Marshal.ReleaseComObject(workbooks);
                }
            }
            catch (Exception ex) when (ex is COMException or TargetInvocationException or InvalidCastException or NullReferenceException)
            {
                return new ExcelState(false, splash, false, 0, ex.GetType().Name);
            }
        }

        public override string ToString() =>
            Error is null
                ? string.Create(CultureInfo.InvariantCulture, $"V={Visible} S={Splash} R={Ready} W={Workbooks}")
                : "error " + Error;

        private static object? GetProperty(object target, string name) =>
            target.GetType().InvokeMember(name, BindingFlags.GetProperty, binder: null, target, args: null, CultureInfo.InvariantCulture);
    }

    /// <summary>Records state transitions only, as "@poll: state", to keep the summary short.</summary>
    private sealed class StateTrace
    {
        private readonly StringBuilder text = new();
        private string? last;

        public void Record(int attempt, ExcelState state)
        {
            var key = state.ToString();
            if (key == last)
            {
                return;
            }

            last = key;
            if (text.Length > 0)
            {
                text.Append(", ");
            }

            text.Append('@').Append(attempt.ToString(CultureInfo.InvariantCulture)).Append(": ").Append(key);
        }

        public override string ToString() => text.ToString();
    }
}
