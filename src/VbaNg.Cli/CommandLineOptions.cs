namespace VbaNg.Cli;

/// <summary>Minimal argument parsing: a command, positionals, <c>--project</c>, <c>--junit</c>, <c>--json</c>, and the flags.</summary>
internal sealed class CommandLineOptions
{
    private CommandLineOptions(string command, IReadOnlyList<string> positionals, string? project, string? junit, bool json, bool ui, bool toXlsx, bool follow, bool vsCode)
    {
        Command = command;
        Positionals = positionals;
        Project = project;
        Junit = junit;
        Json = json;
        Ui = ui;
        ToXlsx = toXlsx;
        Follow = follow;
        VsCode = vsCode;
    }

    public string Command { get; }

    public IReadOnlyList<string> Positionals { get; }

    public string? Project { get; }

    /// <summary>The JUnit XML report file of <c>vbang test</c>.</summary>
    public string? Junit { get; }

    public bool Json { get; }

    /// <summary>Keep dialogs interactive during a run (ARCHITECTURE.md D8).</summary>
    public bool Ui { get; }

    /// <summary>Save a macro-free copy of the workbook an import read (ARCHITECTURE.md D14).</summary>
    public bool ToXlsx { get; }

    /// <summary>Keep printing the log as it grows (<c>vbang logs</c>).</summary>
    public bool Follow { get; }

    /// <summary>Write the VS Code build task and attach configuration (<c>vbang init</c>, ARCHITECTURE.md D11).</summary>
    public bool VsCode { get; }

    public static CommandLineOptions? Parse(string[] args, out string? error)
    {
        error = null;
        if (args.Length == 0)
        {
            error = "No command given.";
            return null;
        }

        var positionals = new List<string>();
        string? project = null;
        string? junit = null;
        var json = false;
        var ui = false;
        var toXlsx = false;
        var follow = false;
        var vsCode = false;

        for (var i = 1; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "--json")
            {
                json = true;
            }
            else if (arg == "--ui")
            {
                ui = true;
            }
            else if (arg == "--to-xlsx")
            {
                toXlsx = true;
            }
            else if (arg == "--follow")
            {
                follow = true;
            }
            else if (arg == "--vscode")
            {
                vsCode = true;
            }
            else if (TryValue(args, ref i, "--project", out var projectValue, out error))
            {
                project = projectValue;
            }
            else if (TryValue(args, ref i, "--junit", out var junitValue, out error))
            {
                junit = junitValue;
            }
            else if (error is not null)
            {
                return null;
            }
            else if (arg.StartsWith("--", StringComparison.Ordinal))
            {
                error = $"Unknown option '{arg}'.";
                return null;
            }
            else
            {
                positionals.Add(arg);
            }
        }

        return new CommandLineOptions(args[0], positionals, project, junit, json, ui, toXlsx, follow, vsCode);
    }

    /// <summary>Matches <c>--name value</c> and <c>--name=value</c>; an option without its value sets <paramref name="error"/>.</summary>
    private static bool TryValue(string[] args, ref int index, string option, out string? value, out string? error)
    {
        value = null;
        error = null;
        var arg = args[index];
        if (arg == option)
        {
            if (index + 1 >= args.Length)
            {
                error = $"{option} requires a value.";
                return false;
            }

            value = args[++index];
            return true;
        }

        if (arg.StartsWith(option + "=", StringComparison.Ordinal))
        {
            value = arg[(option.Length + 1)..];
            return true;
        }

        return false;
    }
}
