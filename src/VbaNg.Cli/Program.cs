namespace VbaNg.Cli;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help" or "help")
        {
            PrintUsage();
            return args.Length == 0 ? ExitCodes.Usage : ExitCodes.Ok;
        }

        if (args[0] is "--version" or "-v" or "version")
        {
            Console.WriteLine(Versions.Vbang);
            return ExitCodes.Ok;
        }

        var options = CommandLineOptions.Parse(args, out var error);
        if (options is null)
        {
            Console.Error.WriteLine("vbang: " + error);
            return ExitCodes.Usage;
        }

        switch (options.Command.ToUpperInvariant())
        {
            case "INIT":
                return InitCommand.Run(options);
            case "REPORT":
                return ReportCommand.Run(options);
            case "BUILD":
                return BuildCommand.Run(options);
            case "RUN":
                return RunCommand.Run(options);
            case "TEST":
                return TestCommand.Run(options);
            case "STATUS":
                return StatusCommand.Run(options);
            case "IMPORT":
                return ImportCommand.Run(options);
            case "LOGS":
                return LogsCommand.Run(options);
            default:
                Console.Error.WriteLine($"vbang: unknown command '{options.Command}'.");
                PrintUsage();
                return ExitCodes.Usage;
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("""
            usage:
              vbang init [<workbook>] [--vscode]
                  Start a project beside a workbook: the manifest, a .gitignore, and with --vscode the build task and the attach configuration for debugging in Excel.
              vbang import <workbook> [<folder>] [--to-xlsx]
                  Read the VBA project out of a workbook and write it as source files next to it, with no Excel. --to-xlsx also saves a macro-free copy; the workbook itself is never changed.
              vbang build [<folder>] [--json]
                  Compile a project folder. Prints diagnostics as path(line,col): error VBA0001: message.
              vbang run <Module.Procedure> [--project <folder>] [--ui] [--json]
                  Build the project, then run the procedure in the running Excel and print its Debug.Print output. MsgBox and InputBox answer their defaults unless --ui restores the dialogs.
              vbang test [--project <folder>] [--junit <file>] [--ui] [--json]
                  Build the project, then run its '@Test procedures in the running Excel and print the report.
              vbang status [--json]
                  Report the add-in version, loaded projects, and the startup run in the running Excel.
              vbang report [<folder>] [--project <folder>]
                  Write a zip a bug report can carry: the manifest, the sources, the build diagnostics, the generated C#, the log, and the add-in status. Never the workbook.
              vbang --version
                  Print this build's version, the one Directory.Build.props declares and CI overrides.
              vbang logs [--project <folder>] [--follow]
                  Print the project's out/output.log: what its code printed and the errors it left unhandled while Excel ran it. --follow keeps printing until Ctrl+C.

            The project folder defaults to the current folder when it ends in .vbang, or to the single
            *.vbang folder in the current folder.

            exit codes: 0 ok, 1 compile errors, 2 usage, 3 Excel not reachable, 4 run-time error, 5 tests failed
            """);
    }
}
