using VbaNg.Compiler;
using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang run Module.Procedure</c>: build the project, then ask the add-in in the running Excel
/// to run the procedure and print its <c>Debug.Print</c> output (ARCHITECTURE.md D6, section 9).
/// </summary>
internal static class RunCommand
{
    public static int Run(CommandLineOptions options)
    {
        if (options.Positionals.Count != 1)
        {
            Console.Error.WriteLine("vbang: usage: vbang run <Module.Procedure> [--project <folder>] [--ui] [--json]");
            return ExitCodes.Usage;
        }

        var procedure = options.Positionals[0];
        var projectDir = ProjectLocator.Resolve(options.Project, out var error);
        if (projectDir is null)
        {
            Console.Error.WriteLine("vbang: " + error);
            return ExitCodes.Usage;
        }

        var build = ProjectCompiler.Build(projectDir, TypeLibraryCache.Resolve);
        if (!build.Success)
        {
            BuildCommand.Report(build, options.Json);
            return ExitCodes.CompileErrors;
        }

        var exit = HostCommandClient.Call("vbang.Run", [projectDir, procedure, options.Ui ? "ui" : string.Empty], out var json);
        if (exit != ExitCodes.Ok)
        {
            return exit;
        }

        var response = RunResponse.FromJson(json);
        if (options.Json)
        {
            Console.WriteLine(json);
            return response.Ok ? ExitCodes.Ok : ExitCodes.RuntimeError;
        }

        if (response.Output.Length > 0)
        {
            Console.WriteLine(response.Output);
        }

        if (!response.Ok)
        {
            Console.Error.WriteLine("vbang: run-time error: " + response.Error);
            if (response.Detail is not null)
            {
                Console.Error.WriteLine(response.Detail);
            }

            return ExitCodes.RuntimeError;
        }

        return ExitCodes.Ok;
    }
}
