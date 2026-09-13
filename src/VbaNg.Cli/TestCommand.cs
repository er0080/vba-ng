using VbaNg.Compiler;
using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang test</c>: build the project, then ask the add-in in the running Excel to run every
/// <c>'@Test</c> procedure and print the report (ARCHITECTURE.md D6, section 9). <c>--junit</c>
/// names a file the add-in writes the JUnit XML report to. Exit code 5 means tests failed.
/// </summary>
internal static class TestCommand
{
    public static int Run(CommandLineOptions options)
    {
        if (options.Positionals.Count != 0)
        {
            Console.Error.WriteLine("vbang: usage: vbang test [--project <folder>] [--junit <file>] [--ui] [--json]");
            return ExitCodes.Usage;
        }

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

        var junitPath = string.Empty;
        if (options.Junit is not null)
        {
            junitPath = Path.GetFullPath(options.Junit);
            Directory.CreateDirectory(Path.GetDirectoryName(junitPath)!);
        }

        var exit = HostCommandClient.Call("vbang.Test", [projectDir, junitPath, options.Ui ? "ui" : string.Empty], out var json);
        if (exit != ExitCodes.Ok)
        {
            return exit;
        }

        var response = RunResponse.FromJson(json);
        if (options.Json)
        {
            Console.WriteLine(json);
            return response.Ok ? ExitCodes.Ok : response.Detail is null ? ExitCodes.TestsFailed : ExitCodes.RuntimeError;
        }

        if (response.Output.Length > 0)
        {
            Console.WriteLine(response.Output);
        }

        if (response.Ok)
        {
            return ExitCodes.Ok;
        }

        // Failed tests come back without a detail; a detail is the stack of an exception in the command itself.
        if (response.Detail is null)
        {
            return ExitCodes.TestsFailed;
        }

        Console.Error.WriteLine("vbang: " + response.Error);
        Console.Error.WriteLine(response.Detail);
        return ExitCodes.RuntimeError;
    }
}
