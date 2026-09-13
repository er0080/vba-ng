using System.Globalization;
using System.Text.Json;

using VbaNg.Compiler;

namespace VbaNg.Cli;

/// <summary><c>vbang build</c>: compile a project folder and print canonical diagnostics (ARCHITECTURE.md D5).</summary>
internal static class BuildCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static int Run(CommandLineOptions options)
    {
        var positional = options.Positionals.Count > 0 ? options.Positionals[0] : null;
        var projectDir = ProjectLocator.Resolve(options.Project ?? positional, out var error);
        if (projectDir is null)
        {
            Console.Error.WriteLine("vbang: " + error);
            return ExitCodes.Usage;
        }

        var result = ProjectCompiler.Build(projectDir, TypeLibraryCache.Resolve);
        Report(result, options.Json);
        return result.Success ? ExitCodes.Ok : ExitCodes.CompileErrors;
    }

    public static void Report(BuildResult result, bool json)
    {
        if (json)
        {
            var payload = new
            {
                success = result.Success,
                project = result.ProjectName,
                assembly = result.AssemblyPath,
                diagnostics = result.Diagnostics.Select(d => new
                {
                    id = d.Id,
                    severity = d.IsError ? "error" : "warning",
                    message = d.Message,
                    file = d.FilePath,
                    line = d.Line,
                    column = d.Column,
                }),
            };
            Console.WriteLine(JsonSerializer.Serialize(payload, JsonOptions));
            return;
        }

        foreach (var diagnostic in result.Diagnostics)
        {
            Console.WriteLine(diagnostic.ToString());
        }

        var errors = result.Diagnostics.Count(d => d.IsError);
        Console.WriteLine(result.Success
            ? "Build succeeded: " + result.AssemblyPath
            : string.Create(CultureInfo.InvariantCulture, $"Build failed: {errors} error(s)."));
    }
}
