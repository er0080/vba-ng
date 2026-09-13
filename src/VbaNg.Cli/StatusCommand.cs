using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary><c>vbang status</c>: report the add-in version, loaded projects, and the startup run.</summary>
internal static class StatusCommand
{
    public static int Run(CommandLineOptions options)
    {
        var exit = HostCommandClient.Call("vbang.Status", [], out var json);
        if (exit != ExitCodes.Ok)
        {
            return exit;
        }

        Console.WriteLine(options.Json ? json : RunResponse.FromJson(json).Output);
        return ExitCodes.Ok;
    }
}
