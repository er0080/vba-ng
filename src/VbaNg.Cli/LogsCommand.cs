using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang logs [--follow]</c>: prints the project's <c>out/output.log</c>, what its code printed
/// and the errors it left unhandled while Excel ran it (ARCHITECTURE.md section 8; ROADMAP.md
/// WP4). <c>--follow</c> keeps printing what arrives until Ctrl+C, starting over when the log does.
/// </summary>
internal static class LogsCommand
{
    private static readonly TimeSpan Poll = TimeSpan.FromMilliseconds(500);

    public static int Run(CommandLineOptions options)
    {
        var projectDir = ProjectLocator.Resolve(options.Project, out var error);
        if (projectDir is null)
        {
            Console.Error.WriteLine("vbang: " + error);
            return ExitCodes.Usage;
        }

        var path = ProjectPaths.OutputLogPath(projectDir);
        if (!options.Follow)
        {
            if (!File.Exists(path))
            {
                Console.Error.WriteLine($"vbang: no log yet at {path}; Excel writes it while it runs the project.");
                return ExitCodes.Ok;
            }

            Console.Write(ReadFrom(path, 0, out _));
            return ExitCodes.Ok;
        }

        using var cancel = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cancel.Cancel();
        };

        var offset = 0L;
        while (!cancel.IsCancellationRequested)
        {
            if (File.Exists(path))
            {
                var length = new FileInfo(path).Length;
                if (length < offset)
                {
                    // The log grew past its limit and started over.
                    offset = 0;
                }

                if (length > offset)
                {
                    Console.Write(ReadFrom(path, offset, out offset));
                }
            }

            cancel.Token.WaitHandle.WaitOne(Poll);
        }

        return ExitCodes.Ok;
    }

    /// <summary>The log from an offset to its end, read while the add-in may be writing or moving it.</summary>
    private static string ReadFrom(string path, long offset, out long end)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        stream.Seek(offset, SeekOrigin.Begin);
        using var reader = new StreamReader(stream);
        var text = reader.ReadToEnd();
        end = stream.Position;
        return text;
    }
}
