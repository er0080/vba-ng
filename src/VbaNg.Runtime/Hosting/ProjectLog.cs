using System.Globalization;
using System.Text;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// The project's log, <c>out/output.log</c> (ARCHITECTURE.md section 3): every line the project
/// prints and every error it leaves unhandled while Excel runs it, each stamped with the local
/// time, so a bug report can carry it and <c>vbang logs</c> can show it (ROADMAP.md WP4). A log
/// past <see cref="MaxBytes"/> moves to <c>output.log.1</c>, replacing the one before. Writing
/// never fails a run: a log that cannot be written is skipped.
/// </summary>
public static class ProjectLog
{
    /// <summary>The size past which the log starts over, keeping one previous file.</summary>
    public const long MaxBytes = 1024 * 1024;

    private static readonly Lock Gate = new();

    public static void Append(string projectDir, string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        ArgumentNullException.ThrowIfNull(text);
        var stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff ", CultureInfo.InvariantCulture);
        var entry = new StringBuilder();
        foreach (var line in text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
        {
            entry.Append(stamp).Append(line).Append(Environment.NewLine);
        }

        var path = ProjectPaths.OutputLogPath(projectDir);
        lock (Gate)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                var existing = new FileInfo(path);
                if (existing.Exists && existing.Length >= MaxBytes)
                {
                    File.Move(path, path + ".1", overwrite: true);
                }

                File.AppendAllText(path, entry.ToString(), new UTF8Encoding(false));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The log is a convenience; the run goes on without it.
            }
        }
    }
}
