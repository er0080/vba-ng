using System.Text.RegularExpressions;

using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The project's log, <c>out/output.log</c> (ARCHITECTURE.md section 3; ROADMAP.md WP4): every
/// line the project prints and every error it leaves unhandled while Excel runs it, each with the
/// time it was written, so a bug report can carry what happened. A log that grows past its limit
/// keeps one previous file.
/// </summary>
public sealed partial class ProjectLogTests : IDisposable
{
    private readonly string projectDir = Path.Combine(Path.GetTempPath(), "vbang-tests", Guid.NewGuid().ToString("N"), "Logged.vbang");

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path.GetDirectoryName(projectDir)!, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    [Fact]
    public void Append_WritesATimestampedLineToOutputLog()
    {
        ProjectLog.Append(projectDir, "hello");
        ProjectLog.Append(projectDir, "world");

        var lines = File.ReadAllLines(ProjectPaths.OutputLogPath(projectDir));
        Assert.Equal(2, lines.Length);
        Assert.Matches(Stamped(), lines[0]);
        Assert.EndsWith(" hello", lines[0], StringComparison.Ordinal);
        Assert.EndsWith(" world", lines[1], StringComparison.Ordinal);
    }

    [Fact]
    public void Append_TextOfSeveralLines_StampsEachLine()
    {
        ProjectLog.Append(projectDir, "first\r\nsecond\nthird");

        var lines = File.ReadAllLines(ProjectPaths.OutputLogPath(projectDir));
        Assert.Equal(3, lines.Length);
        Assert.All(lines, line => Assert.Matches(Stamped(), line));
        Assert.EndsWith(" third", lines[2], StringComparison.Ordinal);
    }

    [Fact]
    public void Append_PastTheLimit_KeepsOnePreviousFile()
    {
        var big = new string('x', (int)ProjectLog.MaxBytes);
        ProjectLog.Append(projectDir, big);
        ProjectLog.Append(projectDir, "next");

        var path = ProjectPaths.OutputLogPath(projectDir);
        Assert.True(File.Exists(path + ".1"), "the full log should move to output.log.1");
        Assert.EndsWith(" next", Assert.Single(File.ReadAllLines(path)), StringComparison.Ordinal);
    }

    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3} ")]
    private static partial Regex Stamped();
}
