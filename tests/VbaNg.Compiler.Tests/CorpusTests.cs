using System.Globalization;
using System.Text;

using VbaNg.Compiler.Syntax;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Whole-file parsing: the repository's fixture modules must parse cleanly and round-trip, and an
/// external corpus (ARCHITECTURE.md section 11, parse-rate corpus) must reach the M1 exit
/// criterion of 99 percent. The corpus is not checked in; point VBANG_CORPUS at a folder of
/// .bas, .cls, and .frm files. Every file, parsed or not, must round-trip byte for byte.
/// </summary>
public sealed class CorpusTests(ITestOutputHelper output)
{
    private static readonly string[] Extensions = [".bas", ".cls", ".frm"];

    public static TheoryData<string> FixtureFiles()
    {
        var root = Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures");
        var data = new TheoryData<string>();
        foreach (var file in EnumerateSources(root).OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            data.Add(Path.GetRelativePath(root, file));
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(FixtureFiles))]
    public void Fixture_ParsesWithoutDiagnosticsAndRoundTrips(string relativePath)
    {
        var path = Path.Combine(TestPaths.RepositoryRoot, "tests", "VbaNg.Compiler.Tests", "Fixtures", relativePath);
        var text = ReadSource(path);

        var tree = SyntaxTree.Parse(text, path);

        Assert.Equal(text, tree.Root.ToFullString());
        Assert.Empty(tree.Diagnostics.Select(d => d.ToString()));
        Assert.False(tree.Root.ContainsMissingTokens);
        Assert.DoesNotContain(tree.Root.DescendantNodes(), n => n is BadStatementSyntax);
    }

    [Fact]
    public void Corpus_ParseRateMeetsTheM1ExitCriterion()
    {
        var root = Environment.GetEnvironmentVariable("VBANG_CORPUS");
        Assert.SkipWhen(string.IsNullOrWhiteSpace(root), "Set VBANG_CORPUS to a folder of .bas, .cls, and .frm files to run the parse-rate corpus.");
        Assert.SkipUnless(Directory.Exists(root), $"VBANG_CORPUS folder not found: {root}");

        var files = EnumerateSources(root!).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList();
        Assert.NotEmpty(files);

        var failures = new List<string>();
        var roundTripFailures = new List<string>();
        var messages = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var file in files)
        {
            var text = ReadSource(file);
            var tree = SyntaxTree.Parse(text, file);
            if (tree.Root.ToFullString() != text)
            {
                roundTripFailures.Add(file);
            }

            if (tree.HasErrors)
            {
                var first = tree.Diagnostics.First(d => d.IsError);
                failures.Add(first.ToString());
                messages[first.Message] = messages.GetValueOrDefault(first.Message) + 1;
            }
        }

        var rate = (files.Count - failures.Count) / (double)files.Count;
        var report = new StringBuilder();
        report.AppendLine(CultureInfo.InvariantCulture, $"Corpus: {root}");
        report.AppendLine(CultureInfo.InvariantCulture, $"Files: {files.Count}, parsed: {files.Count - failures.Count}, failed: {failures.Count}, rate: {rate:P2}");
        report.AppendLine(CultureInfo.InvariantCulture, $"Round-trip failures: {roundTripFailures.Count}");
        report.AppendLine("First error per failing file:");
        foreach (var failure in failures)
        {
            report.AppendLine("  " + failure);
        }

        report.AppendLine("Failures by message:");
        foreach (var (message, count) in messages.OrderByDescending(p => p.Value))
        {
            report.AppendLine(CultureInfo.InvariantCulture, $"  {count,5}  {message}");
        }

        output.WriteLine(report.ToString());
        var reportPath = Path.Combine(AppContext.BaseDirectory, "corpus-report.txt");
        File.WriteAllText(reportPath, report.ToString());
        output.WriteLine("Report written to " + reportPath);

        Assert.Empty(roundTripFailures);
        Assert.True(rate >= 0.99, $"Parse rate {rate:P2} is below the 99% exit criterion; see {reportPath}");
    }

    private static IEnumerable<string> EnumerateSources(string root) =>
        Directory.Exists(root)
            ? Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
                .Where(f => Extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                .Where(f => !f.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                .Where(f => !f.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            : [];

    /// <summary>Reads a module as the VBE would write it: BOM-detected, otherwise the system ANSI code page (CLAUDE.md gotchas).</summary>
    private static string ReadSource(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        try
        {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.Latin1.GetString(bytes);
        }
    }
}

internal static class TestPaths
{
    public static string RepositoryRoot { get; } = FindRepositoryRoot();

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "VbaNg.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Repository root (VbaNg.slnx) not found above " + AppContext.BaseDirectory);
    }
}
