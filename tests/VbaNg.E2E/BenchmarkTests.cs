using System.Globalization;
using System.Reflection;
using System.Text;

using VbaNg.Compiler;
using VbaNg.Interop;
using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.E2E;

/// <summary>
/// The performance baseline of ROADMAP.md WP7: samples/Benchmarks runs under <c>vbang.Run</c> in
/// a hidden Excel with its workbook bound, so the UDF is registered, and the same module runs
/// imported into the VBA project of another hidden Excel. The best of two runs each side and
/// the ratios land next to the test dll as benchmark-report.txt. Numbers, not fixes: the test
/// fails only when a benchmark does not run or computes a wrong result. Skipped unless
/// VBANG_E2E is set (CLAUDE.md R9, R17).
/// </summary>
[Trait("Category", "E2E")]
public sealed class BenchmarkTests : IDisposable
{
    private const string GateVariable = "VBANG_E2E";
    private const int Runs = 2;
    private static readonly TimeSpan CommandTimeout = TimeSpan.FromSeconds(300);
    private static readonly string[] Names = ["Variant loop", "Double arithmetic", "String building", "Range loop", "UDF over 10,000 cells", "Late-bound Collection", "Late-bound Excel"];
    private static readonly string[] Functions = ["VariantLoop", "DoubleArithmetic", "StringBuilding", "RangeLoop", "UdfCells", "LateBoundCollection", "LateBoundExcel"];
    private readonly string workDir = Path.Combine(Path.GetTempPath(), "vbang-e2e", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try
        {
            Directory.Delete(workDir, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Fact]
    public void Benchmarks_RunUnderVbangAndVba_ReportTheRatios()
    {
        var addIn = RequireAddIn();
        var sample = Path.Combine(RepositoryRoot(), "samples", "Benchmarks");
        var sampleDir = Path.Combine(workDir, "Benchmarks");
        var projectDir = Path.Combine(sampleDir, "Benchmarks" + ProjectPaths.FolderSuffix);
        Directory.CreateDirectory(projectDir);
        var workbookPath = Path.Combine(sampleDir, "Benchmarks.xlsx");
        File.Copy(Path.Combine(sample, "Benchmarks.xlsx"), workbookPath);
        foreach (var file in Directory.GetFiles(Path.Combine(sample, "Benchmarks" + ProjectPaths.FolderSuffix)))
        {
            File.Copy(file, Path.Combine(projectDir, Path.GetFileName(file)));
        }

        var build = ProjectCompiler.Build(projectDir, reference => TypeLibraryCache.Resolve(reference.Name, reference.Guid, reference.Version));
        Assert.True(build.Success, string.Join(Environment.NewLine, build.Diagnostics));

        // Under vbang: Main prints one line per benchmark.
        var vbang = new Dictionary<string, double>(StringComparer.Ordinal);
        var responses = new List<RunResponse>();
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.OpenWorkbook(workbookPath);
            for (var run = 0; run < Runs; run++)
            {
                responses.Add(RunResponse.FromJson(excel.RunHostCommand("vbang.Run", CommandTimeout, projectDir, "Bench.Main", string.Empty)));
            }

            ExcelInstance.CloseWorkbook(workbook);
        });

        foreach (var response in responses)
        {
            Assert.True(response.Ok, response.Error + Environment.NewLine + response.Detail + Environment.NewLine + response.Output);
            foreach (var line in response.Output.Split(Environment.NewLine))
            {
                var separator = line.LastIndexOf(':');
                Assert.True(separator > 0 && line.EndsWith(" ms", StringComparison.Ordinal), "Unexpected benchmark line: " + line);
                var name = line[..separator];
                var ms = double.Parse(line[(separator + 1)..^3], CultureInfo.InvariantCulture);
                vbang[name] = vbang.TryGetValue(name, out var best) ? Math.Min(best, ms) : ms;
            }
        }

        // Under VBA: the same module imported into a throwaway workbook's project, each function called by name.
        var vba = new Dictionary<string, double>(StringComparer.Ordinal);
        var modulePath = Path.Combine(projectDir, "Bench.bas");
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(addIn);
            var workbook = excel.ActiveWorkbook();
            var qualifier = "'" + ExcelInstance.Name(workbook) + "'!Bench.";
            ExcelInstance.ImportModule(workbook, modulePath);
            for (var run = 0; run < Runs; run++)
            {
                for (var i = 0; i < Functions.Length; i++)
                {
                    var ms = Convert.ToDouble(excel.RunMacro(qualifier + Functions[i]), CultureInfo.InvariantCulture);
                    vba[Names[i]] = vba.TryGetValue(Names[i], out var best) ? Math.Min(best, ms) : ms;
                }
            }

            ExcelInstance.Release(workbook);
        });

        var report = new StringBuilder();
        report.AppendLine(CultureInfo.InvariantCulture, $"Benchmarks: best of {Runs} runs each, milliseconds, {DateTime.Now:yyyy-MM-dd}");
        report.AppendLine("Benchmark                vbang ms   VBA ms   vbang/VBA");
        foreach (var name in Names)
        {
            Assert.True(vbang.ContainsKey(name), name + " did not run under vbang: " + string.Join(" | ", vbang.Keys));
            Assert.True(vba.ContainsKey(name), name + " did not run under VBA.");
            report.AppendLine(CultureInfo.InvariantCulture, $"{name,-24} {vbang[name],8:0} {vba[name],8:0} {vbang[name] / vba[name],11:0.00}x");
        }

        File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "benchmark-report.txt"), report.ToString());
    }

    private static string RequireAddIn()
    {
        Assert.SkipWhen(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(GateVariable)), $"Set {GateVariable}=1 to run the Excel-driving tests.");
        var configuration = typeof(BenchmarkTests).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? "Debug";
        Assert.SkipWhen(configuration != "Release", "Benchmarks run the Release build: dotnet test tests/VbaNg.E2E -c Release.");
        var path = Path.Combine(RepositoryRoot(), "src", "VbaNg.AddIn", "bin", configuration, "net10.0-windows", "VbaNg.AddIn-AddIn64.xll");
        Assert.True(File.Exists(path), "Add-in not built: " + path);
        return path;
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "VbaNg.slnx")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Repository root (VbaNg.slnx) not found above " + AppContext.BaseDirectory);
    }
}
