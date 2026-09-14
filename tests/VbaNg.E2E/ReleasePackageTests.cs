using System.Diagnostics;
using System.IO.Compression;

using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.E2E;

/// <summary>
/// The release zip, installed as README.md installs it: unpacked, its vba-ng.xll loaded, and the Quickstart workbook in
/// it opened. No build output ships, so the add-in has to build the project with the vbang.exe beside it; the button's
/// macro then writes B2, and vbang --version names the version the zip is named for. Runs when VBANG_E2E is set and
/// VBANG_PACKAGE names a zip, as tools/Invoke-Gate.ps1 -Package does.
/// </summary>
[Trait("Category", "E2E")]
public sealed class ReleasePackageTests : IDisposable
{
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
    public void ReleaseZip_Quickstart_BuildsWithItsOwnCliAndRuns()
    {
        Assert.SkipWhen(string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VBANG_E2E")), "Set VBANG_E2E=1 to run the Excel-driving tests.");
        var zip = Environment.GetEnvironmentVariable("VBANG_PACKAGE");
        Assert.SkipWhen(string.IsNullOrWhiteSpace(zip), "Set VBANG_PACKAGE to a release zip, as tools/Invoke-Gate.ps1 -Package does.");
        Assert.True(File.Exists(zip), "No release zip at " + zip);

        ZipFile.ExtractToDirectory(zip, workDir);
        var root = Assert.Single(Directory.GetDirectories(workDir));
        var projectDir = Path.Combine(root, "samples", "Quickstart", "Quickstart" + ProjectPaths.FolderSuffix);
        Assert.False(Directory.Exists(Path.Combine(projectDir, "out")), "The zip ships build output, so the add-in's build is not tested.");

        var info = new ProcessStartInfo(Path.Combine(root, "vbang.exe"), "--version") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
        using (var cli = Process.Start(info)!)
        {
            var version = cli.StandardOutput.ReadToEnd().Trim();
            cli.WaitForExit();
            Assert.Equal("vbang-" + version + "-win-x64", Path.GetFileName(root));
        }

        string? b2 = null, status = null;
        Sta.Run(() =>
        {
            using var excel = ExcelInstance.Start(Path.Combine(root, "vba-ng.xll"));
            var workbook = excel.OpenWorkbook(Path.Combine(root, "samples", "Quickstart", "Quickstart.xlsx"));
            excel.RunMacro("Greeting.SayHello", attempts: 12);
            b2 = ExcelInstance.CellValue(workbook, "Sheet1", "B2") as string;
            status = RunResponse.FromJson((string)excel.RunMacro("vbang.Status")!).Output;
            ExcelInstance.CloseWorkbook(workbook);
        });

        Assert.True(b2 == "Hello, world", $"B2 was '{b2}'; status: {status}");
        Assert.True(File.Exists(Path.Combine(projectDir, "out", "build.json")), "The add-in did not build the project. Status: " + status);
    }
}
