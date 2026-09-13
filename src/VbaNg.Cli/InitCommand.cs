using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang init [&lt;workbook&gt;] [--vscode]</c> (ARCHITECTURE.md section 8, D11): starts a project
/// beside a workbook, which is what binds the two (D4). Writes the manifest with the references
/// every Excel project has, a <c>.gitignore</c> for the build output, and with <c>--vscode</c> the
/// build task and the attach configuration that debugging inside Excel needs.
/// </summary>
internal static class InitCommand
{
    /// <summary>The workbook formats a project can sit beside (D4); the same list the compiler reads CodeNames from.</summary>
    private static readonly string[] WorkbookExtensions = [".xlsx", ".xlsm", ".xlsb", ".xlam", ".xltx", ".xltm"];

    public static int Run(CommandLineOptions options)
    {
        if (Workbook(options, out var workbook) is { } failure)
        {
            return failure;
        }

        var directory = Path.GetDirectoryName(workbook)!;
        var name = Path.GetFileNameWithoutExtension(workbook);
        var projectDir = Path.Combine(directory, name + ProjectPaths.FolderSuffix);
        var manifestPath = Path.Combine(projectDir, "vbang.json");
        if (File.Exists(manifestPath))
        {
            Console.Error.WriteLine($"vbang: {manifestPath} already exists; init never overwrites a project.");
            return ExitCodes.Usage;
        }

        Directory.CreateDirectory(projectDir);
        var written = new List<string>();
        Write(manifestPath, Manifest(name), written);

        // The build output is derived and rebuilt, so it stays out of git (ARCHITECTURE.md section 3).
        Write(Path.Combine(projectDir, ".gitignore"), "out/\n", written);

        if (options.VsCode)
        {
            var vscodeDir = Path.Combine(directory, ".vscode");
            Directory.CreateDirectory(vscodeDir);
            Write(Path.Combine(vscodeDir, "tasks.json"), Tasks(projectDir, directory), written);
            Write(Path.Combine(vscodeDir, "launch.json"), Launch(), written);
        }

        Console.WriteLine($"Started {Path.GetFileName(projectDir)} beside {Path.GetFileName(workbook)}:");
        foreach (var file in written)
        {
            Console.WriteLine("  " + Path.GetRelativePath(directory, file).Replace('\\', '/'));
        }

        Console.WriteLine($"Put a .bas or .cls in the folder and run vbang build {Path.GetFileName(projectDir)}.");
        return ExitCodes.Ok;
    }

    /// <summary>The workbook the project belongs to: the one named, or the only one in the folder.</summary>
    private static int? Workbook(CommandLineOptions options, out string workbook)
    {
        workbook = string.Empty;
        if (options.Positionals.Count > 0)
        {
            var named = Path.GetFullPath(options.Positionals[0]);
            if (!File.Exists(named))
            {
                Console.Error.WriteLine($"vbang: workbook not found: {named}");
                return ExitCodes.Usage;
            }

            workbook = named;
            return null;
        }

        var here = Directory.GetCurrentDirectory();
        var found = Directory.EnumerateFiles(here)
            .Where(f => WorkbookExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
            .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();

        switch (found.Count)
        {
            case 1:
                workbook = found[0];
                return null;
            case 0:
                Console.Error.WriteLine($"vbang: no workbook in {here}: vbang init <workbook> [--vscode]");
                return ExitCodes.Usage;
            default:
                Console.Error.WriteLine($"vbang: {found.Count} workbooks in {here}; name the one to start a project for: vbang init <workbook> [--vscode]");
                return ExitCodes.Usage;
        }
    }

    private static void Write(string path, string content, List<string> written)
    {
        File.WriteAllText(path, content.ReplaceLineEndings("\n"), new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        written.Add(path);
    }

    /// <summary>
    /// Excel and Office are the references a VBA project in Excel carries by default, so a project
    /// that starts here binds the same names an existing one does (ARCHITECTURE.md section 3).
    /// </summary>
    private static string Manifest(string name) =>
        $$"""
        {
          "name": "{{name}}",
          "references": [
            { "name": "Excel" },
            { "name": "Office" }
          ]
        }

        """;

    /// <summary>The build task, whose problem matcher reads the canonical diagnostics vbang prints (D5, D11).</summary>
    private static string Tasks(string projectDir, string directory) =>
        $$"""
        {
          "version": "2.0.0",
          "tasks": [
            {
              "label": "vbang build",
              "type": "shell",
              "command": "vbang",
              "args": ["build", "{{Path.GetRelativePath(directory, projectDir).Replace('\\', '/')}}"],
              "group": { "kind": "build", "isDefault": true },
              "problemMatcher": "$msCompile"
            }
          ]
        }

        """;

    /// <summary>
    /// Attaching to EXCEL.EXE rather than launching it: the add-in is already loaded there, and the
    /// PDB vbang build writes beside the assembly is what makes a breakpoint in a .bas bind (D11).
    /// </summary>
    private static string Launch() =>
        """
        {
          "version": "0.2.0",
          "configurations": [
            {
              "name": "Attach to Excel",
              "type": "coreclr",
              "request": "attach",
              "processName": "EXCEL.EXE"
            }
          ]
        }

        """;
}
