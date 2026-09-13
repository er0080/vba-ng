using System.Globalization;
using System.Runtime.InteropServices;

using VbaNg.Compiler;
using VbaNg.Import;
using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>
/// <c>vbang import Book.xlsm [folder] [--to-xlsx]</c> (ARCHITECTURE.md section 10): reads the VBA
/// project out of the workbook with no Excel and no "trust access to the VBA project" setting,
/// writes it as source files next to the workbook (D4), and reports what did not come over. The
/// workbook is never modified (D17); <c>--to-xlsx</c> saves a macro-free copy beside it.
/// </summary>
internal static class ImportCommand
{
    public static int Run(CommandLineOptions options)
    {
        if (options.Positionals.Count == 0)
        {
            Console.Error.WriteLine("vbang: import needs a workbook: vbang import Book.xlsm [folder] [--to-xlsx]");
            return ExitCodes.Usage;
        }

        var workbook = Path.GetFullPath(options.Positionals[0]);
        if (!File.Exists(workbook))
        {
            Console.Error.WriteLine($"vbang: workbook not found: {workbook}");
            return ExitCodes.Usage;
        }

        VbaProject? project;
        try
        {
            project = VbaProjectReader.FromWorkbook(workbook);
        }
        catch (Exception ex) when (ex is InvalidDataException or IOException or NotSupportedException)
        {
            Console.Error.WriteLine($"vbang: {workbook} could not be read: {ex.Message}");
            return ExitCodes.CompileErrors;
        }

        if (project is null)
        {
            Console.Error.WriteLine($"vbang: {workbook} has no VBA project.");
            return ExitCodes.CompileErrors;
        }

        var projectDir = options.Positionals.Count > 1
            ? Path.GetFullPath(options.Positionals[1])
            : Path.ChangeExtension(workbook, ProjectPaths.FolderSuffix);
        // The workbook's own parts say which CodeName is the Workbook, a Worksheet, or a Chart,
        // which the vbaProject.bin does not; the manifest carries them so the folder stands on its
        // own in git (D19). An .xls or an unreadable package reads as empty, leaving the name rule.
        var documentKinds = WorkbookCodeNames.Read(workbook);
        var result = ProjectWriter.Write(project, projectDir, documentKinds: documentKinds);
        Console.WriteLine(result.ToText());

        if (!options.ToXlsx)
        {
            return ExitCodes.Ok;
        }

        var copy = Path.ChangeExtension(workbook, ".xlsx");
        try
        {
            SaveMacroFreeCopy(workbook, copy);
            Console.WriteLine($"Wrote {copy}");
            return ExitCodes.Ok;
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException)
        {
            Console.Error.WriteLine($"vbang: the macro-free copy could not be saved: {ex.Message}");
            return ExitCodes.ExcelUnreachable;
        }
    }

    /// <summary>
    /// Saves a macro-free copy through Excel automation, in a hidden instance of its own so the
    /// user's Excel and the workbook itself are left alone (D14, D17). CodeNames are preserved,
    /// because Excel keeps them with the sheets.
    /// </summary>
    private static void SaveMacroFreeCopy(string workbook, string copy)
    {
        var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new InvalidOperationException("Excel is not installed.");
        var application = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Excel could not be started.");
        try
        {
            Set(application, "Visible", false);
            Set(application, "DisplayAlerts", false);
            var workbooks = Get(application, "Workbooks")!;
            var opened = Invoke(workbooks, "Open", workbook)!;
            try
            {
                if (File.Exists(copy))
                {
                    File.Delete(copy);
                }

                // 51 is xlOpenXMLWorkbook, the macro-free format.
                Invoke(opened, "SaveAs", copy, 51);
                Invoke(opened, "Close", false);
            }
            finally
            {
                Release(opened);
                Release(workbooks);
            }
        }
        finally
        {
            Invoke(application, "Quit");
            Release(application);
        }
    }

    private static object? Get(object target, string name) =>
        target.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture);

    private static void Set(object target, string name, object value) =>
        target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, [value], CultureInfo.InvariantCulture);

    private static object? Invoke(object target, string name, params object[] arguments) =>
        target.GetType().InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod, null, target, arguments, CultureInfo.InvariantCulture);

    private static void Release(object? target)
    {
        if (target is not null && Marshal.IsComObject(target))
        {
            Marshal.FinalReleaseComObject(target);
        }
    }
}
