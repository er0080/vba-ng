using System.Runtime.InteropServices;

using VbaNg.Runtime.Hosting;

namespace VbaNg.Cli;

/// <summary>Calls a host command in the running Excel and reports why when that is not possible.</summary>
internal static class HostCommandClient
{
    /// <summary>
    /// Returns <see cref="ExitCodes.Ok"/> with the command's JSON response, or an exit code after
    /// printing the reason to stderr when Excel or the add-in cannot be reached.
    /// </summary>
    public static int Call(string command, object[] arguments, out string json)
    {
        json = string.Empty;
        var application = ExcelAutomation.GetRunningExcel();
        if (application is null)
        {
            Console.Error.WriteLine("vbang: no running Excel instance found. Start Excel with the vba-ng add-in loaded.");
            return ExitCodes.ExcelUnreachable;
        }

        try
        {
            try
            {
                json = ExcelAutomation.Run(application, command, arguments) as string ?? string.Empty;
            }
            catch (COMException ex)
            {
                Console.Error.WriteLine($"vbang: Excel could not run '{command}'. Is the vba-ng add-in loaded? ({ex.Message})");
                return ExitCodes.ExcelUnreachable;
            }

            if (json.Length == 0)
            {
                Console.Error.WriteLine($"vbang: '{command}' returned nothing; the add-in may not be loaded.");
                return ExitCodes.ExcelUnreachable;
            }

            // A long response comes back as a stub naming out/response.json (ARCHITECTURE.md section 7).
            try
            {
                json = RunResponse.ResolveTransport(json);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Console.Error.WriteLine($"vbang: '{command}' wrote its response to a file that could not be read: {ex.Message}");
                return ExitCodes.RuntimeError;
            }

            return ExitCodes.Ok;
        }
        finally
        {
            Marshal.FinalReleaseComObject(application);
        }
    }
}
