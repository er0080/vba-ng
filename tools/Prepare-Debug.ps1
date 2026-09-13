<#
.SYNOPSIS
Prepares an F5 debug session: closes throwaway Excel instances running this repo's dev build of
the add-in, builds the solution, then builds a project folder with vbang.

.DESCRIPTION
Excel keeps the add-in files locked while it runs, so a solution build fails while any Excel has
the dev build loaded. An Excel is a throwaway if it has loaded the add-in from this repo's
src\VbaNg.AddIn\bin folder, which is true of instances started by the VS Code launch
configurations and by Start-ExcelWithAddIn.ps1, and never of an Excel used for real work. Such
instances are asked to quit through COM (unsaved changes are discarded) and terminated if they
cannot be reached. Every other Excel instance is left alone. Windows PowerShell 5.1 compatible
(CLAUDE.md R22).

.PARAMETER Project
Project folder to build with vbang after the solution build. Optional.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\Prepare-Debug.ps1 -Project samples\Hello\Hello.vbang
#>
param(
    [string]$Project
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$addInBin = Join-Path $root 'src\VbaNg.AddIn\bin'

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class VbaNgExcelWindow
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string windowName);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object dispatch);

    // Excel registers in the ROT only after losing focus, so reach the object model through the
    // workbook window (XLMAIN > XLDESK > EXCEL7) of the given process. Null when it has no workbook.
    public static object ApplicationForProcess(uint pid)
    {
        IntPtr main = IntPtr.Zero;
        while ((main = FindWindowEx(IntPtr.Zero, main, "XLMAIN", null)) != IntPtr.Zero)
        {
            uint owner; GetWindowThreadProcessId(main, out owner);
            if (owner != pid) continue;
            IntPtr desk = FindWindowEx(main, IntPtr.Zero, "XLDESK", null);
            if (desk == IntPtr.Zero) continue;
            IntPtr book = FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
            if (book == IntPtr.Zero) continue;
            Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
            object window;
            if (AccessibleObjectFromWindow(book, 0xFFFFFFF0, ref iid, out window) != 0) continue;
            return window.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, window, null);
        }
        return null;
    }
}
"@

foreach ($excel in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
    $isDevInstance = $false
    try {
        $isDevInstance = [bool]($excel.Modules | Where-Object { $_.FileName -like "$addInBin\*" })
    }
    catch {
        # Module enumeration can fail for processes of another user or bitness; not ours then.
    }

    if (-not $isDevInstance) {
        Write-Output "Leaving Excel $($excel.Id) alone: it does not have the dev add-in loaded."
        continue
    }

    Write-Output "Closing throwaway Excel $($excel.Id): it has the dev add-in loaded and locks the build output."
    $app = [VbaNgExcelWindow]::ApplicationForProcess([uint32]$excel.Id)
    if ($app -ne $null) {
        try {
            $app.DisplayAlerts = $false
            $app.Quit()
        }
        catch {
            Write-Output "  COM quit failed ($($_.Exception.Message.Trim())); terminating."
        }
        finally {
            [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
        }
    }

    if (-not $excel.WaitForExit(15000)) {
        Write-Output "  Still running; terminating."
        Stop-Process -Id $excel.Id -Force
        $excel.WaitForExit(15000) | Out-Null
    }
}

Push-Location $root
try {
    dotnet build VbaNg.slnx
    $exit = $LASTEXITCODE
    if ($exit -eq 0 -and $Project) {
        dotnet run --project src/VbaNg.Cli --no-build -- build $Project
        $exit = $LASTEXITCODE
    }
}
finally {
    Pop-Location
}

exit $exit
