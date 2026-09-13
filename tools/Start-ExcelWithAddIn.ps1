<#
.SYNOPSIS
Starts a throwaway Excel instance with the vba-ng add-in loaded (CLAUDE.md R17, R31).

.DESCRIPTION
Creates Excel through COM, adds a blank workbook so the CLI can find the instance, and loads the
add-in with Application.RegisterXLL. Registering programmatically avoids the security notice Excel
shows when an unsigned .xll is opened as a file. Prints the Excel process id and the add-in's
vbang.Status response. Windows PowerShell 5.1 compatible (CLAUDE.md R22).

.PARAMETER AddInPath
Path to the built .xll. Defaults to the Debug build output.

.PARAMETER Hidden
Keep the Excel window hidden (for unattended runs).

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\Start-ExcelWithAddIn.ps1

.NOTES
The default execution policy on Windows client editions is Restricted, which blocks every script
file. -ExecutionPolicy Bypass lifts it for this one process without changing the machine setting.
#>
param(
    [string]$AddInPath = (Join-Path $PSScriptRoot '..\src\VbaNg.AddIn\bin\Debug\net10.0-windows\VbaNg.AddIn-AddIn64.xll'),
    [switch]$Hidden
)

$ErrorActionPreference = 'Stop'
$AddInPath = [IO.Path]::GetFullPath($AddInPath)
if (-not (Test-Path $AddInPath)) {
    throw "Add-in not found: $AddInPath. Run 'dotnet build' first."
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject Excel.Application
$app.Visible = -not $Hidden
$app.UserControl = $true
$null = $app.Workbooks.Add()
if (-not $app.RegisterXLL($AddInPath)) {
    throw "Excel refused to register $AddInPath"
}

$status = $app.Run('vbang.Status')
$pid_ = Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1 -ExpandProperty Id
Write-Output "Excel PID: $pid_"
Write-Output "Add-in: $AddInPath"
Write-Output $status
[Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
