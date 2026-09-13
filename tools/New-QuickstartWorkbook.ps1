<#
.SYNOPSIS
Creates samples\Quickstart\Quickstart.xlsx: a macro-free workbook with one Form control button,
the workbook the README quickstart walks through (CLAUDE.md R17, R31).

.DESCRIPTION
The button's OnAction names Greeting.SayHello from the Quickstart.vbang project next to the
workbook; clicking it writes "Hello, world" to B2 and prints the same line. Runs a throwaway
hidden Excel through COM and quits it. Windows PowerShell 5.1 compatible (CLAUDE.md R22).

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\New-QuickstartWorkbook.ps1
#>
param(
    [string]$Path = (Join-Path $PSScriptRoot '..\samples\Quickstart\Quickstart.xlsx')
)

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject Excel.Application
$pid_ = Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1 -ExpandProperty Id
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Sheet1'
    $ws.Range('A2').Value2 = 'Greeting'
    $ws.Range('A2').Font.Bold = $true

    $button = $ws.Buttons().Add(240, 10, 100, 30)
    $button.Name = 'SayHelloButton'
    $button.Caption = 'Say hello'
    $button.OnAction = 'Greeting.SayHello'

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $wb.SaveAs($Path, 51)   # xlOpenXMLWorkbook
    $wb.Close($false)
    Write-Output "Wrote $Path"
}
finally {
    $app.Quit()
    foreach ($object in @($button, $ws, $wb, $app)) {
        if ($null -ne $object) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) | Out-Null }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # The COM references above keep Excel alive until they are gone; a throwaway that still lingers is ended (CLAUDE.md R17).
    if ($pid_ -and (Get-Process -Id $pid_ -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 2
        Stop-Process -Id $pid_ -Force -ErrorAction SilentlyContinue
    }
}
