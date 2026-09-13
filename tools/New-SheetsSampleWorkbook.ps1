<#
.SYNOPSIS
Creates samples\Sheets\Sheets.xlsx: a macro-free workbook with a Form control button and an
ActiveX CommandButton on Sheet1, the fixture of the M4 lifecycle E2E test (CLAUDE.md R17, R31).

.DESCRIPTION
The button's OnAction names Module1.ButtonClick from the Sheets.vbang project next to the
workbook; the ActiveX control is named CommandButton1, whose Click handler lives in Sheet1.cls.
Runs a throwaway hidden Excel through COM and quits it. Windows PowerShell 5.1 compatible
(CLAUDE.md R22).

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\New-SheetsSampleWorkbook.ps1
#>
param(
    [string]$Path = (Join-Path $PSScriptRoot '..\samples\Sheets\Sheets.xlsx')
)

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$missing = [Type]::Missing
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject Excel.Application
$pid_ = Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1 -ExpandProperty Id
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Sheet1'

    $button = $ws.Buttons().Add(150, 10, 100, 30)
    $button.Name = 'FormButton'
    $button.Caption = 'Form button'
    $button.OnAction = 'Module1.ButtonClick'

    $ole = $ws.OLEObjects().Add('Forms.CommandButton.1', $missing, $missing, $missing, $missing, $missing, $missing, 150, 60, 100, 30)
    $ole.Name = 'CommandButton1'
    $ole.Object.Caption = 'ActiveX button'

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $wb.SaveAs($Path, 51)   # xlOpenXMLWorkbook
    $wb.Close($false)
    Write-Output "Wrote $Path"
}
finally {
    $app.Quit()
    foreach ($object in @($ole, $button, $ws, $wb, $app)) {
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
