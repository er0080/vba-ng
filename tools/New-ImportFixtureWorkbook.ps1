<#
.SYNOPSIS
Creates tests\VbaNg.Import.Tests\Fixtures\Legacy.xlsm: a macro-enabled workbook whose VBA project
holds a standard module, a class module, a document module, a form, a reference, and a Declare,
with a Form button and an ActiveX button on the sheet: the fixture the MS-OVBA importer and the
import round trip are tested against (CLAUDE.md R17, R31).

.DESCRIPTION
Writes the modules through the VBE object model, which needs "Trust access to the VBA project
object model" (Excel: Trust Center, Macro Settings). Runs a throwaway hidden Excel through COM
and quits it. Windows PowerShell 5.1 compatible (CLAUDE.md R22).

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\New-ImportFixtureWorkbook.ps1
#>
param(
    [string]$Path = (Join-Path $PSScriptRoot '..\tests\VbaNg.Import.Tests\Fixtures\Legacy.xlsm')
)

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$null = New-Item -ItemType Directory -Force -Path (Split-Path $Path)

$standard = @'
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)

Public Const TaxRate As Double = 0.2

Public Function Gross(ByVal net As Currency) As Currency
    Gross = net * (1 + TaxRate)
End Function

Public Sub Wait()
    Sleep 0
End Sub

Public Sub ButtonClick()
    Worksheets("Sheet1").Range("D1").Value = "form button"
End Sub
'@

$class = @'
Public Total As Currency

Private Sub Class_Initialize()
    Total = 0
End Sub

Public Sub Add(ByVal amount As Currency)
    Total = Total + amount
End Sub
'@

$sheet = @'
Private Sub Worksheet_Change(ByVal Target As Range)
    ' The write below raises Change again; without this guard Workbook_Open takes Excel down.
    Application.EnableEvents = False
    Range("B1").Value = "changed"
    Application.EnableEvents = True
End Sub

Private Sub CommandButton1_Click()
    Range("C1").Value = "activex"
End Sub
'@

$workbook = @'
Private Sub Workbook_Open()
    Worksheets(1).Range("A1").Value = "opened"
End Sub
'@

$form = @'
Private Sub UserForm_Click()
    Me.Caption = "clicked"
End Sub
'@

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject Excel.Application
$pid_ = Get-Process EXCEL | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1 -ExpandProperty Id
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Add()
    $wb.Worksheets.Item(1).Name = 'Sheet1'
    $project = $wb.VBProject
    $project.Name = 'LegacyBook'

    # A registered reference the importer must find in the dir stream: the Scripting runtime.
    $null = $project.References.AddFromGuid('{420B2830-E718-11CF-893D-00A0C9054228}', 1, 0)

    $module = $project.VBComponents.Add(1)      # vbext_ct_StdModule
    $module.Name = 'Sales'
    $module.CodeModule.AddFromString($standard)

    $customer = $project.VBComponents.Add(2)    # vbext_ct_ClassModule
    $customer.Name = 'Basket'
    $customer.CodeModule.AddFromString($class)

    $userForm = $project.VBComponents.Add(3)    # vbext_ct_MSForm
    $userForm.Name = 'Dialog'
    $userForm.CodeModule.AddFromString($form)

    $project.VBComponents.Item('Sheet1').CodeModule.AddFromString($sheet)
    $project.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString($workbook)

    # The two ways a button reaches code: a Form control, whose OnAction names a macro, and an
    # ActiveX control, whose events bind by name in the sheet's module (ROADMAP.md WP6).
    $sheet1 = $wb.Worksheets.Item('Sheet1')
    $formButton = $sheet1.Buttons().Add(200, 10, 100, 30)
    $formButton.OnAction = 'ButtonClick'
    $formButton.Caption = 'Run'
    $missing = [Type]::Missing
    $activeX = $sheet1.OLEObjects().Add('Forms.CommandButton.1', $missing, $missing, $missing, $missing, $missing, $missing, 200, 60, 100, 30)
    $activeX.Name = 'CommandButton1'

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $wb.SaveAs($Path, 52)   # xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    Write-Output "Wrote $Path"
}
finally {
    $app.Quit()
    foreach ($object in @($activeX, $formButton, $sheet1, $userForm, $customer, $module, $project, $wb, $app)) {
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
