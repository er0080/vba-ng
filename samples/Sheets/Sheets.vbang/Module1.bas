Attribute VB_Name = "Module1"
Option Explicit

' The Form control button's OnAction is Module1.ButtonClick; Auto_Open runs after Workbook_Open.

Public Sub ButtonClick()
    ActiveSheet.Range("D1").Value = "button"
End Sub

Public Sub Auto_Open()
    Worksheets(1).Range("E1").Value = "auto"
End Sub

' UDFs: a typed parameter receives the cell's value, a Variant parameter the Range itself, as in VBA.

Public Function Twice(x As Double) As Double
    Twice = x * 2
End Function

Public Function Describe(cell As Variant) As String
    Describe = TypeName(cell) & " " & cell.Address(False, False) & " " & cell.Value
End Function

Public Function Failing() As Double
    Failing = 1 / 0
End Function

' Application.Volatile, Caller, and ThisCell inside a UDF (ROADMAP.md WP4): COM cannot answer
' these while Excel recalculates, so the host answers them from the call Excel is making.

Public Function VolatileCount() As Double
    Static evaluations As Double
    Application.Volatile
    evaluations = evaluations + 1
    VolatileCount = evaluations
End Function

Public Function CallerAddress() As String
    CallerAddress = Application.Caller.Address(False, False)
End Function

Public Function ThisCellAddress() As String
    ThisCellAddress = Application.ThisCell.Address(False, False)
End Function
