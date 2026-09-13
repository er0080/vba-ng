VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Sheet events and an ActiveX button: Range and Cells here mean the sheet's own.

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$2" Then
        Range("B1").Value = "changed " & Target.Address & " to " & Target.Value
    End If
End Sub

Private Sub CommandButton1_Click()
    Cells(1, 3).Value = "clicked"
End Sub
