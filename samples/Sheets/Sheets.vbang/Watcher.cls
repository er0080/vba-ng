VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Watcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' WithEvents on a library type: Set advises the sheet's events, and the handler runs when a cell changes.

Private WithEvents ws As Worksheet

Public Sub Watch(target As Worksheet)
    Set ws = target
End Sub

Private Sub ws_Change(ByVal Target As Range)
    ' Writing a cell fires Change again, so only the cell the test changes is watched.
    If Target.Address = "$A$2" Then ws.Range("F1").Value = "watched " & Target.Address
End Sub
