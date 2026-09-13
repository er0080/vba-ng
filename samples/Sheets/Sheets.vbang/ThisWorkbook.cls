VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Workbook events: Open runs when the add-in binds the workbook (ARCHITECTURE.md D16).
' The WithEvents variables watch the workbook itself, the Application, and Sheet1
' through a class (F1, F2, F3).

Private WithEvents Book As Workbook
Private WithEvents App As Application
Private watcher As Watcher

Private Sub Workbook_Open()
    Worksheets("Sheet1").Range("A1").Value = "opened"
    Set Book = Me
    Set App = Application
    Set watcher = New Watcher
    watcher.Watch Worksheets("Sheet1")
End Sub

Private Sub Book_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' A two-argument event, so the order Excel passes them in is checked too; guarded like Sheet1's handler.
    If Target.Address = "$A$2" Then Worksheets("Sheet1").Range("F2").Value = Sh.Name & "!" & Target.Address(False, False)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Worksheets("Sheet1").Range("A3").Value = "closing"
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Application events reach a document module too (ROADMAP.md WP4), guarded like the handlers above.
    If Target.Address = "$A$2" Then Worksheets("Sheet1").Range("F3").Value = "app " & Sh.Name
End Sub
