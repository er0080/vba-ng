VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A predeclared instance: Settings.Unit reads without New.

Public Unit As String

Private Sub Class_Initialize()
    Unit = "EUR"
End Sub
