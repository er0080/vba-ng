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

' A class with a predeclared instance, reached by the class name.

Public Level As Long

Private Sub Class_Initialize()
    Trace = Trace & "S"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "X"
End Sub
