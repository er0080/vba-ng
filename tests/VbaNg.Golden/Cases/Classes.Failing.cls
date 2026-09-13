VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Failing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class_Initialize raises an error.

Private Sub Class_Initialize()
    Trace = Trace & "F"
    Err.Raise 1001, , "init failed"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "G"
End Sub
