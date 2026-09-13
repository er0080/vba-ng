VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Logs its lifetime to the recorder module's Trace, so the object round trip through a LongPtr
' in the Memory golden shows when Class_Terminate fires.

Public Id As Long

Private Sub Class_Initialize()
    Id = 7
    Trace = Trace & "I"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "T"
End Sub
