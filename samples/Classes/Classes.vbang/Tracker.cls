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

' Counts the instances that are alive, so the tests can see Class_Terminate run.

Private Sub Class_Initialize()
    Live = Live + 1
End Sub

Private Sub Class_Terminate()
    Live = Live - 1
End Sub
