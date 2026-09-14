VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phoenix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class_Terminate stores Me when Revive is set, as VBA-TDD's SpecDefinition hands itself to its suite.

Public Tag As String
Public Child As Collection

Private Sub Class_Initialize()
    Set Child = New Collection
    Child.Add "kept"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "T" & Tag
    If Revive Then Keep.Add Me
End Sub
