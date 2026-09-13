VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Two Nodes can reference each other, which is the cycle the reference count never breaks.

Public Tag As String
Public Other As Node

Private Sub Class_Terminate()
    Trace = Trace & "T" & Tag
End Sub
