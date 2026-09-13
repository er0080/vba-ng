VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Outer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Creates a Counter in its Class_Initialize; the Counter goes when Outer does.

Private inner As Counter

Private Sub Class_Initialize()
    Trace = Trace & "O"
    Set inner = New Counter
    inner.Tag = "in"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "P"
End Sub

Public Property Get Inside() As Counter
    Set Inside = inner
End Property
