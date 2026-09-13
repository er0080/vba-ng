VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The workhorse of the Classes golden: logs its lifetime to Trace, exposes a default
' property, properties of every kind, two events, and Me.

Public Event Changed(ByVal newValue As Long)
Public Event Adjust(ByRef delta As Long)

Public Tag As String
Private valueStore As Long
Private childStore As Counter

Private Sub Class_Initialize()
    Trace = Trace & "I"
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "T" & Tag
End Sub

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = valueStore
End Property

Public Property Let Value(ByVal newValue As Long)
    valueStore = newValue
    RaiseEvent Changed(newValue)
End Property

Public Property Get Child() As Counter
    Set Child = childStore
End Property

Public Property Set Child(ByVal newChild As Counter)
    Set childStore = newChild
End Property

Public Property Get Doubled(ByVal times As Long) As Long
    Doubled = valueStore * times
End Property

Public Property Let Doubled(ByVal times As Long, ByVal total As Long)
    valueStore = total \ times
End Property

Public Function Self() As Counter
    Set Self = Me
End Function

Public Sub Bump(Optional ByVal by As Long = 1)
    valueStore = valueStore + by
    RaiseEvent Changed(valueStore)
End Sub

Public Function Nudge() As Long
    Dim delta As Long
    delta = 1
    RaiseEvent Adjust(delta)
    valueStore = valueStore + delta
    Nudge = delta
End Function

Public Sub Drain(ByRef target As Long)
    target = valueStore
    valueStore = 0
End Sub

Public Function Describe(ParamArray parts() As Variant) As String
    Dim i As Long
    Describe = Tag & ":" & valueStore
    For i = LBound(parts) To UBound(parts)
        Describe = Describe & "," & parts(i)
    Next
End Function
