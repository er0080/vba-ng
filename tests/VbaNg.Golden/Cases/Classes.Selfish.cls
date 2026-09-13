VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Selfish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class_Initialize takes and drops a reference to Me (TypeName(Me)), as VBA-Better-Array's
' constructor does; the object must outlive it (ROADMAP.md WP3).

Private Sub Class_Initialize()
    Trace = Trace & "I" & TypeName(Me)
End Sub

Private Sub Class_Terminate()
    Trace = Trace & "T"
End Sub

Public Function Ping() As String
    Ping = "ping"
End Function
