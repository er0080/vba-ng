VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Watcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Handles a Counter's events through WithEvents.

Private WithEvents source As Counter
Attribute source.VB_VarHelpID = -1

Public Property Get Target() As Counter
    Set Target = source
End Property

Public Property Set Target(ByVal newTarget As Counter)
    Set source = newTarget
End Property

Private Sub source_Changed(ByVal newValue As Long)
    Trace = Trace & "C" & newValue
End Sub

Private Sub source_Adjust(ByRef delta As Long)
    delta = delta * 10
End Sub
