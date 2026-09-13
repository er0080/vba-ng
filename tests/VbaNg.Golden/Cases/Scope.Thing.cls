VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Friend members (MS-VBAL 5.2.4.1.1): reachable from every module of the project.

Private hitCount As Long

Friend Function Secret() As String
    Secret = "friend"
End Function

Friend Property Get Hits() As Long
    Hits = hitCount
End Property

Friend Sub Hit()
    hitCount = hitCount + 1
End Sub
