Attribute VB_Name = "Hidden"
Option Explicit
Option Private Module

' Option Private Module (MS-VBAL 4.4): the module's public members stay reachable inside the
' project; only other projects lose them.

Public HiddenCount As Long

Public Function Secret() As String
    Secret = "still visible in the project"
End Function
