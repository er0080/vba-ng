Attribute VB_Name = "Helpers"
Option Explicit

' The other side of the Scope golden: public members the recorder module reaches bare and as
' Helpers.Name, a member sharing its name with one in Other, a function named like the module
' Other, a property and a Static procedure in a standard module, and a private variable the
' accessor counts.

Public Type Point
    X As Long
    Y As Long
End Type

Public Enum Colors
    Red = 1
    Green = 2
End Enum

Public Const Limit As Long = 100
Public Const Tag As String = "Helpers"
Global GlobalLimit As Long
Public Counter As Long
Public Values(1 To 3) As Long
Private ownCount As Long
Private settingValue As Long

Public Function Twice(n As Long) As Long
    Twice = n * 2
End Function

Public Function Both() As String
    Both = "Helpers"
End Function

Public Function CallsBoth() As String
    CallsBoth = Both()
End Function

Public Function Nearby() As String
    Nearby = "Helpers.Nearby"
End Function

Public Function Other() As String
    Other = "Helpers.Other"
End Function

Public Function Own() As Long
    ownCount = ownCount + 1
    Own = ownCount
End Function

Public Property Get Setting() As Long
    Setting = settingValue
End Property

Public Property Let Setting(value As Long)
    settingValue = value * 10
End Property

Public Static Function StaticCalls() As Long
    Dim n As Long
    n = n + 1
    StaticCalls = n
End Function

Public Sub Reset()
    Counter = 0
End Sub
