VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Members that take fewer arguments than a late-bound call passes them. Calls counts the procedures that ran.

Public Data As Object
Public Items As Variant
Public Tag As String
Public Calls As Long

Public Property Get Prop() As Object
    Set Prop = Data
End Property

Public Property Get Named(ByVal key As String) As Collection
    Calls = Calls + 1
    Set Named = Data
End Property

Public Function Make() As Collection
    Set Make = Data
End Function

Public Function Number() As Long
    Calls = Calls + 1
    Number = 7
End Function

Public Function Pick(ByVal i As Long) As Collection
    Calls = Calls + 1
    Set Pick = Data
End Function

Public Sub Act()
    Calls = Calls + 1
End Sub
