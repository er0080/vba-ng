VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A default property with a Let and a Set accessor both: a late-bound Set reaches the Set, a
' late-bound Let the Let, as VBA-Dictionary's Item needs (ROADMAP.md WP3).

Public LastAccessor As String
Private held As Variant

Public Property Get Item(ByVal key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(held) Then
        Set Item = held
    Else
        Item = held
    End If
End Property

Public Property Let Item(ByVal key As String, ByVal newValue As Variant)
    LastAccessor = "Let " & key
    held = newValue
End Property

Public Property Set Item(ByVal key As String, ByVal newValue As Variant)
    LastAccessor = "Set " & key
    Set held = newValue
End Property

Public Function HandMe() As String
    ReleaseSlot Me
    HandMe = TypeName(Me)
End Function
