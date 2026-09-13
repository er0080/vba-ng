VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A default member that takes a string, which the ! operator names.

Public Function Item(ByVal key As String) As String
Attribute Item.VB_UserMemId = 0
    Item = UCase(key)
End Function
