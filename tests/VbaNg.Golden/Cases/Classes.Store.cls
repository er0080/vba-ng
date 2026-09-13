VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A keyed store whose default member has both a Get and a Let, so s(key) = value assigns through it.

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Public Property Get Item(ByVal key As String) As Variant
Attribute Item.VB_UserMemId = 0
    Item = items(key)
End Property

Public Property Let Item(ByVal key As String, ByVal value As Variant)
    items.Add value, key
End Property
