VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A collection wrapper with a default Item and a NewEnum for For Each.

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Public Sub Add(ByVal item As Variant)
    items.Add item
End Sub

Public Property Get Count() As Long
    Count = items.Count
End Property

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Item = items(index)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = items.[_NewEnum]
End Function
