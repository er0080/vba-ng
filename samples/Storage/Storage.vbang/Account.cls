VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A class whose variables lie in its instance's own block (D20): expand an Account in the
' Locals pane to see them.

Private owner As String
Private amounts() As Currency
Private total As Currency
Private lastPosted As Variant

Public Sub Deposit(ByVal amount As Currency)
    owner = "Eric"
    If total = 0 Then ReDim amounts(0 To 0) Else ReDim Preserve amounts(0 To UBound(amounts) + 1)
    amounts(UBound(amounts)) = amount
    total = total + amount
    lastPosted = amount
End Sub

Public Property Get Balance() As Currency
    Balance = total
End Property
