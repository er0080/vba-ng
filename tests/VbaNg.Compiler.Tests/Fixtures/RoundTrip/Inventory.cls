VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' An inventory of items keyed by SKU, raising an event when the count changes.

Public Event CountChanged(ByVal NewCount As Long)

Private Type Entry
    Sku As String
    Quantity As Long
    Price As Currency
End Type

Private mEntries() As Entry
Private mCount As Long
Private WithEvents mApp As Excel.Application
Attribute mApp.VB_VarHelpID = -1

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    If Index < 1 Or Index > mCount Then Err.Raise 9
    Item = mEntries(Index).Sku
End Property

Public Property Let Price(ByVal Sku As String, ByVal Value As Currency)
    Dim i As Long
    i = IndexOf(Sku)
    If i = 0 Then
        Err.Raise vbObjectError + 1, "Inventory", "Unknown SKU: " & Sku
    End If
    mEntries(i).Price = Value
End Property

Public Sub Add(ByVal Sku As String, Optional ByVal Quantity As Long = 1, Optional ByVal Price As Currency = 0@)
    Dim i As Long
    i = IndexOf(Sku)
    If i > 0 Then
        mEntries(i).Quantity = mEntries(i).Quantity + Quantity
        Exit Sub
    End If

    mCount = mCount + 1
    ReDim Preserve mEntries(1 To mCount)
    With mEntries(mCount)
        .Sku = Sku
        .Quantity = Quantity
        .Price = Price
    End With
    RaiseEvent CountChanged(mCount)
End Sub

Public Function Total() As Currency
    Dim i As Long, sum As Currency
    For i = 1 To mCount
        sum = sum + mEntries(i).Quantity * mEntries(i).Price
    Next i
    Total = sum
End Function

Private Function IndexOf(ByVal Sku As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mEntries(i).Sku, Sku, vbTextCompare) = 0 Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    ReDim mEntries(1 To 1)
    Set mApp = Application
End Sub

Private Sub Class_Terminate()
    Set mApp = Nothing
End Sub

Private Sub mApp_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print "Opened "; Wb.Name
End Sub
