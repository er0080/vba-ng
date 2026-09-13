VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' An account that records every posting and raises an event on each one.

Implements IAccount

Public Event Posted(ByVal amount As Currency, ByRef cancel As Boolean)

Private nameStore As String
Private entries As Collection

Private Sub Class_Initialize()
    Set entries = New Collection
    nameStore = "ledger"
End Sub

Public Property Get Name() As String
    Name = nameStore
End Property

Public Property Let Name(ByVal newName As String)
    nameStore = newName
End Property

Public Property Get Count() As Long
    Count = entries.Count
End Property

Public Property Get Entry(ByVal index As Long) As Currency
Attribute Entry.VB_UserMemId = 0
    Entry = entries(index)
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = entries.[_NewEnum]
End Function

Public Sub Post(ByVal amount As Currency)
    Dim cancel As Boolean
    RaiseEvent Posted(amount, cancel)
    If Not cancel Then entries.Add amount
End Sub

Public Sub Clear()
    Set entries = New Collection
End Sub

Private Property Get IAccount_Label() As String
    IAccount_Label = "ledger " & nameStore
End Property

Private Property Get IAccount_Balance() As Currency
    Dim total As Currency
    Dim i As Long
    For i = 1 To entries.Count
        total = total + entries(i)
    Next
    IAccount_Balance = total
End Property

Private Sub IAccount_Post(ByVal amount As Currency)
    Post amount
End Sub
