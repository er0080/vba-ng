Attribute VB_Name = "Bank"
Option Explicit

' The state the classes share and the helpers the tests use.

Public Live As Long

Public Function OpenLedger(ByVal name As String) As Ledger
    Set OpenLedger = New Ledger
    OpenLedger.Name = name
End Function

Public Function TotalOf(ByVal account As IAccount) As Currency
    TotalOf = account.Balance
End Function

Public Function DescribeAny(ByVal item As Variant) As String
    DescribeAny = TypeName(item)
End Function
