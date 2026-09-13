VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Auditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Watches a ledger: counts its postings and cancels the ones over its limit.

Private WithEvents source As Ledger
Attribute source.VB_VarHelpID = -1

Public Seen As Long
Public Limit As Currency

Private Sub Class_Initialize()
    Limit = 1000
End Sub

Public Property Get Target() As Ledger
    Set Target = source
End Property

Public Property Set Target(ByVal newTarget As Ledger)
    Set source = newTarget
End Property

Private Sub source_Posted(ByVal amount As Currency, ByRef cancel As Boolean)
    Seen = Seen + 1
    If amount > Limit Then cancel = True
End Sub
