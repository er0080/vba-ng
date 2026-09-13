VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The interface every account implements.

Public Property Get Label() As String
End Property

Public Property Get Balance() As Currency
End Property

Public Sub Post(ByVal amount As Currency)
End Sub
