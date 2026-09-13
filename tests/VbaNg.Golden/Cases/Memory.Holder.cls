VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Holder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A class variable the Memory golden reaches through its address.

Private m_value As Long

Public Function Address() As LongPtr
    Address = VarPtr(m_value)
End Function

Public Property Get Value() As Long
    Value = m_value
End Property
