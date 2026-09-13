Attribute VB_Name = "Parsing"
Option Explicit

' Error handling as legacy code writes it: On Error GoTo with Exit Function before the
' handler, On Error Resume Next with Err.Number, Resume to a label, and GoSub.

Public Function ParseNumbers(ByVal text As String) As Long()
    Dim parts As Variant
    Dim result() As Long
    Dim i As Long
    parts = Split(text, ",")
    ReDim result(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        result(i) = CLng(Trim$(parts(i)))
    Next i
    ParseNumbers = result
End Function

Public Function SafeDivide(ByVal a As Double, ByVal b As Double, ByRef failed As Boolean) As Double
    On Error GoTo Fail
    failed = False
    SafeDivide = a / b
    Exit Function
Fail:
    failed = True
    SafeDivide = 0
End Function

Public Function ErrorNumberOf(ByVal text As String) As Long
    Dim n As Long
    On Error Resume Next
    n = CLng(text)
    ErrorNumberOf = Err.Number
End Function

Public Function Classify(ByVal n As Long) As String
    Dim result As String
    If n < 0 Then
        GoSub Negative
    Else
        GoSub NonNegative
    End If
    Classify = result
    Exit Function
Negative:
    result = "negative"
    Return
NonNegative:
    If n = 0 Then
        result = "zero"
    Else
        result = "positive"
    End If
    Return
End Function

Public Function RetryCount() As Long
    Dim attempts As Long
    On Error GoTo Handler
Attempt:
    attempts = attempts + 1
    If attempts < 3 Then Err.Raise 1001
    RetryCount = attempts
    Exit Function
Handler:
    Resume Attempt
End Function
