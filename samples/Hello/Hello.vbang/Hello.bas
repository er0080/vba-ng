Attribute VB_Name = "Hello"
Option Explicit

Public Sub Main()
    Dim i As Long
    Dim total As Long
    For i = 1 To 5
        total = total + i
    Next i
    Debug.Print "Hello from vba-ng"
    Debug.Print "Total = " & total
End Sub
