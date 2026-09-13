Attribute VB_Name = "Locals"
Option Explicit

' The locals window check of ARCHITECTURE.md section 10: every kind of VBA storage the memory
' model (D20) keeps in native memory, filled with known values, then Stop. Under F5 ("Debug
' Storage sample in Excel") the debugger breaks on the Stop and the Locals pane should read
' each variable as VBA's locals window would.

Private Type Part
    Code As String * 6
    Name As String
    Count As Long
    Price As Currency
    Active As Boolean
    Owner As Account
    Extra As Variant
End Type

Private moduleTotal As Double
Private moduleParts(1 To 2) As Part

Public Sub Inspect()
    Static calls As Long
    Dim i As Long
    Dim total As Double
    Dim when As Date
    Dim price As Currency
    Dim flag As Boolean
    Dim text As String
    Dim fixed As String * 4
    Dim number As Variant
    Dim word As Variant
    Dim list As Variant
    Dim holder As Variant
    Dim numbers() As Long
    Dim grid(1 To 2, 1 To 2) As String
    Dim acct As Account
    Dim late As Object
    Dim items As Collection
    Dim part As Part

    calls = calls + 1
    For i = 1 To 3
        total = total + i * 1.5
    Next
    when = #9/11/2026 2:30:00 PM#
    price = 19.99
    flag = True
    text = "hello"
    fixed = "ab"
    number = 42
    word = "forty-two"
    list = Array(1, "two", 3.5)
    ReDim numbers(1 To 3)
    numbers(1) = 10: numbers(2) = 20: numbers(3) = 30
    grid(1, 1) = "a": grid(2, 1) = "b": grid(1, 2) = "c": grid(2, 2) = "d"
    Set acct = New Account
    acct.Deposit 100
    acct.Deposit 25
    Set late = acct
    Set holder = acct
    Set items = New Collection
    items.Add "one"
    items.Add 2
    part.Code = "X-1"
    part.Name = "widget"
    part.Count = 7
    part.Price = 2.5
    part.Active = True
    Set part.Owner = acct
    part.Extra = "note"
    moduleTotal = total
    moduleParts(1) = part
    Stop
    Debug.Print calls; total; when; price; flag; text; fixed; number; word; list(1); numbers(2); grid(2, 2); _
        acct.Balance; late.Balance; items.Count; part.Name; moduleParts(1).Code
End Sub
