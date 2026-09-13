Attribute VB_Name = "Bench"
Option Explicit

' The performance baseline of ROADMAP.md WP7: four loops timed the same way under vbang and
' under VBA. Each function returns its elapsed milliseconds and checks its own result, so a
' wrong answer raises instead of being timed. Main prints the four lines;
' tests/VbaNg.E2E/BenchmarkTests.cs runs Main under vbang.Run with Benchmarks.xlsx bound (so the
' UDF is registered) and calls the same functions with this module imported into a VBA project,
' then writes benchmark-report.txt with the ratios, which ROADMAP.md publishes.
'
' Native VBA on the development machine, Excel 16.0 x64, best of two runs on 2026-09-09:
'   Variant loop               61 ms
'   String building            81 ms
'   Range loop              2,936 ms
'   UDF over 10,000 cells     302 ms
' and, added on 2026-09-11 for ROADMAP.md D-L:
'   Late-bound Collection      93 ms
'   Late-bound Excel          115 ms
' and, added on 2026-09-11 for ROADMAP.md D-M:
'   Double arithmetic          54 ms

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef count As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long

Public Sub Main()
    Report "Variant loop", VariantLoop()
    Report "Double arithmetic", DoubleArithmetic()
    Report "String building", StringBuilding()
    Report "Range loop", RangeLoop()
    Report "UDF over 10,000 cells", UdfCells()
    Report "Late-bound Collection", LateBoundCollection()
    Report "Late-bound Excel", LateBoundExcel()
End Sub

' A million iterations of Variant assignment and arithmetic: the scalar hot path of CLAUDE.md R20.
Public Function VariantLoop() As Double
    Dim started As Currency
    Dim i As Long
    Dim v As Variant
    Dim total As Variant
    started = Ticks()
    total = 0#
    For i = 1 To 1000000
        v = i
        total = total + v * 2 - 1
    Next i
    VariantLoop = Elapsed(started)
    If total <> 1E12 Then Fail "Variant loop"
End Function

' A million iterations of typed Double arithmetic, four operations each and a running total: the
' path ROADMAP.md D-M evaluates in the x87 extended format, as VBA does.
Public Function DoubleArithmetic() As Double
    Dim started As Currency
    Dim i As Long
    Dim x As Double
    Dim y As Double
    Dim total As Double
    started = Ticks()
    y = 3
    For i = 1 To 1000000
        x = i
        total = total + x / y * 1.5 - x * 0.5
    Next i
    DoubleArithmetic = Elapsed(started)
    If Abs(total) > 1 Then Fail "Double arithmetic"
End Function

' Ten thousand concatenations, then the result scanned twice, once with InStr and once with Mid$.
Public Function StringBuilding() As Double
    Dim started As Currency
    Dim i As Long
    Dim s As String
    Dim commas As Long
    Dim p As Long
    started = Ticks()
    For i = 1 To 10000
        s = s & CStr(i) & ","
    Next i
    p = InStr(s, ",")
    Do While p > 0
        commas = commas + 1
        p = InStr(p + 1, s, ",")
    Loop
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "," Then commas = commas + 1
    Next i
    StringBuilding = Elapsed(started)
    If commas <> 20000 Or Len(s) <> 48894 Then Fail "String building"
End Function

' A hundred by a hundred cells written one at a time and read back one at a time: twenty
' thousand calls into Excel's object model.
Public Function RangeLoop() As Double
    Dim ws As Worksheet
    Dim started As Currency
    Dim r As Long
    Dim c As Long
    Dim total As Double
    Set ws = ActiveSheet
    ws.Cells.ClearContents
    started = Ticks()
    For r = 1 To 100
        For c = 1 To 100
            ws.Cells(r, c).Value = r * c
        Next c
    Next r
    For r = 1 To 100
        For c = 1 To 100
            total = total + ws.Cells(r, c).Value
        Next c
    Next r
    RangeLoop = Elapsed(started)
    If total <> 25502500 Then Fail "Range loop"
End Function

' The function the formulas call.
Public Function BenchUdf(x As Double) As Double
    BenchUdf = x * 2 + 1
End Function

' Ten thousand cells whose formula calls BenchUdf, recalculated in full.
Public Function UdfCells() As Double
    Dim ws As Worksheet
    Dim started As Currency
    Set ws = ActiveSheet
    ws.Cells.ClearContents
    ws.Range("A1:J1000").Formula = "=BenchUdf(ROW())"
    started = Ticks()
    Application.CalculateFull
    UdfCells = Elapsed(started)
    If ws.Evaluate("SUM(A1:J1000)") <> 10020000 Then Fail "UDF over 10,000 cells"
End Function

' A hundred thousand calls into a Collection through an Object variable, so each one is bound
' by name at run time (ROADMAP.md D-L): fifty thousand Adds, then fifty thousand Counts.
Public Function LateBoundCollection() As Double
    Dim started As Currency
    Dim items As Object
    Dim i As Long
    Dim total As Double
    Set items = New Collection
    started = Ticks()
    For i = 1 To 50000
        items.Add i
    Next i
    For i = 1 To 50000
        total = total + items.Count
    Next i
    LateBoundCollection = Elapsed(started)
    If total <> 2500000000# Then Fail "Late-bound Collection"
End Function

' Twenty thousand property reads from a worksheet through an Object variable: late binding into
' Excel's object model, where each call is a name lookup and an IDispatch Invoke (ROADMAP.md D-L).
Public Function LateBoundExcel() As Double
    Dim started As Currency
    Dim sheet As Object
    Dim i As Long
    Dim total As Long
    Set sheet = ActiveSheet
    started = Ticks()
    For i = 1 To 20000
        total = total + Len(sheet.Name)
    Next i
    LateBoundExcel = Elapsed(started)
    If total <> 20000 * Len(ActiveSheet.Name) Then Fail "Late-bound Excel"
End Function

Private Function Ticks() As Currency
    Dim count As Currency
    QueryPerformanceCounter count
    Ticks = count
End Function

Private Function Elapsed(started As Currency) As Double
    Dim frequency As Currency
    QueryPerformanceFrequency frequency
    Elapsed = (Ticks() - started) * 1000 / frequency
End Function

Private Sub Report(name As String, ms As Double)
    Debug.Print name & ": " & Format(ms, "0") & " ms"
End Sub

Private Sub Fail(name As String)
    Err.Raise vbObjectError + 513, "Bench", name & " computed the wrong result"
End Sub
