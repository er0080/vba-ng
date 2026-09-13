Attribute VB_Name = "Utils"
Option Explicit
Option Compare Text
Option Base 1

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Enum LogLevel
    llDebug = 0
    llInfo
    llWarning
    llError = 10
    [_Last] = llError
End Enum

Public Type Settings
    Name As String * 32
    Level As LogLevel
    Thresholds(1 To 3) As Double
End Type

Private Const LOG_FILE As String = "log.txt"
Public Const MAX_RETRIES = 3, TIMEOUT_MS As Long = 30000&

Public Sub Log(ByVal Level As LogLevel, ByVal Message As String, ParamArray Args() As Variant)
    Dim fn As Integer, i As Long, text As String
    text = Message
    For i = LBound(Args) To UBound(Args)
        text = Replace(text, "{" & i - LBound(Args) & "}", CStr(Args(i)))
    Next i

    Select Case Level
        Case llDebug
            Debug.Print "[debug] "; text
        Case llInfo, llWarning
            Debug.Print "[info] " & text
        Case Is >= llError
            fn = FreeFile
            Open ThisWorkbook.Path & "\" & LOG_FILE For Append As #fn
            Print #fn, Format$(Now, "yyyy-mm-dd hh:nn:ss"); Tab(22); text
            Close #fn
        Case Else
            ' ignore
    End Select
End Sub

Public Function Retry(ByVal Action As String, Optional ByVal Attempts As Long = MAX_RETRIES) As Boolean
    Dim attempt As Long
    On Error GoTo Failed
    Do
        attempt = attempt + 1
        Application.Run Action
        Retry = True
        Exit Function
Failed:
        If attempt >= Attempts Then Exit Do
        Sleep 100 * attempt
        Resume Next
    Loop While attempt < Attempts
    On Error GoTo 0
End Function

Public Function ReadLines(ByVal PathName As String) As String()
    Dim fn As Integer, line As String, lines() As String, n As Long
    fn = FreeFile
    Open PathName For Input Access Read Shared As #fn
    Do Until EOF(fn)
        Line Input #fn, line
        ReDim Preserve lines(0 To n)
        lines(n) = line
        n = n + 1
    Loop
    Close #fn
    ReadLines = lines
End Function

Public Function Clamp(ByVal Value As Double, ByVal Low As Double, ByVal High As Double) As Double
    If Value < Low Then
        Clamp = Low
    ElseIf Value > High Then
        Clamp = High
    Else
        Clamp = Value
    End If
End Function

Public Sub Demo()
    Dim s As Settings, cell As Range, d As Date, elapsed As Long
    s.Name = "default"
    s.Level = llInfo
    s.Thresholds(1) = 0.5: s.Thresholds(2) = 1.5: s.Thresholds(3) = 2.5
    d = #1/1/2000 12:00:00 PM#
    elapsed = GetTickCount()
    With ActiveSheet
        For Each cell In .Range("A1:A10").Cells
            If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
                cell.Offset(0, 1).Value = Clamp(cell.Value, s.Thresholds(1), s.Thresholds(3))
            End If
        Next cell
        .Range("B1").Select
    End With
    Log llInfo, "Took {0} ms for {1}", GetTickCount() - elapsed, [A1].Address
    If Retry("Utils.Demo") Then Log llDebug, "ok" Else Log llError, "failed"
    Mid$(s.Name, 1, 3) = "DEF"
    On Error Resume Next
    Name "old.txt" As "new.txt"
    Kill "temp.txt"
    Err.Clear
End Sub
