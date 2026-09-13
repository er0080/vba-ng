using System.Globalization;
using System.Text;

using VbaNg.Compiler.Syntax;

namespace VbaNg.Golden.Harness;

/// <summary>The recorder module text for one area, plus the module line on which each case's code starts.</summary>
public sealed record GeneratedModule(string Text, IReadOnlyList<int> CaseLines);

/// <summary>
/// Generates the VBA module that real Excel runs for one area: one procedure per case under an
/// error handler that records <c>Err</c>, and the recorder that describes values as JSON lines.
/// Every <c>?</c> line becomes a call to <c>GoldenRecord</c>, so the expression is evaluated in the
/// case's own procedure, with the case's declared variables, exactly as the author wrote it.
/// </summary>
public static class RecorderModule
{
    public const string RunProcedure = "GoldenRun";
    public const string RecordProcedure = "GoldenRecord";
    private const string ErrorProcedure = "GoldenError";
    private const string FailLabel = "GoldenFail";

    public static string ModuleName(string area) => "Golden" + area;

    /// <summary>A companion module's text as the VBE exported it, CRLF line ends included.</summary>
    public static string CompanionText(CompanionSource source)
    {
        ArgumentNullException.ThrowIfNull(source);
        return string.Join("\r\n", source.Lines) + "\r\n";
    }

    public static GeneratedModule Generate(CaseArea area, string resultsPath)
    {
        ArgumentNullException.ThrowIfNull(area);
        ArgumentException.ThrowIfNullOrWhiteSpace(resultsPath);
        if (resultsPath.Contains('"', StringComparison.Ordinal))
        {
            throw new ArgumentException("The results path cannot contain a double quote.", nameof(resultsPath));
        }

        var builder = new StringBuilder();
        var lineNumber = 0;
        void Line(string text = "")
        {
            builder.Append(text).Append("\r\n");
            lineNumber++;
        }

        foreach (var option in area.Options)
        {
            Line(option);
        }

        if (area.Options.Count > 0)
        {
            Line();
        }

        Line("Private Declare PtrSafe Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)");
        Line();
        Line($"Private Const GoldenPath As String = \"{resultsPath}\"");
        Line();
        foreach (var declaration in area.Declarations)
        {
            Line(declaration);
        }

        if (area.Declarations.Count > 0)
        {
            Line();
        }

        // An error raised inside a case's own handler escapes the case; Resume Next here turns it
        // into that case's recorded error instead of a modal run-time error dialog.
        Line($"Public Sub {RunProcedure}()");
        Line("    On Error Resume Next");
        Line("    Dim goldenFile As Integer: goldenFile = FreeFile");
        Line("    Open GoldenPath For Output As #goldenFile: Close #goldenFile");
        for (var i = 1; i <= area.Cases.Count; i++)
        {
            Line(Invariant($"    GoldenCase{i}"));
            Line(Invariant($"    If Err.Number <> 0 Then {ErrorProcedure} {i}, Err.Number, Err.Description, Err.Source: Err.Clear"));
        }

        Line("    GoldenWrite \"{\"\"done\"\":true}\"");
        Line("End Sub");

        var caseLines = new List<int>();
        for (var i = 0; i < area.Cases.Count; i++)
        {
            var number = i + 1;
            var source = area.Cases[i];
            Line();
            Line(Invariant($"' {source.Name}"));
            Line(Invariant($"Private Sub GoldenCase{number}()"));
            Line($"    On Error GoTo {FailLabel}");
            caseLines.Add(lineNumber + 1);
            var resultIndex = 0;
            foreach (var line in source.Lines)
            {
                if (GoldenCaseSource.IsResultLine(line))
                {
                    Line(Invariant($"    {RecordProcedure} {number}, {resultIndex}, {GoldenCaseSource.ResultExpression(line)}"));
                    resultIndex++;
                }
                else
                {
                    Line("    " + line.Trim());
                }
            }

            Line("    Exit Sub");
            Line($"{FailLabel}:");
            Line(Invariant($"    {ErrorProcedure} {number}, Err.Number, Err.Description, Err.Source"));
            Line("End Sub");
        }

        Line();
        builder.Append(Recorder.ReplaceLineEndings("\r\n"));
        return new GeneratedModule(builder.ToString(), caseLines);
    }

    /// <summary>
    /// Checks the generated module with the vba-ng parser before it goes anywhere near Excel: a
    /// syntax error in one case would otherwise stop the whole area behind a VBE dialog. Returns
    /// one message per problem, naming the case.
    /// </summary>
    public static IReadOnlyList<string> Validate(CaseArea area)
    {
        ArgumentNullException.ThrowIfNull(area);
        var module = Generate(area, @"C:\golden.jsonl");
        var tree = SyntaxTree.Parse(module.Text, area.Name + CaseFile.Extension);
        var problems = new List<string>();

        foreach (var diagnostic in tree.Diagnostics)
        {
            problems.Add($"{Describe(area, module, diagnostic.Line)}: {diagnostic.Message}");
        }

        foreach (var source in area.Companions)
        {
            var companionTree = SyntaxTree.Parse(CompanionText(source), source.FileName);
            foreach (var diagnostic in companionTree.Diagnostics)
            {
                problems.Add(string.Create(CultureInfo.InvariantCulture, $"{area.Name}.{source.FileName}({diagnostic.Line}): {diagnostic.Message}"));
            }
        }

        foreach (var call in tree.Root.DescendantNodes().OfType<CallStatementSyntax>())
        {
            if (call.Expression is IdentifierNameSyntax { Name: RecordProcedure } && call.Arguments?.Arguments.Count != 3)
            {
                var token = call.FirstToken()!;
                var line = tree.GetLinePosition(token).Line;
                problems.Add($"{Describe(area, module, line)}: a ? line records exactly one expression.");
            }
        }

        return problems;
    }

    private static string Describe(CaseArea area, GeneratedModule module, int moduleLine)
    {
        var index = -1;
        for (var i = 0; i < module.CaseLines.Count && module.CaseLines[i] <= moduleLine; i++)
        {
            index = i;
        }

        if (index < 0)
        {
            return Invariant($"{area.Name}{CaseFile.Extension}: recorder line {moduleLine}");
        }

        var source = area.Cases[index];
        var caseLine = source.Line + (moduleLine - module.CaseLines[index]);
        return Invariant($"{area.Name}{CaseFile.Extension}({caseLine}): case '{source.Name}'");
    }

    private static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// The recorder, in VBA. Values are described by <c>TypeName</c> plus a readable <c>Str</c> form and,
    /// where text loses information, the exact bits read with <c>CopyMemory</c>. Strings are escaped
    /// to ASCII JSON; a lone surrogate, which JSON cannot carry, becomes U+FFFD in the value and
    /// the exact UTF-16 units travel in <c>units</c>. Arrays are walked by dimension; the dimension
    /// count comes from the SAFEARRAY header rather than from probing <c>LBound</c>, so describing
    /// a value never disturbs the case's own <c>Err</c> state.
    /// </summary>
    private const string Recorder = """""
        Private Sub GoldenRecord(ByVal caseIndex As Long, ByVal resultIndex As Long, ByVal value As Variant)
            GoldenWrite "{""case"":" & caseIndex & ",""index"":" & resultIndex & ",""value"":" & GoldenDescribe(value) & "}"
        End Sub

        Private Sub GoldenError(ByVal caseIndex As Long, ByVal number As Long, ByVal description As String, ByVal source As String)
            GoldenWrite "{""case"":" & caseIndex & ",""error"":{""number"":" & number & ",""description"":" & GoldenQuote(description) & ",""source"":" & GoldenQuote(source) & "}}"
        End Sub

        ' Each record opens the results file on its own so that a case's Close or Reset cannot end the recording.
        Private Sub GoldenWrite(ByVal text As String)
            Dim goldenFile As Integer
            goldenFile = FreeFile
            Open GoldenPath For Append As #goldenFile
            Print #goldenFile, text
            Close #goldenFile
        End Sub

        Private Function GoldenDescribe(ByRef value As Variant) As String
            If IsObject(value) Then
                If value Is Nothing Then
                    GoldenDescribe = "{""type"":""Nothing""}"
                Else
                    GoldenDescribe = "{""type"":""Object"",""class"":" & GoldenQuote(TypeName(value)) & "}"
                End If
            ElseIf IsArray(value) Then
                GoldenDescribe = GoldenDescribeArray(value)
            Else
                GoldenDescribe = GoldenDescribeScalar(value)
            End If
        End Function

        Private Function GoldenDescribeScalar(ByRef value As Variant) As String
            Dim kind As String
            Dim text As String
            Dim single4 As Single
            Dim double8 As Double
            Dim currency8 As Currency
            Dim bits32 As Long
            Dim bits64 As LongLong
            Dim raw(0 To 15) As Byte
            Dim code As Long
            kind = TypeName(value)
            Select Case VarType(value)
                Case vbEmpty, vbNull
                    GoldenDescribeScalar = "{""type"":" & GoldenQuote(kind) & "}"
                Case vbBoolean
                    If value Then text = "True" Else text = "False"
                    GoldenDescribeScalar = "{""type"":""Boolean"",""value"":""" & text & """}"
                Case vbByte, vbInteger, vbLong, vbLongLong
                    GoldenDescribeScalar = "{""type"":" & GoldenQuote(kind) & ",""value"":""" & Trim$(Str$(value)) & """}"
                Case vbSingle
                    single4 = value
                    CopyMemory VarPtr(bits32), VarPtr(single4), 4
                    GoldenDescribeScalar = "{""type"":""Single"",""value"":""" & Trim$(Str$(value)) & """,""bits"":""" & GoldenHex(bits32, 8) & """}"
                Case vbDouble
                    double8 = value
                    CopyMemory VarPtr(bits64), VarPtr(double8), 8
                    GoldenDescribeScalar = "{""type"":""Double"",""value"":""" & Trim$(Str$(value)) & """,""bits"":""" & GoldenHex(bits64, 16) & """}"
                Case vbDate
                    double8 = value
                    CopyMemory VarPtr(bits64), VarPtr(double8), 8
                    GoldenDescribeScalar = "{""type"":""Date"",""value"":""" & Format$(value, "yyyy\-mm\-dd hh:nn:ss") & """,""bits"":""" & GoldenHex(bits64, 16) & """}"
                Case vbCurrency
                    currency8 = value
                    CopyMemory VarPtr(bits64), VarPtr(currency8), 8
                    GoldenDescribeScalar = "{""type"":""Currency"",""value"":""" & Trim$(Str$(value)) & """,""bits"":""" & GoldenHex(bits64, 16) & """}"
                Case vbDecimal
                    CopyMemory VarPtr(raw(0)), VarPtr(value), 16
                    GoldenDescribeScalar = "{""type"":""Decimal"",""value"":""" & Trim$(Str$(value)) & """,""bytes"":""" & GoldenBytes(raw) & """}"
                Case vbError
                    CopyMemory VarPtr(code), VarPtr(value) + 8, 4
                    GoldenDescribeScalar = "{""type"":""Error"",""value"":""" & Trim$(Str$(code)) & """}"
                Case vbString
                    GoldenDescribeScalar = "{""type"":""String"",""value"":" & GoldenQuote(value)
                    If GoldenHasLoneSurrogate(value) Then GoldenDescribeScalar = GoldenDescribeScalar & ",""units"":""" & GoldenUnits(value) & """"
                    GoldenDescribeScalar = GoldenDescribeScalar & "}"
                Case Else
                    GoldenDescribeScalar = "{""type"":" & GoldenQuote(kind) & ",""value"":""" & VarType(value) & """}"
            End Select
        End Function

        Private Function GoldenDescribeArray(ByRef value As Variant) As String
            Dim dimensions As Long
            Dim i As Long
            Dim j As Long
            Dim bounds As String
            Dim items As String
            Dim separator As String
            dimensions = GoldenDimensions(value)
            For i = 1 To dimensions
                If i > 1 Then bounds = bounds & ", "
                bounds = bounds & LBound(value, i) & " To " & UBound(value, i)
            Next
            GoldenDescribeArray = "{""type"":" & GoldenQuote(TypeName(value)) & ",""bounds"":""" & bounds & """"
            If dimensions = 1 Then
                For i = LBound(value) To UBound(value)
                    items = items & separator & GoldenDescribe(value(i))
                    separator = ","
                Next
                GoldenDescribeArray = GoldenDescribeArray & ",""items"":[" & items & "]"
            ElseIf dimensions = 2 Then
                For i = LBound(value, 1) To UBound(value, 1)
                    For j = LBound(value, 2) To UBound(value, 2)
                        items = items & separator & GoldenDescribe(value(i, j))
                        separator = ","
                    Next
                Next
                GoldenDescribeArray = GoldenDescribeArray & ",""items"":[" & items & "]"
            End If
            GoldenDescribeArray = GoldenDescribeArray & "}"
        End Function

        Private Function GoldenDimensions(ByRef value As Variant) As Long
            Dim vt As Integer
            Dim pointer As LongPtr
            Dim dimensions As Integer
            CopyMemory VarPtr(vt), VarPtr(value), 2
            CopyMemory VarPtr(pointer), VarPtr(value) + 8, LenB(pointer)
            If pointer = 0 Then Exit Function
            If (vt And &H4000) <> 0 Then
                CopyMemory VarPtr(pointer), pointer, LenB(pointer)
                If pointer = 0 Then Exit Function
            End If
            CopyMemory VarPtr(dimensions), pointer, 2
            GoldenDimensions = dimensions
        End Function

        Private Function GoldenHex(ByVal number As LongLong, ByVal digits As Long) As String
            GoldenHex = Right$(String$(digits, "0") & Hex$(number), digits)
        End Function

        Private Function GoldenBytes(ByRef raw() As Byte) As String
            Dim i As Long
            For i = LBound(raw) To UBound(raw)
                GoldenBytes = GoldenBytes & Right$("0" & Hex$(raw(i)), 2)
            Next
        End Function

        Private Function GoldenQuote(ByVal text As String) As String
            Dim i As Long
            Dim code As Long
            Dim nextCode As Long
            Dim result As String
            result = """"
            i = 1
            Do While i <= Len(text)
                code = AscW(Mid$(text, i, 1)) And &HFFFF&
                Select Case code
                    Case 34
                        result = result & "\"""
                    Case 92
                        result = result & "\\"
                    Case 32 To 33, 35 To 91, 93 To 126
                        result = result & Chr$(code)
                    Case &HD800& To &HDBFF&
                        nextCode = 0
                        If i < Len(text) Then nextCode = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
                        If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                            result = result & "\u" & Right$("000" & Hex$(code), 4) & "\u" & Right$("000" & Hex$(nextCode), 4)
                            i = i + 1
                        Else
                            result = result & "\u" & "FFFD"
                        End If
                    Case &HDC00& To &HDFFF&
                        result = result & "\u" & "FFFD"
                    Case Else
                        result = result & "\u" & Right$("000" & Hex$(code), 4)
                End Select
                i = i + 1
            Loop
            GoldenQuote = result & """"
        End Function

        Private Function GoldenHasLoneSurrogate(ByVal text As String) As Boolean
            Dim i As Long
            Dim code As Long
            i = 1
            Do While i <= Len(text)
                code = AscW(Mid$(text, i, 1)) And &HFFFF&
                If code >= &HD800& And code <= &HDBFF& Then
                    If i = Len(text) Then GoldenHasLoneSurrogate = True: Exit Function
                    code = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
                    If code < &HDC00& Or code > &HDFFF& Then GoldenHasLoneSurrogate = True: Exit Function
                    i = i + 1
                ElseIf code >= &HDC00& And code <= &HDFFF& Then
                    GoldenHasLoneSurrogate = True
                    Exit Function
                End If
                i = i + 1
            Loop
        End Function

        Private Function GoldenUnits(ByVal text As String) As String
            Dim i As Long
            For i = 1 To Len(text)
                GoldenUnits = GoldenUnits & Right$("000" & Hex$(AscW(Mid$(text, i, 1)) And &HFFFF&), 4)
            Next
        End Function
        """"";
}
