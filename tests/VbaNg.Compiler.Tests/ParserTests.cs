using VbaNg.Compiler.Syntax;

using Xunit;

namespace VbaNg.Compiler.Tests;

/// <summary>
/// Parser tests for MS-VBAL sections 4 and 5. Every grammar production has a round-trip case
/// (CLAUDE.md R14): parse then print must equal the input byte for byte, with no diagnostics.
/// </summary>
public sealed class ParserTests
{
    private const string FilePath = @"C:\test\Module1.bas";

    private static SyntaxTree Parse(string text) => SyntaxTree.Parse(text, FilePath);

    private static ProcedureDeclarationSyntax FirstProcedure(SyntaxTree tree) => tree.Root.Procedures.First();

    private static T Statement<T>(string body)
        where T : StatementSyntax
    {
        var tree = Parse("Sub M()\r\n" + body + "\r\nEnd Sub\r\n");
        Assert.Empty(tree.Diagnostics);
        return Assert.IsType<T>(FirstProcedure(tree).Body[0]);
    }

    private static ExpressionSyntax Expression(string text)
    {
        var statement = Statement<AssignmentStatementSyntax>("x = " + text);
        return statement.Value;
    }

    // Round trips, grouped by MS-VBAL section.

    [Theory]
    // 4.2 module header and attributes
    [InlineData("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Foo\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\n")]
    [InlineData("Attribute VB_Name = \"Module1\"\r\nAttribute VB_Description = \"A module\"\r\nAttribute VB_Ext_KEY = \"Key\", \"Value\"\r\n")]
    [InlineData("Sub Main()\r\nAttribute Main.VB_Description = \"Runs\"\r\nAttribute Main.VB_ProcData.VB_Invoke_Func = \"a\\n14\"\r\nEnd Sub\r\n")]
    [InlineData("Public Counter As Long\r\nAttribute Counter.VB_VarUserMemId = 0\r\nAttribute Counter.VB_VarDescription = \"n\"\r\n")]
    // 5.2.1 options, 5.2.2 DefType
    [InlineData("Option Explicit\r\nOption Base 1\r\nOption Compare Text\r\nOption Compare Binary\r\nOption Compare Database\r\nOption Private Module\r\n")]
    [InlineData("DefInt A-Z\r\nDefStr S, T-V\r\nDefLng L\r\nDefBool B: DefVar V\r\n")]
    // 5.2.3 variables and constants
    [InlineData("Dim a\r\nDim b As Long, c As String * 20, d() As Variant, e(10), f(1 To 5, 0 To 2) As Byte\r\nPrivate g As New Collection\r\nPublic h As Excel.Range\r\nGlobal i As Object\r\nDim j%, k&, l$, m!, n#, o@, p^\r\nPrivate WithEvents app As Excel.Application\r\nStatic q As Boolean\r\nDim [My Var] As Long\r\n")]
    [InlineData("Const A = 1\r\nPrivate Const B As String = \"x\", C = -2.5\r\nPublic Const D As Long = &HFF&\r\nGlobal Const E = A + 1\r\n")]
    [InlineData("Private Type Point\r\n    X As Double\r\n    Y As Double\r\n    Name As String * 10\r\n    Items(1 To 3) As Long\r\n    Inner As Other\r\nEnd Type\r\nPublic Type Empty1\r\nEnd Type\r\n")]
    [InlineData("Private Type PICTDESC\r\n    Size As Long\r\n    Type As Long\r\n    hPic As LongPtr\r\n    Next As Long\r\n    Global As Boolean\r\n    Date As Date\r\n    String As String\r\nEnd Type\r\n")]
    [InlineData("Sub M()\r\n    Err.Raise &H80040201, \"src\", \"failed\"\r\n    Foo &HFF&, &O17, &7\r\n    x = &H1 And &HFF Or &O7\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    ReDim Preserve List(pt.hash).Elements(UBound(List(pt.hash).Elements) + GRANULARITY)\r\n    ReDim This.Items(0 To n) As Long, arr(1)(2)(3 To 4)\r\n    If x Then ReDim a(1) Else ReDim b(2)\r\nEnd Sub\r\n")]
    [InlineData("Public Enum Color\r\n    Red\r\n    Green = 2\r\n    Blue = Green + 1\r\n    [_Hidden] = -1\r\n    Mask = &H1000\r\nEnd Enum\r\nPrivate Enum E2: A: End Enum\r\n")]
    [InlineData("#If VBA7 Then\r\n    Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\r\n#Else\r\n    Private Declare Function GetTickCount Lib \"kernel32\" () As Long\r\n#End If\r\nDeclare Sub Sleep Lib \"kernel32\" Alias \"Sleep\" (ByVal ms As Long)\r\nPublic Declare PtrSafe Function F CDecl Lib \"x.dll\" (ByVal a As LongPtr, ByRef b As Any) As LongPtr\r\n")]
    [InlineData("Public Event Changed(ByVal Index As Long, Cancel As Boolean)\r\nEvent Done()\r\nEvent Bare\r\nImplements IComparable\r\nImplements Lib.IFoo\r\n")]
    // 5.3 procedures
    [InlineData("Sub A()\r\nEnd Sub\r\nPublic Sub B\r\nEnd Sub\r\nPrivate Static Function C(x) As Long\r\n    C = x\r\nEnd Function\r\nFriend Function D() As String()\r\nEnd Function\r\nPublic Property Get Count() As Long\r\n    Count = 1\r\nEnd Property\r\nProperty Let Count(ByVal v As Long)\r\nEnd Property\r\nProperty Set Item(ByVal i As Long, ByRef v As Object)\r\nEnd Property\r\nFunction Left$(s As String) As String\r\nEnd Function\r\n")]
    [InlineData("Sub P(a, ByVal b As Long, ByRef c As String, Optional d As Variant, Optional ByVal e As Long = 5, Optional f = \"x\", ParamArray g() As Variant)\r\nEnd Sub\r\nSub Q(arr() As Long, Optional o As Object = Nothing)\r\nEnd Sub\r\n")]
    // 5.4.1 labels, 5.4.2 control flow
    [InlineData("Sub M()\r\nStart:\r\n10 x = 1\r\n20: y = 2\r\nRetry: On Error GoTo Handler\r\n    GoTo Start\r\n    GoSub 20\r\n    Return\r\nHandler:\r\n    Resume Next\r\n    Resume\r\n    Resume Retry\r\n    Resume 10\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    If a Then\r\n        x = 1\r\n    ElseIf b Then\r\n        x = 2\r\n    ElseIf c Then\r\n    Else\r\n        x = 3\r\n    End If\r\n    If a Then x = 1\r\n    If a Then x = 1 Else x = 2\r\n    If a Then x = 1: y = 2 Else z = 3: w = 4\r\n    If a Then Else x = 5\r\n    If a Then If b Then x = 1 Else x = 2\r\n    If a Then 10\r\n    If a Then: x = 1\r\n    If a Then ' comment\r\n        x = 1\r\n    End If\r\n    If a Then _\r\n        x = 1\r\n10  If a Then Exit Sub\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    Select Case x\r\n        Case 1\r\n            y = 1\r\n        Case 2, 3, 4 To 6, Is > 10, Is <= -1, \"a\" To \"z\"\r\n            y = 2\r\n        Case Is = 7: y = 7\r\n        Case Else\r\n            y = 0\r\n    End Select\r\n    Select Case True\r\n    End Select\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    For i = 1 To 10\r\n        For j = i To 1 Step -1\r\n            x = x + 1\r\n        Next j\r\n    Next i\r\n    For i = 1 To 3: Next\r\n    For i = 1 To 3\r\n        For j = 1 To 3\r\n            For k = 1 To 3\r\n    Next k, j, i\r\n    For Each c In coll\r\n        Debug.Print c\r\n    Next c\r\n    For Each Me.Item(1) In arr(): Next\r\n    For x! = 0.5 To 2.5 Step 0.5: Next x!\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    Do While x < 10\r\n        x = x + 1\r\n    Loop\r\n    Do Until done\r\n    Loop\r\n    Do\r\n        x = x + 1\r\n    Loop While x < 10\r\n    Do\r\n    Loop Until True\r\n    Do: Loop\r\n    While x < 5\r\n        x = x + 1\r\n    Wend\r\n    Do\r\n        Exit Do\r\n    Loop\r\n    For i = 1 To 2\r\n        Exit For\r\n    Next\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    With Sheet1.Range(\"A1\")\r\n        .Value = 1\r\n        .Font.Bold = True\r\n        x = .Offset(1, 0).Value + !Field\r\n        With .Interior\r\n            .Color = vbRed\r\n        End With\r\n        .Select\r\n        .Cells(1).Clear\r\n    End With\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    On Error GoTo 0\r\n    On Error GoTo -1\r\n    On Error GoTo Handler\r\n    On Error Resume Next\r\n    On x GoTo L1, L2, 30\r\n    On y GoSub L1\r\n    Error 5\r\n    Error (Err.Number + 1)\r\n    Err.Raise 5, \"src\"\r\n    Stop\r\n    End\r\n    Exit Sub\r\nL1:\r\nL2:\r\n30\r\nEnd Sub\r\nFunction F()\r\n    Exit Function\r\nEnd Function\r\nProperty Get P()\r\n    Exit Property\r\nEnd Property\r\n")]
    [InlineData("Sub M()\r\n    RaiseEvent Changed(1, False)\r\n    RaiseEvent Done\r\n    RaiseEvent Bare()\r\nEnd Sub\r\n")]
    // 5.4.2.1 call statements, 5.4.3 assignments
    [InlineData("Sub M()\r\n    Foo\r\n    Foo 1, 2\r\n    Foo (1), 2\r\n    Foo (1) And (2), 3\r\n    Check (a = 0 Or b) _\r\n        And (c), m\r\n    Foo(1)\r\n    Foo (1)\r\n    Foo 1, , 3\r\n    Foo a:=1, b:=\"x\"\r\n    Foo ByVal x, AddressOf Handler\r\n    Call Foo\r\n    Call Foo(1, 2)\r\n    Call obj.Method(x)(2)\r\n    obj.Method 1, 2\r\n    obj!Item.Method\r\n    Me.Refresh\r\n    Sheet1.Range(\"A1\").Select\r\n    Foo -1\r\n    Foo 1,\r\n    [A1].Select\r\n    Beep\r\n    Debug.Assert x > 0\r\n    Sheets.Add Type:=\"Worksheet\"\r\n    CellRange.FormatConditions.Add Type:=xlExpression, Formula1:=\"=1\"\r\n    Foo Name:=1, Print:=2, Set:=3\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    x = 1\r\n    Let y = 2\r\n    Set o = New Collection\r\n    Set o = Nothing\r\n    Set Me.Parent = obj\r\n    arr(1, 2) = 3\r\n    obj.Prop(1).Sub = 4\r\n    .Value = 5\r\n    !Field = 6\r\n    s$ = \"a\"\r\n    LSet a = b\r\n    RSet a = b\r\n    Mid(s, 1, 2) = \"ab\"\r\n    Mid$(s, 3) = \"c\"\r\n    MidB(s, 1) = \"d\"\r\n    x = Mid(s, 1, 2)\r\n    ReDim a(10)\r\n    ReDim Preserve a(1 To n) As Long, b(2, 3)\r\n    ReDim obj.Items(5)\r\n    ReDim Me!Cache(0 To 9)\r\n    Erase a, b\r\nEnd Sub\r\n")]
    // 5.4.5 file statements
    [InlineData("Sub M()\r\n    Open \"a.txt\" For Input As #1\r\n    Open path For Output Access Write Lock Read Write As #2 Len = 128\r\n    Open path For Binary Access Read Shared As fn\r\n    Open path For Random As #3 Len = reclen\r\n    Open path For Append As 4\r\n    Open path As #5\r\n    Close #1, #2\r\n    Close 3\r\n    Close\r\n    Reset\r\n    Seek #1, 10\r\n    Seek fn, pos\r\n    Lock #1\r\n    Lock #1, 5\r\n    Lock #1, 1 To 10\r\n    Unlock #1, To 10\r\n    Line Input #1, s\r\n    Width #1, 80\r\n    Print #1, a; b, Spc(3); Tab(10); c; Tab; d,\r\n    Print #1,\r\n    Print #1, \"x\";\r\n    Write #1, a, b, \"c\"\r\n    Write #2,\r\n    Input #1, a, b(1), c.d\r\n    Put #1, 5, rec\r\n    Put #1, , rec\r\n    Get #1, , buffer\r\n    Get fn, recno, buffer\r\n    x = Input(5, #1)\r\n    y = Seek(1)\r\n    z = EOF(1) Or LOF(1) > Loc(1)\r\n    Name \"old.txt\" As \"new.txt\"\r\n    Name a & \".bak\" As b\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    Debug.Print\r\n    Debug.Print \"a\"\r\n    Debug.Print \"a\"; \"b\", \"c\";\r\n    Debug.Print Spc(2); x, Tab(5); y; Tab\r\n    Debug.Print ; \"leading\"\r\n    Debug.Print a & b, -c\r\n    If x Then Debug.Print \"y\" Else Debug.Print \"n\"\r\nEnd Sub\r\n")]
    // 5.6 expressions
    [InlineData("Sub M()\r\n    x = 1 + 2 * 3 - 4 / 5 \\ 6 Mod 7 ^ 8 & \"s\"\r\n    x = -2 ^ 2 + -a * -(b) - +c\r\n    x = Not a = b And c <> d Or e < f Xor g > h Eqv i <= j Imp k >= l\r\n    x = a Like \"a*\" And o Is Nothing And TypeOf o Is Excel.Range And TypeOf o Is Object\r\n    x = (1 + 2) * ((3))\r\n    x = a =< b Or c => d Or e >< f\r\n    x = 1.5E+10# + &HFF + &O17 + 1@ + 1! + .5 + 1D3 + 100& + 32767%\r\n    x = #1/1/2000# + #12:30:45 PM# + #Jan 1, 2000 3:00 PM#\r\n    x = \"quoted \"\"string\"\"\" & vbCrLf\r\n    x = True Or False Or Null Or Empty Or Nothing Is Nothing\r\n    x = obj.Prop.Method(1, , 3).Item(\"k\")!Field(2).Value\r\n    x = arr(1)(2)\r\n    x = [Sheet1!A1] + [Foo Bar].Value\r\n    x = Me.Name & Me!Field & .Item & !Key\r\n    x = f(a:=1, b:=2) + g(, 2) + h()\r\n    x = New Collection Is Nothing\r\n    x = Len(s) + Abs(-1) + CLng(\"1\") + Int(2.5) + Fix(-2.5) + Sgn(x) + String(3, \"a\") + Date + Array(1, 2)(0) + LBound(a) + UBound(a, 2) + Seek(1)\r\n    x = a.Print + a.Sub + a.End + a.[Foo] + a.Next\r\n    x = 2 ^ 3 ^ 2\r\n    x = a& + b& + a & b\r\n    x = a!b!c\r\n    x = -x ^ 2\r\n    x = a AND b\r\nEnd Sub\r\n")]
    [InlineData("Sub M()\r\n    x = 1 + _\r\n        2 + _\r\n        3\r\n    Foo a, _\r\n        b\r\n    If x _\r\n        Then y = 1\r\n    Dim s As String, _\r\n        t As Long\r\n    x = \"a\" _\r\n      & \"b\"\r\nEnd Sub\r\n")]
    // 3.4 conditional compilation
    [InlineData("#Const DEBUGGING = 1\r\n#Const NAME = \"x\" & \"y\"\r\n#If DEBUGGING = 1 And Win64 Then\r\nPublic Const Mode = 1\r\n#ElseIf Mac Then\r\nthis is not valid vba ( at all\r\n#Else\r\nPublic Const Mode = 3\r\n#End If\r\n#If Not VBA7 Then\r\n  garbage \"\r\n  #If Nested Then\r\n  more\r\n  #End If\r\n#End If\r\n#If 0 Then\r\n#ElseIf Len(NAME) = 2 Then\r\nPublic Const Mode2 = 2\r\n#End If\r\n")]
    // comments and separators everywhere
    [InlineData("' leading comment\r\n\r\nOption Explicit ' trailing\r\n\r\n    ' indented comment\r\nRem old style\r\nSub M() ' after header\r\n    x = 1: y = 2 ' two statements\r\n    x = 1:: y = 2\r\n    : z = 3\r\n    Rem inside\r\n    If a Then ' comment after Then\r\n        b = 1 ' comment\r\n    Else ' comment after Else\r\n        b = 2\r\n    End If ' comment after End If\r\nEnd Sub ' comment after End Sub\r\n' trailing comment without newline")]
    [InlineData("")]
    [InlineData("\r\n\r\n")]
    [InlineData("Sub M()\r\nEnd Sub")]
    [InlineData("x = 1\ny = 2\rz = 3\r\n")]
    [InlineData("Sub M(): x = 1: End Sub\r\nSub N(): End Sub\r\nFunction F(): F = 1: End Function\r\n")]
    [InlineData("Sub M()\r\n    Select Case x: Case 1: y = 1: Case Else: y = 2: End Select\r\n    If a Then Do: x = x + 1: Loop While x < 3\r\n    For i = 1 To 3: If i = 2 Then Exit For: Next\r\n    Do: If x Then Exit Do: Loop\r\nEnd Sub\r\n")]
    // VBA checks neither which End keyword closes a procedure nor which Exit leaves one, and the
    // file is printed back with the keyword it was written with (docs/vba-quirks.md).
    [InlineData("Property Get P() As Long\r\n    P = 1\r\nEnd Function\r\n")]
    [InlineData("Sub S()\r\n    Exit Property\r\nEnd Function\r\n")]
    public void Parse_RoundTripsWithoutDiagnostics(string text)
    {
        var tree = Parse(text);

        Assert.Equal(text, tree.Root.ToFullString());
        Assert.Empty(tree.Diagnostics);
        Assert.False(tree.Root.ContainsMissingTokens);
    }

    [Theory]
    [InlineData("Sub M()\r\n    If x\r\n    End If\r\nEnd Sub\r\n", 2, 9, "Expected: Then")]
    [InlineData("Sub M()\r\n    x = 1 2\r\nEnd Sub\r\n", 2, 11, "Expected: end of statement")]
    [InlineData("Sub M()\r\n    x = \r\nEnd Sub\r\n", 2, 8, "Expected: expression")]
    [InlineData("Sub M()\r\n    x = (1\r\nEnd Sub\r\n", 2, 11, "Expected: )")]
    [InlineData("Sub M()\r\n    Next\r\nEnd Sub\r\n", 2, 5, "Next without For")]
    [InlineData("Sub M()\r\n    End If\r\nEnd Sub\r\n", 2, 5, "End If without block If")]
    [InlineData("Sub M()\r\n    For i = 1 To 3\r\nEnd Sub\r\n", 3, 1, "Expected: Next")]
    [InlineData("Sub M()\r\n    Dim Local As Long\r\nEnd Sub\r\n", 2, 9, "Expected: identifier")]
    [InlineData("Sub M()\r\n    x = 1\r\n", 2, 10, "Expected: End Sub")]
    [InlineData("Sub M()\r\nSub N()\r\nEnd Sub\r\n", 2, 1, "Expected: End Sub")]
    [InlineData("Sub M()\r\n    ? = 1\r\nEnd Sub\r\n", 2, 5, "Unexpected character '?'.")]
    [InlineData("Sub M()\r\nEnd Sub\r\nDim x As Long\r\n", 3, 1, "Only comments may appear after End Sub, End Function, or End Property.")]
    [InlineData("Option Fancy\r\n", 1, 8, "Expected: Explicit, Base, Compare, or Private")]
    [InlineData("Sub M()\r\n    Select Case x\r\n        y = 1\r\n        Case 1\r\n    End Select\r\nEnd Sub\r\n", 3, 9, "Expected: Case")]
    [InlineData("Sub M()\r\n    Exit Loop\r\nEnd Sub\r\n", 2, 10, "Expected: Sub, Function, Property, Do, or For")]
    [InlineData("Sub M()\r\n    Do\r\n    Wend\r\nEnd Sub\r\n", 3, 5, "Wend without While")]
    [InlineData("Sub M()\r\n    Do\r\nEnd Sub\r\n", 3, 1, "Expected: Loop")]
    [InlineData("Sub M()\r\n    While x\r\nEnd Sub\r\n", 3, 1, "Expected: Wend")]
    [InlineData("#If Foo( Then\r\n#End If\r\n", 1, 9, "Expected: )")]
    [InlineData("#If Len(Foo()) Then\r\n#End If\r\n", 1, 5, "'Foo' is not allowed in a conditional compilation expression.")]
    public void Parse_ReportsTheFirstErrorAndStillRoundTrips(string text, int line, int column, string message)
    {
        var tree = Parse(text);

        Assert.Equal(text, tree.Root.ToFullString());
        Assert.NotEmpty(tree.Diagnostics);
        var first = tree.Diagnostics[0];
        Assert.Equal((line, column, message), (first.Line, first.Column, first.Message));
        Assert.StartsWith("VBA000", first.Id, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Sub M(\r\n")]
    [InlineData("If If If\r\nElse Else\r\nEnd End End If\r\n")]
    [InlineData("For For For To To Next Next\r\n")]
    [InlineData(")))(((\r\n")]
    [InlineData("Sub M()\r\n    With\r\n    .\r\n    !\r\n    x = a.\r\n    y = b!\r\n    z = c(\r\nEnd Sub\r\n")]
    [InlineData("Select Case\r\nCase\r\nCase Is\r\nCase 1 To\r\nEnd Select\r\n")]
    [InlineData("Dim As\r\nDim x As\r\nDim x As String *\r\nConst\r\nConst x\r\nType\r\nEnum\r\nDeclare\r\nEvent\r\nImplements\r\n")]
    [InlineData("Open\r\nOpen x For\r\nPrint\r\nGet #\r\nLine Input\r\nName x\r\n")]
    [InlineData("#If\r\n#ElseIf\r\n#Else\r\n#End If\r\n#End If\r\n#Const\r\n#Const x\r\n")]
    [InlineData("Attribute\r\nAttribute VB_Name\r\nOption\r\nDefInt\r\nDefInt A-\r\n")]
    [InlineData("Sub M()\r\n    Property Get X()\r\n    End Property\r\nEnd Sub\r\n")]
    [InlineData("\0\u0001\uFFFF\r\n")]
    public void Parse_NeverThrowsAndAlwaysRoundTrips(string text)
    {
        var tree = Parse(text);

        Assert.Equal(text, tree.Root.ToFullString());
        Assert.NotEmpty(tree.Diagnostics);
    }

    // Tree shapes.

    [Fact]
    public void Module_ExposesHeaderNameAndProcedures()
    {
        var tree = Parse("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Widget\"\r\nOption Explicit\r\nPublic Sub A()\r\nEnd Sub\r\nPrivate Function B() As Long\r\nEnd Function\r\n");

        Assert.NotNull(tree.Root.Header);
        Assert.Equal(4, tree.Root.Header!.Lines.Count);
        Assert.Equal("Widget", tree.Root.Name);
        Assert.Equal(["A", "B"], tree.Root.Procedures.Select(p => p.Name.Text));
        Assert.True(tree.Root.Procedures.First().IsSub);
        Assert.True(tree.Root.Procedures.Last().IsFunction);
    }

    [Fact]
    public void Precedence_FollowsTheTableInSection569()
    {
        // -2 ^ 2 is -(2 ^ 2); Not a = b is Not (a = b); a + b * c is a + (b * c); 2 ^ 3 ^ 2 is (2 ^ 3) ^ 2.
        var negate = Assert.IsType<UnaryExpressionSyntax>(Expression("-2 ^ 2"));
        Assert.IsType<BinaryExpressionSyntax>(negate.Operand);

        var not = Assert.IsType<UnaryExpressionSyntax>(Expression("Not a = b"));
        Assert.Equal(SyntaxKind.EqualsToken, Assert.IsType<BinaryExpressionSyntax>(not.Operand).OperatorToken.Kind);

        var sum = Assert.IsType<BinaryExpressionSyntax>(Expression("a + b * c"));
        Assert.Equal(SyntaxKind.PlusToken, sum.OperatorToken.Kind);
        Assert.Equal(SyntaxKind.AsteriskToken, Assert.IsType<BinaryExpressionSyntax>(sum.Right).OperatorToken.Kind);

        var power = Assert.IsType<BinaryExpressionSyntax>(Expression("2 ^ 3 ^ 2"));
        Assert.IsType<BinaryExpressionSyntax>(power.Left);
        Assert.IsType<LiteralExpressionSyntax>(power.Right);

        var and = Assert.IsType<BinaryExpressionSyntax>(Expression("a And Not b Or c"));
        Assert.Equal(SyntaxKind.OrKeyword, and.OperatorToken.Kind);
        Assert.Equal(SyntaxKind.AndKeyword, Assert.IsType<BinaryExpressionSyntax>(and.Left).OperatorToken.Kind);

        var concat = Assert.IsType<BinaryExpressionSyntax>(Expression("a & b + c = d"));
        Assert.Equal(SyntaxKind.EqualsToken, concat.OperatorToken.Kind);
        Assert.Equal(SyntaxKind.AmpersandToken, Assert.IsType<BinaryExpressionSyntax>(concat.Left).OperatorToken.Kind);
    }

    [Fact]
    public void CallStatement_ParenthesizedFirstArgumentFollowedByComma()
    {
        var call = Statement<CallStatementSyntax>("Foo (1), 2");

        Assert.Equal("Foo", Assert.IsType<IdentifierNameSyntax>(call.Expression).Name);
        Assert.NotNull(call.Arguments);
        Assert.Equal(2, call.Arguments!.Arguments.Count);
        Assert.IsType<ParenthesizedExpressionSyntax>(call.Arguments.Arguments[0].Expression);
        Assert.Null(call.Arguments.OpenParenToken);
    }

    [Fact]
    public void CallStatement_ParenthesizedFirstArgumentContinuedByAnOperator()
    {
        // VBA-TDD's TestCase.cls: Check (Number = 0 Or Err.Number = Number) And (Source = ""), Message.
        var call = Statement<CallStatementSyntax>("Check (a = 0 Or b) And (c), m, 2");

        Assert.Equal("Check", Assert.IsType<IdentifierNameSyntax>(call.Expression).Name);
        Assert.NotNull(call.Arguments);
        Assert.Equal(3, call.Arguments!.Arguments.Count);
        var first = Assert.IsType<BinaryExpressionSyntax>(call.Arguments.Arguments[0].Expression);
        Assert.Equal(SyntaxKind.AndKeyword, first.OperatorToken.Kind);
        Assert.IsType<ParenthesizedExpressionSyntax>(first.Left);
        Assert.IsType<ParenthesizedExpressionSyntax>(first.Right);
        Assert.Null(call.Arguments.OpenParenToken);

        var lone = Statement<CallStatementSyntax>("Foo (x) + 1");
        Assert.Equal(SyntaxKind.PlusToken, Assert.IsType<BinaryExpressionSyntax>(Assert.Single(lone.Arguments!.Arguments).Expression).OperatorToken.Kind);
    }

    [Fact]
    public void CallStatement_WithParenthesizedArgumentsIsAnIndexExpression()
    {
        var call = Statement<CallStatementSyntax>("Foo(1, 2)");

        Assert.IsType<IndexExpressionSyntax>(call.Expression);
        Assert.Null(call.Arguments);
    }

    [Fact]
    public void CallStatement_NamedOmittedAndByValArguments()
    {
        var call = Statement<CallStatementSyntax>("Foo a:=1, , ByVal x");

        var arguments = call.Arguments!.Arguments.Nodes.ToList();
        Assert.True(arguments[0].IsNamed);
        Assert.Equal("a", arguments[0].Name!.Text);
        Assert.True(arguments[1].IsOmitted);
        Assert.NotNull(arguments[2].ByValKeyword);
    }

    [Fact]
    public void Assignment_LetAndSetAreDistinguished()
    {
        Assert.False(Statement<AssignmentStatementSyntax>("x = 1").IsSet);
        Assert.NotNull(Statement<AssignmentStatementSyntax>("Let x = 1").Keyword);
        Assert.True(Statement<AssignmentStatementSyntax>("Set x = New Collection").IsSet);
    }

    [Fact]
    public void SingleLineIf_HoldsColonSeparatedStatementsAndElse()
    {
        var statement = Statement<SingleLineIfStatementSyntax>("If a Then x = 1: y = 2 Else z = 3: w = 4");

        Assert.Equal(2, statement.Statements.Count);
        Assert.NotNull(statement.ElseClause);
        Assert.Equal(2, statement.ElseClause!.Statements.Count);
    }

    [Fact]
    public void SingleLineIf_WithLineNumberIsAnImplicitGoTo()
    {
        var statement = Statement<SingleLineIfStatementSyntax>("If a Then 100");

        var jump = Assert.IsType<GoToStatementSyntax>(Assert.Single(statement.Statements));
        Assert.Equal("100", jump.Label.Text);
        Assert.Equal(string.Empty, jump.Keyword.Text);
    }

    [Fact]
    public void BlockIf_CollectsElseIfAndElse()
    {
        var statement = Statement<IfBlockSyntax>("If a Then\r\n    x = 1\r\nElseIf b Then\r\n    x = 2\r\nElse\r\n    x = 3\r\nEnd If");

        Assert.Single(statement.Statements);
        Assert.Single(statement.ElseIfBlocks.AsEnumerable());
        Assert.NotNull(statement.ElseBlock);
        Assert.Equal(SyntaxKind.IfKeyword, statement.EndIf.BlockKeyword!.Kind);
    }

    [Fact]
    public void Next_WithSeveralVariablesClosesSeveralLoops()
    {
        var outer = Statement<ForBlockSyntax>("For i = 1 To 2\r\n    For j = 1 To 2\r\n        x = 1\r\n    Next j, i");

        var inner = Assert.IsType<ForBlockSyntax>(Assert.Single(outer.Statements));
        Assert.Null(inner.Next);
        Assert.NotNull(outer.Next);
        Assert.Equal(2, outer.Next!.Variables.Count);
    }

    [Fact]
    public void SelectCase_ClauseForms()
    {
        var select = Statement<SelectCaseBlockSyntax>("Select Case x\r\n    Case 1, 2 To 3, Is > 4\r\n    Case Else\r\nEnd Select");

        var clauses = select.CaseBlocks[0].Clauses.Nodes.ToList();
        Assert.IsType<ValueCaseClauseSyntax>(clauses[0]);
        Assert.IsType<RangeCaseClauseSyntax>(clauses[1]);
        Assert.Equal(SyntaxKind.GreaterThanToken, Assert.IsType<IsCaseClauseSyntax>(clauses[2]).OperatorToken.Kind);
        Assert.True(select.CaseBlocks[1].IsCaseElse);
    }

    [Fact]
    public void Labels_AndLineNumbers()
    {
        var tree = Parse("Sub M()\r\nRetry:\r\n10 x = 1\r\n20: y = 2\r\nEnd Sub\r\n");
        Assert.Empty(tree.Diagnostics);

        var body = FirstProcedure(tree).Body;
        Assert.Equal(
            [SyntaxKind.LabelStatement, SyntaxKind.LabelStatement, SyntaxKind.LetAssignmentStatement, SyntaxKind.LabelStatement, SyntaxKind.LetAssignmentStatement],
            body.Select(s => s.Kind));
        Assert.True(((LabelStatementSyntax)body[1]).IsLineNumber);
        Assert.Null(((LabelStatementSyntax)body[1]).ColonToken);
        Assert.NotNull(((LabelStatementSyntax)body[3]).ColonToken);
    }

    [Fact]
    public void ContextualStatements_NameErrorMidLineInputWidthReset()
    {
        var tree = Parse("Sub M()\r\n    Name a As b\r\n    Error 5\r\n    Mid(s, 1) = \"x\"\r\n    Line Input #1, s\r\n    Width #1, 80\r\n    Reset\r\n    Name = 1\r\n    Error.Foo\r\n    x = Mid(s, 1)\r\n    Mid$(s, 2, 1) = \"y\"\r\n    MidB$(s, 1) = \"z\"\r\n    Sheets.Add Type:=\"Worksheet\"\r\nEnd Sub\r\n");
        Assert.Empty(tree.Diagnostics);

        Assert.Equal(
            [
                SyntaxKind.NameStatement, SyntaxKind.ErrorStatement, SyntaxKind.MidStatement, SyntaxKind.LineInputStatement,
                SyntaxKind.WidthStatement, SyntaxKind.ResetStatement, SyntaxKind.LetAssignmentStatement, SyntaxKind.CallStatement,
                SyntaxKind.LetAssignmentStatement, SyntaxKind.MidStatement, SyntaxKind.MidStatement, SyntaxKind.CallStatement,
            ],
            FirstProcedure(tree).Body.Select(s => s.Kind));
    }

    [Fact]
    public void Declarations_CarryModifiersAndTypes()
    {
        var tree = Parse("Private WithEvents app As Excel.Application\r\nDim s As String * 10, a(1 To 5) As Long\r\nPublic Const C As Long = 1\r\n");
        Assert.Empty(tree.Diagnostics);

        var first = Assert.IsType<VariableDeclarationSyntax>(tree.Root.Members[0]);
        Assert.True(first.Modifiers.Any(SyntaxKind.PrivateKeyword));
        var declarator = Assert.Single(first.Declarators);
        Assert.NotNull(declarator.WithEventsKeyword);
        Assert.IsType<NamedTypeSyntax>(declarator.AsClause!.Type);

        var second = Assert.IsType<VariableDeclarationSyntax>(tree.Root.Members[1]);
        var fixedString = Assert.IsType<BuiltinTypeSyntax>(second.Declarators[0].AsClause!.Type);
        Assert.NotNull(fixedString.Length);
        Assert.Single(second.Declarators[1].Bounds!.Bounds);

        var constant = Assert.IsType<ConstDeclarationSyntax>(tree.Root.Members[2]);
        Assert.Equal("C", constant.Declarators[0].Name.Text);
    }

    [Fact]
    public void Procedure_ParametersAndReturnType()
    {
        var tree = Parse("Public Static Function F(ByVal a As Long, Optional b = 2, ParamArray c()) As String()\r\nEnd Function\r\n");
        Assert.Empty(tree.Diagnostics);

        var function = FirstProcedure(tree);
        Assert.Equal(2, function.Modifiers.Count);
        var parameters = function.Parameters!.Parameters.Nodes.ToList();
        Assert.True(parameters[0].Modifiers.Any(SyntaxKind.ByValKeyword));
        Assert.NotNull(parameters[1].DefaultValue);
        Assert.NotNull(parameters[2].Bounds);
        Assert.NotNull(function.AsClause!.ArrayDesignator);
    }

    [Fact]
    public void Property_AccessorsAreRecognized()
    {
        var tree = Parse("Property Get X() As Long\r\nEnd Property\r\nProperty Let X(v As Long)\r\nEnd Property\r\nProperty Set X(v As Object)\r\nEnd Property\r\n");
        Assert.Empty(tree.Diagnostics);

        Assert.Equal(
            [SyntaxKind.GetKeyword, SyntaxKind.LetKeyword, SyntaxKind.SetKeyword],
            tree.Root.Procedures.Select(p => p.AccessorKeyword!.Kind));
        Assert.All(tree.Root.Procedures, p => Assert.True(p.IsProperty));
    }

    // Conditional compilation (MS-VBAL 3.4).

    [Fact]
    public void ConditionalCompilation_KeepsTheTakenBranchOnly()
    {
        var tree = Parse("#If Win64 Then\r\nConst A = 1\r\n#Else\r\nConst A = 2\r\n#End If\r\n");
        Assert.Empty(tree.Diagnostics);

        var constant = Assert.IsType<ConstDeclarationSyntax>(Assert.Single(tree.Root.Members));
        Assert.Equal("1", Assert.IsType<LiteralExpressionSyntax>(constant.Declarators[0].Value).Token.Text);
        var trivia = tree.Root.EndOfFileToken.LeadingTrivia;
        Assert.Contains(trivia, t => t.Kind == SyntaxKind.ConditionalDirectiveTrivia);
        Assert.Contains(trivia, t => t.Kind == SyntaxKind.DisabledTextTrivia && t.Text == "Const A = 2\r\n");
    }

    [Fact]
    public void ConditionalCompilation_ElseIfAndConst()
    {
        var tree = Parse("#Const LEVEL = 2\r\n#If LEVEL = 1 Then\r\nConst A = 1\r\n#ElseIf LEVEL = 2 Then\r\nConst A = 2\r\n#ElseIf LEVEL = 2 Then\r\nConst A = 22\r\n#Else\r\nConst A = 3\r\n#End If\r\n");
        Assert.Empty(tree.Diagnostics);

        var constant = Assert.IsType<ConstDeclarationSyntax>(Assert.Single(tree.Root.Members));
        Assert.Equal("2", Assert.IsType<LiteralExpressionSyntax>(constant.Declarators[0].Value).Token.Text);
    }

    [Fact]
    public void ConditionalCompilation_UndefinedConstantIsFalseAndOptionsOverride()
    {
        var text = "#If CUSTOM Then\r\nConst A = 1\r\n#End If\r\n";

        Assert.Empty(Parse(text).Root.Members);
        var options = new ParseOptions(new Dictionary<string, object?> { ["CUSTOM"] = true });
        Assert.Single(SyntaxTree.Parse(text, FilePath, options).Root.Members);
    }

    [Fact]
    public void ConditionalCompilation_ExcludedTextNeedNotBeValid()
    {
        var tree = Parse("#If Mac Then\r\n    this is \"not valid ( vba\r\n    #If Nested Then\r\n    #End If\r\n#End If\r\nConst A = 1\r\n");

        Assert.Empty(tree.Diagnostics);
        Assert.Single(tree.Root.Members);
    }

    [Fact]
    public void Diagnostics_UseTheCanonicalFormat()
    {
        var tree = Parse("Sub M()\r\n    If x\r\nEnd Sub\r\n");

        Assert.Equal(FilePath + "(2,9): error VBA0001: Expected: Then", tree.Diagnostics[0].ToString());
    }

    [Fact]
    public void ErrorRecovery_ContinuesWithTheNextStatement()
    {
        var tree = Parse("Sub M()\r\n    x = = 1\r\n    y = 2\r\nEnd Sub\r\nSub N()\r\nEnd Sub\r\n");

        Assert.Equal(2, tree.Root.Procedures.Count());
        var body = FirstProcedure(tree).Body;
        Assert.Contains(body, s => s is AssignmentStatementSyntax a && a.Target is IdentifierNameSyntax { Name: "y" });
        Assert.Contains(body, s => s is BadStatementSyntax || s.ContainsMissingTokens);
    }
}
