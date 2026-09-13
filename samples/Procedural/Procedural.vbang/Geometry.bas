Attribute VB_Name = "Geometry"
Option Explicit

' Points, shapes, and the procedure features M3 covers: user-defined types, enums,
' ByRef records, Optional and ParamArray parameters, and a Static local.

Public Type TPoint
    X As Double
    Y As Double
End Type

Public Enum ShapeKind
    skPoint
    skLine
    skTriangle = 3
End Enum

Public Function MakePoint(ByVal x As Double, ByVal y As Double) As TPoint
    MakePoint.X = x
    MakePoint.Y = y
End Function

Public Sub MovePoint(ByRef p As TPoint, ByVal dx As Double, ByVal dy As Double)
    p.X = p.X + dx
    p.Y = p.Y + dy
End Sub

Public Function Distance(ByRef a As TPoint, ByRef b As TPoint) As Double
    Distance = Sqr((a.X - b.X) ^ 2 + (a.Y - b.Y) ^ 2)
End Function

Public Function Perimeter(ParamArray sides() As Variant) As Double
    Dim i As Long
    For i = LBound(sides) To UBound(sides)
        Perimeter = Perimeter + sides(i)
    Next i
End Function

Public Function KindName(ByVal kind As ShapeKind) As String
    Select Case kind
        Case skPoint
            KindName = "point"
        Case skLine
            KindName = "line"
        Case skTriangle
            KindName = "triangle"
        Case Else
            KindName = "unknown"
    End Select
End Function

Public Function Describe(ByVal kind As ShapeKind, Optional ByVal label As String = "shape") As String
    Describe = label & ": " & KindName(kind)
End Function

Public Function NextId() As Long
    Static counter As Long
    counter = counter + 1
    NextId = counter
End Function
