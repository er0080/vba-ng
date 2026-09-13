Attribute VB_Name = "Tests"
Option Explicit

' Test procedures: Public Subs without parameters under a '@Test annotation, run by
' vbang test. Assert compares with VBA's = under this module's Option Compare.

'@Test
Public Sub MovePoint_AddsTheOffset()
    Dim p As TPoint
    p = MakePoint(1, 2)
    MovePoint p, 3, 4
    Assert.AreEqual 4, p.X
    Assert.AreEqual 6, p.Y
End Sub

'@Test
Public Sub Distance_IsEuclidean()
    Dim origin As TPoint
    Dim p As TPoint
    p = MakePoint(3, 4)
    Assert.AreEqual 5, Distance(origin, p)
End Sub

'@Test
Public Sub Perimeter_SumsEverySide()
    Assert.AreEqual 12, Perimeter(3, 4, 5)
    Assert.AreEqual 0, Perimeter()
End Sub

'@Test
Public Sub KindName_CoversTheEnum()
    Assert.AreEqual "point", KindName(skPoint)
    Assert.AreEqual "triangle", KindName(skTriangle)
    Assert.AreEqual "unknown", KindName(7)
End Sub

'@Test
Public Sub Describe_UsesTheDefaultLabel()
    Assert.AreEqual "shape: line", Describe(skLine)
    Assert.AreEqual "edge: line", Describe(skLine, "edge")
End Sub

'@Test
Public Sub NextId_CountsAcrossCalls()
    Dim first As Long
    first = NextId()
    Assert.AreEqual first + 1, NextId()
End Sub

'@Test
Public Sub ParseNumbers_ReadsEveryItem()
    Dim values() As Long
    values = ParseNumbers("1, 2, 3")
    Assert.AreEqual 0, LBound(values)
    Assert.AreEqual 2, UBound(values)
    Assert.AreEqual 6, values(0) + values(1) + values(2)
End Sub

'@Test
Public Sub SafeDivide_ReportsDivisionByZero()
    Dim failed As Boolean
    Assert.AreEqual 2, SafeDivide(6, 3, failed)
    Assert.IsFalse failed
    Assert.AreEqual 0, SafeDivide(1, 0, failed)
    Assert.IsTrue failed, "division by zero must set failed"
End Sub

'@Test
Public Sub ErrorNumberOf_SeesTheSkippedError()
    Assert.AreEqual 13, ErrorNumberOf("abc")
    Assert.AreEqual 0, ErrorNumberOf("42")
End Sub

'@Test
Public Sub Classify_UsesGoSub()
    Assert.AreEqual "negative", Classify(-1)
    Assert.AreEqual "zero", Classify(0)
    Assert.AreEqual "positive", Classify(5)
End Sub

'@Test
Public Sub RetryCount_ResumesUntilItSucceeds()
    Assert.AreEqual 3, RetryCount()
End Sub

'@Test
Public Sub Collection_KeepsInsertionOrder()
    Dim items As New Collection
    items.Add "b", "second"
    items.Add "a", "first", 1
    Assert.AreEqual "a", items(1)
    Assert.AreEqual 2, items.Count
    Assert.IsNotNothing items
End Sub
