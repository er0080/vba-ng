Attribute VB_Name = "Main"
Option Explicit

' The entry point vbang run calls: vbang run Main.Main --project samples/Procedural/Procedural.vbang

Public Sub Main()
    Dim origin As TPoint
    Dim p As TPoint
    p = MakePoint(3, 4)
    Debug.Print "Distance from origin: " & Distance(origin, p)
    Debug.Print Describe(skTriangle, "first")
    Debug.Print "Next id: " & NextId()
End Sub
