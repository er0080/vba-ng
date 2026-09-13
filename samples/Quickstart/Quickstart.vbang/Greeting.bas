Attribute VB_Name = "Greeting"
Option Explicit

' The button on Sheet1 has OnAction = Greeting.SayHello, which is how Excel runs a Form control's
' macro. Debug.Print goes to the console: the terminal under `vbang run`, the Debug Console under F5.

Public Sub SayHello()
    Worksheets("Sheet1").Range("B2").Value = "Hello, world"
    Debug.Print "Hello, world"
End Sub
