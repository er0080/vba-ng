VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Disc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Implements IShape; its own Area differs from the interface's so the dispatch is visible.

Implements IShape

Public Radius As Double

Private Function IShape_Area() As Double
    IShape_Area = 3 * Radius * Radius
End Function

Private Property Get IShape_Kind() As String
    IShape_Kind = "circle"
End Property

Public Function Area() As Double
    Area = 0
End Function
