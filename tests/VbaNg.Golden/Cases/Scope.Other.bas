Attribute VB_Name = "Other"
Option Explicit

' A second standard module: a function and a constant with the same names as Helpers has, so
' only the qualified forms reach them, and a module name that Helpers also uses for a function.

Public Const Tag As String = "Other"

Public Function Both() As String
    Both = "Other"
End Function
