VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Sign in"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCancelled As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Private Sub OkButton_Click()
    If Len(Trim$(UserName.Text)) = 0 Then
        MsgBox "Enter a user name.", vbExclamation
        UserName.SetFocus
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    mCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CancelButton_Click
    End If
End Sub
