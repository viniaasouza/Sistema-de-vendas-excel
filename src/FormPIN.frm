VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPIN 
   Caption         =   "UserForm1"
   ClientHeight    =   2863
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "FormPIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sPINDigitado As String
Public bConfirmado As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Autorização Administrativa"
    Me.Width = 300
    Me.Height = 205
    bConfirmado = False
    sPINDigitado = ""
    TxtPIN.PasswordChar = "*"
    TxtPIN.Text = ""
End Sub

Private Sub UserForm_Activate()
    TxtPIN.PasswordChar = "*"
    TxtPIN.Text = ""
    On Error Resume Next
    TxtPIN.SetFocus
    On Error GoTo 0
End Sub

Public Sub BtnConfirmar_Click()
    sPINDigitado = Trim(TxtPIN.Text)
    bConfirmado = True
    Me.Hide
End Sub

Public Sub BtnCancelar_Click()
    sPINDigitado = ""
    bConfirmado = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        sPINDigitado = ""
        bConfirmado = False
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub TxtPIN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call BtnConfirmar_Click
    ElseIf KeyCode = vbKeyEscape Then
        Call BtnCancelar_Click
    End If
End Sub

