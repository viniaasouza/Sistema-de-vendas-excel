VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormLogin 
   Caption         =   "UserForm1"
   ClientHeight    =   5705
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6818
   OleObjectBlob   =   "FormLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bModoTroca As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Identificação do Operador - PDV"
    
    CboOperador.Clear
    Dim wsU As Worksheet
    Set wsU = ModSeguranca.PlanUsuarios()
    
    If Not wsU Is Nothing Then
        Dim ultLin As Long, r As Long
        ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
        For r = 2 To ultLin
            If UCase(Trim(wsU.Cells(r, "F").Text)) = "SIM" Then
                Dim n As String, p As String
                n = Trim(wsU.Cells(r, "B").Text)
                p = Trim(wsU.Cells(r, "E").Text)
                CboOperador.AddItem n & " (" & p & ")"
            End If
        Next r
    End If
    
    If CboOperador.ListCount > 0 Then
        Dim idxSel As Long, i As Long
        idxSel = 0
        If ModSeguranca.vNomeOperador <> "" Then
            For i = 0 To CboOperador.ListCount - 1
                If InStr(1, CboOperador.List(i), ModSeguranca.vNomeOperador, vbTextCompare) > 0 Then
                    idxSel = i
                    Exit For
                End If
            Next i
        End If
        CboOperador.ListIndex = idxSel
    End If
    
    TxtSenha.Text = ""
    LblMensagem.Caption = ""
End Sub

Private Sub CboOperador_Change()
    TxtSenha.Text = ""
    LblMensagem.Caption = ""
    On Error Resume Next
    TxtSenha.SetFocus
    On Error GoTo 0
End Sub

Private Sub BtEntrar_Click()
    If CboOperador.ListIndex = -1 Or Trim(CboOperador.Text) = "" Then
        LblMensagem.Caption = "Selecione o operador!"
        CboOperador.SetFocus
        Exit Sub
    End If
    
    If Trim(TxtSenha.Text) = "" Then
        LblMensagem.Caption = "Digite o PIN / Senha!"
        TxtSenha.SetFocus
        Exit Sub
    End If
    
    Dim txtOp As String, nomeOp As String
    txtOp = CboOperador.Text
    If InStr(txtOp, "(") > 0 Then
        nomeOp = Trim(Left(txtOp, InStr(txtOp, "(") - 1))
    Else
        nomeOp = Trim(txtOp)
    End If
    
    Dim autenticado As Boolean
    autenticado = ModSeguranca.AutenticarUsuario(nomeOp, Trim(TxtSenha.Text))
    
    If autenticado Then
        LblMensagem.Caption = ""
        Unload Me
        
        If Not bModoTroca Then
            FormVendas.Show
        Else
            MsgBox "Operador alterado com sucesso!" & vbCrLf & vbCrLf & _
                   "• Operador Ativo: " & ModSeguranca.vNomeOperador & vbCrLf & _
                   "• Perfil: " & ModSeguranca.vPerfilAtivo, vbInformation, "Operador Atualizado"
        End If
    Else
        LblMensagem.Caption = "PIN / Senha incorreta para o operador!"
        TxtSenha.Text = ""
        TxtSenha.SetFocus
    End If
End Sub

Private Sub TxtSenha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call BtEntrar_Click
    End If
End Sub

Private Sub BtSair_Click()
    Unload Me
End Sub


