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
Private bFechamentoAutorizado As Boolean

Private Sub UserForm_Initialize()
    bFechamentoAutorizado = False
    
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        Me.Caption = "Acesso ao Sistema - Identificação do Operador"
        BtEntrar.Caption = "ENTRAR NO SISTEMA"
    ElseIf bModoTroca Then
        Me.Caption = "Trocar Operador Ativo"
        BtEntrar.Caption = "TROCAR OPERADOR"
    Else
        Me.Caption = "Identificação do Operador - PDV"
        BtEntrar.Caption = "ENTRAR (PDV)"
    End If
    
    On Error Resume Next
    LblInfo.Caption = "Acesso Individual Protegido por PIN:" & vbCrLf & _
                      "• Recepção: Vitória ou Geovana" & vbCrLf & _
                      "• Administração: Administrador (PIN Restrito)"
    On Error GoTo 0
    
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

Private Sub UserForm_Activate()
    On Error Resume Next
    TxtSenha.SetFocus
    On Error GoTo 0
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
    
    Dim eraLoginInicial As Boolean
    eraLoginInicial = (Trim(ModSeguranca.vUsuarioAtivo) = "")
    
    Dim autenticado As Boolean
    autenticado = ModSeguranca.AutenticarUsuario(nomeOp, Trim(TxtSenha.Text))
    
    If autenticado Then
        LblMensagem.Caption = ""
        bFechamentoAutorizado = True
        
        On Error Resume Next
        ThisWorkbook.Sheets("INÍCIO").Range("B5").Value = "  SISTEMA ONLINE  |  OPERADOR ATIVO: " & UCase(ModSeguranca.vNomeOperador) & " (" & ModSeguranca.vPerfilAtivo & ")"
        On Error GoTo 0
        
        Unload Me
        
        If Not bModoTroca Then
            FormVendas.Show
        ElseIf eraLoginInicial Then
            MsgBox "Bem-vindo(a), " & ModSeguranca.vNomeOperador & "!" & vbCrLf & vbCrLf & _
                   "• Perfil de Acesso: " & ModSeguranca.vPerfilAtivo, vbInformation, "Acesso Autorizado"
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
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        If MsgBox("O login é obrigatório para utilizar a planilha." & vbCrLf & vbCrLf & _
                  "Deseja fechar a planilha agora?", vbQuestion + vbYesNo, "Login Obrigatório") = vbYes Then
            bFechamentoAutorizado = True
            Unload Me
            ThisWorkbook.Close SaveChanges:=False
        End If
        Exit Sub
    End If
    bFechamentoAutorizado = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And Not bFechamentoAutorizado Then
        If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
            Cancel = True
            LblMensagem.Caption = "Informe seu PIN para entrar ou clique em FECHAR!"
            On Error Resume Next
            TxtSenha.SetFocus
            On Error GoTo 0
        End If
    End If
End Sub


