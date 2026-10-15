Attribute VB_Name = "ModSeguranca"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE SEGURANÇA, CONTROLE DE ACESSO E OPERADORES (FASE 5 & 6)
' =============================================================
Public vUsuarioAtivo As String
Public vNomeOperador As String
Public vPerfilAtivo As String
Public vIDUsuarioAtivo As Long

Public Function PlanUsuarios() As Worksheet
    On Error Resume Next
    Set PlanUsuarios = ThisWorkbook.Sheets("USUARIOS")
    On Error GoTo 0
End Function

Public Sub ProtegerAbaUsuarios()
    ' Oculta a aba USUARIOS de forma blindada (xlSheetVeryHidden = 2)
    ' Impede visualização ou desocultação pelo menu normal do Excel
    On Error Resume Next
    Dim wsU As Worksheet
    Set wsU = PlanUsuarios()
    If Not wsU Is Nothing Then
        wsU.Visible = 2 ' 2 = xlSheetVeryHidden
    End If
    On Error GoTo 0
End Sub

Public Sub InicializarSessaoPadrao()
    Call ProtegerAbaUsuarios
    If Trim(vUsuarioAtivo) = "" Then
        vUsuarioAtivo = "admin"
        vNomeOperador = "Administrador"
        vPerfilAtivo = "ADMIN"
        vIDUsuarioAtivo = 1
    End If
End Sub

Public Sub AbrirLogin(Optional ByVal bTrocaUsuario As Boolean = False)
    FormLogin.bModoTroca = bTrocaUsuario
    FormLogin.Show
End Sub

Public Sub TrocarOperador()
    Call AbrirLogin(True)
End Sub

Public Function AutenticarUsuario(ByVal loginOuNome As String, ByVal pinSenha As String) As Boolean
    Dim wsU As Worksheet
    Set wsU = PlanUsuarios()
    
    If wsU Is Nothing Then
        vUsuarioAtivo = "admin"
        vNomeOperador = "Administrador"
        vPerfilAtivo = "ADMIN"
        vIDUsuarioAtivo = 1
        AutenticarUsuario = True
        Call AtualizarBadgesFormsAbertos
        Exit Function
    End If
    
    Dim ultLin As Long, r As Long
    ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
    
    Dim achou As Boolean
    achou = False
    
    For r = 2 To ultLin
        Dim strNome As String, strUser As String, strPass As String, strPerfil As String, strAtivo As String
        Dim idU As Long
        idU = Val(wsU.Cells(r, "A").Value)
        strNome = Trim(wsU.Cells(r, "B").Text)
        strUser = Trim(wsU.Cells(r, "C").Text)
        strPass = Trim(wsU.Cells(r, "D").Text)
        strPerfil = UCase(Trim(wsU.Cells(r, "E").Text))
        strAtivo = UCase(Trim(wsU.Cells(r, "F").Text))
        
        If (LCase(strUser) = LCase(loginOuNome) Or LCase(strNome) = LCase(loginOuNome)) And strAtivo = "SIM" Then
            If strPass = Trim(pinSenha) Then
                vIDUsuarioAtivo = idU
                vUsuarioAtivo = strUser
                vNomeOperador = strNome
                vPerfilAtivo = strPerfil
                AutenticarUsuario = True
                achou = True
                Exit For
            End If
        End If
    Next r
    
    If achou Then
        AutenticarUsuario = True
        Call AtualizarBadgesFormsAbertos
    Else
        AutenticarUsuario = False
    End If
End Function

Public Function VerificarAutorizacaoAdmin(ByVal operacaoDescricao As String, Optional ByVal senhaDireta As String = "") As Boolean
    If UCase(Trim(vPerfilAtivo)) = "ADMIN" Then
        VerificarAutorizacaoAdmin = True
        Exit Function
    End If
    
    Dim pinInformado As String
    If senhaDireta <> "" Then
        pinInformado = senhaDireta
    Else
        pinInformado = InputBox("Esta operação exige autorização de Administrador!" & vbCrLf & vbCrLf & _
                                "• Ação: " & operacaoDescricao & vbCrLf & vbCrLf & _
                                "Digite o PIN de 4 dígitos do Administrador:", "Autorização Administrativa")
    End If
    
    If Trim(pinInformado) = "" Then
        VerificarAutorizacaoAdmin = False
        Exit Function
    End If
    
    Dim wsU As Worksheet
    Set wsU = PlanUsuarios()
    If wsU Is Nothing Then
        VerificarAutorizacaoAdmin = (pinInformado = "1234")
        Exit Function
    End If
    
    Dim ultLin As Long, r As Long
    ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
    For r = 2 To ultLin
        If UCase(Trim(wsU.Cells(r, "E").Text)) = "ADMIN" And UCase(Trim(wsU.Cells(r, "F").Text)) = "SIM" Then
            If Trim(wsU.Cells(r, "D").Text) = Trim(pinInformado) Then
                VerificarAutorizacaoAdmin = True
                Exit Function
            End If
        End If
    Next r
    
    MsgBox "PIN de Administrador incorreto! Operação cancelada.", vbCritical, "Acesso Negado"
    VerificarAutorizacaoAdmin = False
End Function

Public Sub AtualizarBadgesFormsAbertos()
    On Error Resume Next
    Dim textoBadge As String
    textoBadge = "👤 Operador: " & vNomeOperador & " (" & IIf(vPerfilAtivo = "ADMIN", "Admin", "Recepção") & ")"
    
    ' Atualiza FormVendas se aberto
    FormVendas.LblOperadorBadge.Caption = textoBadge
    If vPerfilAtivo = "ADMIN" Then
        FormVendas.LblOperadorBadge.ForeColor = &H800000
    Else
        FormVendas.LblOperadorBadge.ForeColor = &H7A3B00
    End If
    
    ' Atualiza FormRelatorios se aberto
    FormRelatorios.LblOperadorBadge.Caption = textoBadge
    Call ConfigurarPermissoesRelatorios
    On Error GoTo 0
End Sub

Public Sub ConfigurarPermissoesRelatorios()
    On Error Resume Next
    With FormRelatorios
        If UCase(Trim(vPerfilAtivo)) = "RECEPCAO" Then
            .TxtTotalVendido.Text = "R$ ***"
            .TxtTotalTaxas.Text = "R$ ***"
            .TxtTotalReceber.Text = "R$ ***"
            .TxtTotalVendido.ForeColor = &H7F7F7F
            .TxtTotalTaxas.ForeColor = &H7F7F7F
            .TxtTotalReceber.ForeColor = &H7F7F7F
            .LblCardTotal.Caption = "Total Vendido: (Admin)"
            .LblCardTaxa.Caption = "Taxas Cartão: (Admin)"
            .LblCardReceber.Caption = "LÍQ. A RECEBER: (Admin)"
        Else
            .TxtTotalVendido.ForeColor = 13107200
            .TxtTotalTaxas.ForeColor = 255
            .TxtTotalReceber.ForeColor = 24608
            .LblCardTotal.Caption = "Total Vendido:"
            .LblCardTaxa.Caption = "Total Taxas Cartão:"
            .LblCardReceber.Caption = "LÍQUIDO A RECEBER:"
        End If
    End With
    On Error GoTo 0
End Sub

Public Sub AbrirGerenciamentoUsuarios()
    ' Só abre a planilha de usuários se o Administrador confirmar a senha
    If VerificarAutorizacaoAdmin("Gerenciamento e Acesso à Base de Usuários") Then
        Dim wsU As Worksheet
        Set wsU = PlanUsuarios()
        If Not wsU Is Nothing Then
            wsU.Visible = -1 ' xlSheetVisible
            wsU.Activate
            MsgBox "Aba USUARIOS desocultada com sucesso!" & vbCrLf & vbCrLf & _
                   "• Para alterar senhas ou nomes, edite diretamente as células." & vbCrLf & _
                   "• Ao terminar ou fechar o sistema, ela será novamente ocultada de forma blindada.", vbInformation, "Gerenciamento de Acessos"
        End If
    End If
End Sub
