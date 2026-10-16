Attribute VB_Name = "ModSeguranca"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE SEGURANÇA, CONTROLE DE ACESSO E OPERADORES (FASE 8)
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
    Call ProtegerAbaUsuarios
    FormLogin.bModoTroca = bTrocaUsuario
    FormLogin.Show
End Sub

Public Sub TrocarOperador()
    Call ProtegerAbaUsuarios
    Call AbrirLogin(True)
    Call AplicarVisibilidadeAbasPorPerfil
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
        Call ProtegerAbaUsuarios
        Call AplicarVisibilidadeAbasPorPerfil
        Call AtualizarBadgesFormsAbertos
        Exit Function
    End If
    
    Dim rInicio As Long
    Dim celID As Range
    Set celID = wsU.Range("A:A").Find(What:="ID", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celID Is Nothing Then rInicio = celID.Row + 1 Else rInicio = 5
    
    Dim ultLin As Long, r As Long
    ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
    
    Dim achou As Boolean
    achou = False
    
    For r = rInicio To ultLin
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
        Call ProtegerAbaUsuarios
        Call AplicarVisibilidadeAbasPorPerfil
        Call AtualizarBadgesFormsAbertos
        
        If UCase(Trim(vPerfilAtivo)) = "RECEPCAO" Then
            On Error Resume Next
            ThisWorkbook.Sheets("INÍCIO").Activate
            On Error GoTo 0
        End If
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
        pinInformado = SolicitarPINAdmin(operacaoDescricao)
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
    
    Dim rInicio As Long
    Dim celID As Range
    Set celID = wsU.Range("A:A").Find(What:="ID", LookIn:=xlValues, LookAt:=xlWhole)
    If Not celID Is Nothing Then rInicio = celID.Row + 1 Else rInicio = 5
    
    Dim ultLin As Long, r As Long
    ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
    For r = rInicio To ultLin
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
    textoBadge = "?? Operador: " & vNomeOperador & " (" & IIf(vPerfilAtivo = "ADMIN", "Admin", "Recepção") & ")"
    
    FormVendas.LblOperadorBadge.Caption = textoBadge
    If vPerfilAtivo = "ADMIN" Then
        FormVendas.LblOperadorBadge.ForeColor = &H800000
    Else
        FormVendas.LblOperadorBadge.ForeColor = &H7A3B00
    End If
    
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
    If VerificarAutorizacaoAdmin("Gerenciamento e Acesso à Base de Usuários") Then
        Dim wsU As Worksheet
        Set wsU = PlanUsuarios()
        If Not wsU Is Nothing Then
            wsU.Visible = -1 ' xlSheetVisible
            wsU.Activate
            MsgBox "Aba USUARIOS desocultada com sucesso!" & vbCrLf & vbCrLf & _
                   "• Para alterar senhas ou nomes, edite diretamente as células." & vbCrLf & _
                   "• Ao trocar de operador ou sair desta aba, ela se OCULTARÁ AUTOMATICAMENTE.", vbInformation, "Gerenciamento de Acessos"
        End If
    End If
End Sub

Public Sub FecharEProtegerUsuarios()
    Call ProtegerAbaUsuarios
    On Error Resume Next
    ThisWorkbook.Sheets("INÍCIO").Activate
    On Error GoTo 0
    MsgBox "Planilha de Usuários blindada e ocultada com sucesso!", vbInformation, "Segurança Ativa"
End Sub




Public Sub AplicarVisibilidadeAbasPorPerfil()
    On Error Resume Next
    Dim isAdmin As Boolean
    isAdmin = (UCase(Trim(vPerfilAtivo)) = "ADMIN")
    
    ' INÍCIO permanece sempre visível como painel principal
    ThisWorkbook.Sheets("INÍCIO").Visible = -1 ' xlSheetVisible
    ThisWorkbook.Sheets("INÍCIO").Activate
    
    ' Planilhas restritas exclusivamente para o Administrador
    Dim abasRestritas As Variant
    abasRestritas = Array("VENDAS", "CONFIG", "CLIENTES", "INVENTARIO", "PRODUTOS_SERVICOS", "ITENS_VENDA", "R_Vendas")
    
    Dim i As Long, nomeAba As String, ws As Worksheet
    For i = LBound(abasRestritas) To UBound(abasRestritas)
        nomeAba = CStr(abasRestritas(i))
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(nomeAba)
        If Not ws Is Nothing Then
            If isAdmin Then
                ws.Visible = -1 ' xlSheetVisible (Admin pode visualizar)
            Else
                ws.Visible = 2  ' xlSheetVeryHidden (Totalmente invisivel para Recepcao)
            End If
        End If
    Next i
    
    ' USUARIOS sempre permanece VeryHidden por padrão para todos
    Call ProtegerAbaUsuarios
    
    ' Proteção de células do inventário
    Call ModProdutosServicos.AplicarProtecaoInventario
    On Error GoTo 0
End Sub

Public Function SolicitarPINAdmin(ByVal operacaoDescricao As String) As String
    Load FormPIN
    FormPIN.LblOperacao.Caption = "Ação: " & operacaoDescricao
    FormPIN.TxtPIN.PasswordChar = "*"
    FormPIN.TxtPIN.Text = ""
    FormPIN.bConfirmado = False
    FormPIN.sPINDigitado = ""
    FormPIN.Show
    
    If FormPIN.bConfirmado Then
        SolicitarPINAdmin = Trim(FormPIN.sPINDigitado)
    Else
        SolicitarPINAdmin = ""
    End If
    Unload FormPIN
End Function


