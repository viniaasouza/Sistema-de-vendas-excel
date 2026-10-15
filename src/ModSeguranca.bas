Attribute VB_Name = "ModSeguranca"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE SEGURANÇA, CONTROLE DE ACESSO E OPERADORES (FASE 5)
' =============================================================
Public vUsuarioAtivo As String      ' Ex: "admin", "recepcao1", "recepcao2"
Public vNomeOperador As String      ' Ex: "Administrador", "Ana Silva", "Bruna Costa"
Public vPerfilAtivo As String       ' "ADMIN" ou "RECEPCAO"
Public vIDUsuarioAtivo As Long

Public Function PlanUsuarios() As Worksheet
    On Error Resume Next
    Set PlanUsuarios = ThisWorkbook.Sheets("USUARIOS")
    On Error GoTo 0
End Function

Public Sub InicializarSessaoPadrao()
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
        
        If (strNome = loginOuNome Or strUser = loginOuNome) And strPass = pinSenha Then
            If strAtivo = "NÃO" Or strAtivo = "NAO" Then
                MsgBox "Este operador está inativo no sistema! Contate o administrador.", vbExclamation, "Acesso Bloqueado"
                AutenticarUsuario = False
                Exit Function
            End If
            
            vIDUsuarioAtivo = idU
            vNomeOperador = strNome
            vUsuarioAtivo = strUser
            vPerfilAtivo = strPerfil
            achou = True
            Exit For
        End If
    Next r
    
    If achou Then
        AutenticarUsuario = True
        Call AtualizarBadgesFormsAbertos
    Else
        AutenticarUsuario = False
    End If
End Function

Public Function VerificarAutorizacaoAdmin(ByVal motivoAcao As String) As Boolean
    ' Perfil ADMIN possui autorização imediata irrestrita
    If UCase(Trim(vPerfilAtivo)) = "ADMIN" Then
        VerificarAutorizacaoAdmin = True
        Exit Function
    End If
    
    ' Perfil RECEPCAO exige validação da senha do Administrador
    Dim senhaDigitada As String
    senhaDigitada = InputBox("AÇÃO RESTRITA AO ADMINISTRADOR:" & vbCrLf & _
                             "• " & motivoAcao & vbCrLf & vbCrLf & _
                             "Operador atual: " & vNomeOperador & " (Recepção)" & vbCrLf & vbCrLf & _
                             "Digite a senha/PIN de autorização do Administrador:", _
                             "Autorização Administrativa Necessária")
                             
    If Trim(senhaDigitada) = "" Then
        MsgBox "Ação cancelada pelo usuário ou senha não informada.", vbInformation, "Autorização Cancelada"
        VerificarAutorizacaoAdmin = False
        Exit Function
    End If
    
    Dim wsU As Worksheet
    Set wsU = PlanUsuarios()
    If wsU Is Nothing Then
        If senhaDigitada = "1234" Or senhaDigitada = "admin" Then
            VerificarAutorizacaoAdmin = True
            Exit Function
        Else
            MsgBox "Senha de autorização incorreta!", vbCritical, "Acesso Negado"
            VerificarAutorizacaoAdmin = False
            Exit Function
        End If
    End If
    
    Dim ultLin As Long, r As Long, autorizou As Boolean
    ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
    autorizou = False
    
    For r = 2 To ultLin
        If UCase(Trim(wsU.Cells(r, "E").Text)) = "ADMIN" And _
           UCase(Trim(wsU.Cells(r, "F").Text)) = "SIM" Then
            If Trim(wsU.Cells(r, "D").Text) = Trim(senhaDigitada) Then
                autorizou = True
                Exit For
            End If
        End If
    Next r
    
    If autorizou Then
        VerificarAutorizacaoAdmin = True
    Else
        MsgBox "Senha de autorização de Administrador inválida!", vbCritical, "Acesso Negado"
        VerificarAutorizacaoAdmin = False
    End If
End Function

Public Sub AtualizarBadgesFormsAbertos()
    On Error Resume Next
    Dim textoBadge As String
    textoBadge = "?? Operador: " & vNomeOperador & " (" & IIf(vPerfilAtivo = "ADMIN", "Admin", "Recepção") & ")"
    
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


