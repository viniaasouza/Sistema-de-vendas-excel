Attribute VB_Name = "ModFuncoes"
Option Explicit

Public Sub AbrirDespesas()
    Call ModFinanceiro.AbrirPainelDespesas
End Sub

Public Sub AbrirDRE()
    Call ModFinanceiro.IrParaDRE
End Sub

Public Sub AbrirFluxoCaixa()
    Call ModFinanceiro.IrParaFluxoCaixa
End Sub


Public Function InspecionarRuntimeFormProdutos() As String
    Load FormProdutos
    Dim c As Object
    Dim s As String
    For Each c In FormProdutos.Controls
        If TypeName(c) = "Label" Then
            s = s & c.Name & ": Cap=[" & c.Caption & "] L=" & c.Left & " T=" & c.Top & " W=" & c.Width & " H=" & c.Height & " FC=" & c.ForeColor & " BS=" & c.BackStyle & " Vis=" & c.Visible & vbCrLf
        End If
    Next c
    Unload FormProdutos
    InspecionarRuntimeFormProdutos = s
End Function


Public Sub AbrirInventario()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        MsgBox "Acesso Restrito: A visualização da planilha de Inventário é permitida exclusivamente para o Administrador." & vbCrLf & vbCrLf & _
               "Solicite autorização com PIN para visualizar.", vbExclamation, "Visualização Restrita"
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Visualizar Planilha de Inventário") Then
            Exit Sub
        End If
        ThisWorkbook.Sheets("INVENTARIO").Visible = -1
    End If
    On Error Resume Next
    ThisWorkbook.Sheets("INVENTARIO").Activate
    Call ModProdutosServicos.AtualizarInventario
    On Error GoTo 0
End Sub

Public Sub IrParaInicio()
    On Error Resume Next
    ThisWorkbook.Sheets("INÍCIO").Activate
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        Call ModSeguranca.AplicarVisibilidadeAbasPorPerfil
    End If
    On Error GoTo 0
End Sub


Public Sub AbrirProdutosServicos()
    FormProdutos.Show
End Sub


' ========================================================
' MACROS PÚBLICAS DE NAVEGAÇÃO E ACESSO AO SISTEMA (FASE 5)
' ========================================================
Public Sub AbrirPDV()
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        Call ModSeguranca.AbrirLogin(False)
    Else
        FormVendas.Show
    End If
End Sub

Public Sub AbrirRelatorios()
    Call ModRelatorios.AbrirRelatorios
End Sub

Public Sub AbrirProdutos()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        ' Recepção pode consultar catálogo de produtos, mas preço fica travado
        FormProdutos.Show
    Else
        FormProdutos.Show
    End If
End Sub

Public Sub AbrirClientes()
    FormClientes.Show
End Sub

Public Sub AbrirLogin()
    Call ModSeguranca.AbrirLogin(False)
End Sub

Public Sub TrocarOperador()
    Call ModSeguranca.TrocarOperador
End Sub

Public Sub ConfigurarLogo()
    Call ModRelatorios.SelecionarLogoEmpresa
End Sub

' ========================================================
' MOTOR DE CONSULTAS SQL COM ADODB (AULAS 10 A 15)
' ========================================================
Public Function ConexaoExcel() As Object
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")
    Dim strCon As String
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
             ";Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";"
    con.Open strCon
    Set ConexaoExcel = con
End Function

Public Function ExecutarConsultaSQL(ByVal sqlQuery As String) As Object
    On Error GoTo TrataErro
    Dim con As Object, rs As Object
    Set con = ConexaoExcel()
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, con, 3, 1
    Set ExecutarConsultaSQL = rs
    Exit Function
TrataErro:
    Set ExecutarConsultaSQL = Nothing
End Function

Public Function AcSQL(Valor As String) As String
    Dim n As Long, t As String, v As Long
    t = ""
    For n = 1 To VBA.Len(Valor)
        v = VBA.Asc(VBA.Mid(Valor, n, 1))
        Select Case v
            Case 39: t = t & "''"
            Case 65: t = t & "[AÁÀÂÃÄ]"
            Case 67: t = t & "[CÇ]"
            Case 69: t = t & "[EÉÈÊË]"
            Case 73: t = t & "[IÍÌÎÏ]"
            Case 79: t = t & "[OÓÒÔÕÖ]"
            Case 85: t = t & "[UÚÙÛÜ]"
            Case 97: t = t & "[aáàâãä]"
            Case 99: t = t & "[cç]"
            Case 101: t = t & "[eéèêë]"
            Case 105: t = t & "[iíìîï]"
            Case 111: t = t & "[oóòôõö]"
            Case 117: t = t & "[uúùûü]"
            Case Else
                If v > 31 And v < 127 Then
                    t = t & VBA.Chr(v)
                Else
                    t = t & "_"
                End If
        End Select
    Next n
    AcSQL = t
End Function
' ========================================================
' CONVERSOR NUMERICO ROBUSTO E UNIVERSAL (ANTI-ERRO 13)
' ========================================================
Public Function ConverterParaDouble(ByVal v As Variant) As Double
    On Error GoTo TrataErro
    
    If IsError(v) Then
        ConverterParaDouble = 0#
        Exit Function
    End If
    
    Dim s As String
    s = Trim(CStr(v))
    If s = "" Then
        ConverterParaDouble = 0#
        Exit Function
    End If
    
    ' Remover prefixos de moeda, porcentagem, quebras de linha e espacos comuns/invisiveis
    s = Replace(s, "R$", "", 1, -1, vbTextCompare)
    s = Replace(s, "r$", "", 1, -1, vbTextCompare)
    s = Replace(s, "$", "")
    s = Replace(s, "%", "")
    s = Replace(s, Chr(160), "") ' Non-breaking space
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, " ", "")
    s = Trim(s)
    
    If s = "" Then
        ConverterParaDouble = 0#
        Exit Function
    End If
    
    ' Se houver anotacoes entre parenteses como "(1x)" ou "(3,50%)", extrair apenas a parte principal
    Dim posParen As Long
    posParen = InStr(s, "(")
    If posParen > 0 Then
        s = Trim(Left(s, posParen - 1))
    End If
    
    If s = "" Then
        ConverterParaDouble = 0#
        Exit Function
    End If
    
    ' Tratar separadores decimais e de milhares
    Dim posVirgula As Long, posPonto As Long
    posVirgula = InStrRev(s, ",")
    posPonto = InStrRev(s, ".")
    
    Dim sepDecLocal As String
    sepDecLocal = Application.International(xlDecimalSeparator)
    
    If posVirgula > 0 And posPonto > 0 Then
        If posVirgula > posPonto Then
            ' Padrao brasileiro/europeu: 1.250,50 -> remove '.', deixa ','
            s = Replace(s, ".", "")
            s = Replace(s, ",", sepDecLocal)
        Else
            ' Padrao americano: 1,250.50 -> remove ',', deixa '.'
            s = Replace(s, ",", "")
            s = Replace(s, ".", sepDecLocal)
        End If
    ElseIf posVirgula > 0 Then
        s = Replace(s, ",", sepDecLocal)
    ElseIf posPonto > 0 Then
        s = Replace(s, ".", sepDecLocal)
    End If
    
    If IsNumeric(s) Then
        ConverterParaDouble = CDbl(s)
    Else
        s = Replace(s, ",", ".")
        ConverterParaDouble = Val(s)
    End If
    Exit Function

TrataErro:
    ConverterParaDouble = 0#
End Function


