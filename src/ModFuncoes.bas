Attribute VB_Name = "ModFuncoes"
Option Explicit

Public Function TestarPDVFluxoCompleto() As String
    On Error GoTo TrataErro
    
    ' 1. Configurar sessao da Vitoria Karollyne (RECEPCAO)
    ModSeguranca.vUsuarioAtivo = "vitoria"
    ModSeguranca.vNomeOperador = "Vitória Karollyne Bezerra da Fonseca"
    ModSeguranca.vPerfilAtivo = "RECEPCAO"
    
    ' 2. Carregar FormVendas e Iniciar PDV
    Load FormVendas
    Call ModVendas.IniciarNovaVenda
    
    ' 3. Preencher dados da venda
    FormVendas.CboCliente.Text = "Carlos Alberto da Silva (1)"
    FormVendas.CboProduto.ListIndex = 0
    Call ModVendas.ProdutoSelecionado
    Call ModVendas.AdicionarItem
    
    FormVendas.CboFormaPagto.Text = "Cartão de Crédito"
    Call ModVendas.FormaPagtoAlterada
    
    ' DESCONTO VAZIO ("") intencional
    FormVendas.TxtDesconto.Text = ""
    Call ModVendas.RecalcularTotaisVenda
    
    ' 4. Extrair valores com ConverterParaDouble
    Dim bVal As Double, dVal As Double, lVal As Double, tVal As Double, rVal As Double
    bVal = ConverterParaDouble(FormVendas.TxtTotalBruto.Text)
    dVal = ConverterParaDouble(FormVendas.TxtDesconto.Text)
    lVal = ConverterParaDouble(FormVendas.TxtTotalLiquido.Text)
    tVal = ConverterParaDouble(FormVendas.TxtTaxaCartao.Text) / 100
    rVal = ConverterParaDouble(FormVendas.TxtValorReceber.Text)
    
    If bVal <= 0 Or lVal <= 0 Then
        Unload FormVendas
        TestarPDVFluxoCompleto = "ERRO: Valores invalidos Bruto=" & bVal & " Liq=" & lVal
        Exit Function
    End If
    
    ' 5. Gravar na aba VENDAS
    Dim wsV As Worksheet, wsI As Worksheet
    Set wsV = ThisWorkbook.Sheets("VENDAS")
    Set wsI = ThisWorkbook.Sheets("ITENS_VENDA")
    
    Dim idVenda As Long, linV As Long
    idVenda = Val(FormVendas.TxtIDVenda.Text)
    linV = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row + 1
    
    wsV.Cells(linV, "A").Value = idVenda
    wsV.Cells(linV, "B").Value = Date
    wsV.Cells(linV, "C").Value = "1"
    wsV.Cells(linV, "D").Value = "Carlos Alberto da Silva"
    wsV.Cells(linV, "E").Value = bVal
    wsV.Cells(linV, "F").Value = dVal
    wsV.Cells(linV, "G").Value = lVal
    wsV.Cells(linV, "H").Value = FormVendas.CboFormaPagto.Text
    wsV.Cells(linV, "I").Value = "Concluída"
    wsV.Cells(linV, "J").Value = "1x"
    wsV.Cells(linV, "K").Value = tVal
    wsV.Cells(linV, "L").Value = rVal
    wsV.Cells(linV, "M").Value = ModSeguranca.vNomeOperador
    
    ' Gravar na aba ITENS_VENDA
    Dim linI As Long
    linI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row + 1
    wsI.Cells(linI, "A").Value = linI - 1
    wsI.Cells(linI, "B").Value = idVenda
    wsI.Cells(linI, "C").Value = FormVendas.LstItens.List(0, 1)
    wsI.Cells(linI, "D").Value = FormVendas.LstItens.List(0, 2)
    wsI.Cells(linI, "E").Value = ConverterParaDouble(FormVendas.LstItens.List(0, 3))
    wsI.Cells(linI, "F").Value = ConverterParaDouble(FormVendas.LstItens.List(0, 4))
    wsI.Cells(linI, "G").Value = ConverterParaDouble(FormVendas.LstItens.List(0, 5))
    
    ' Baixar estoque de teste
    Call ModVendas.BaixarEstoqueItemVenda(CStr(FormVendas.LstItens.List(0, 2)), 1)
    
    ' Ler confirmacao
    Dim cID As Long, cCli As String, cBruto As Double, cDesc As Double, cLiq As Double, cOp As String, cStatus As String
    cID = wsV.Cells(linV, "A").Value
    cCli = wsV.Cells(linV, "D").Value
    cBruto = wsV.Cells(linV, "E").Value
    cDesc = wsV.Cells(linV, "F").Value
    cLiq = wsV.Cells(linV, "G").Value
    cStatus = wsV.Cells(linV, "I").Value
    cOp = wsV.Cells(linV, "M").Value
    
    ' Reverter o estoque baixado no teste e deletar linhas
    Call ModVendas.BaixarEstoqueItemVenda(CStr(FormVendas.LstItens.List(0, 2)), -1)
    wsV.Rows(linV).Delete
    wsI.Rows(linI).Delete
    
    Unload FormVendas
    
    TestarPDVFluxoCompleto = "OK: ID=" & cID & " | Cliente=" & cCli & " | Bruto=" & cBruto & _
                             " | Desconto=" & cDesc & " | Liq=" & cLiq & " | Op=" & cOp & _
                             " | Status=" & cStatus
    Exit Function

TrataErro:
    On Error Resume Next
    Unload FormVendas
    TestarPDVFluxoCompleto = "ERRO (" & Err.Number & "): " & Err.Description
End Function


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
