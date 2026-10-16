Attribute VB_Name = "ModVendas"
Option Explicit
Option Compare Text

Public vItemIndexEmEdicao As Long

Public Sub BaixarEstoqueItemVenda(ByVal descProduto As String, ByVal qtdVendida As Long)
    On Error Resume Next
    Dim wsP As Worksheet
    Set wsP = ModProdutosServicos.PlanProdutosServicos()
    If wsP Is Nothing Then Exit Sub
    
    Dim cel As Range
    Set cel = wsP.Range("C:C").Find(What:=Trim(descProduto), LookIn:=xlValues, LookAt:=xlWhole)
    If Not cel Is Nothing Then
        Dim tipoP As String
        tipoP = UCase(Trim(cel.Offset(0, -1).Text))
        If tipoP <> "SERVIÇO" And tipoP <> "SERVICO" Then
            Dim estAtual As Long
            estAtual = Val(cel.Offset(0, 4).Value)
            estAtual = estAtual - qtdVendida
            If estAtual < 0 Then estAtual = 0
            wsP.Unprotect "VendasSeguras2026"
            cel.Offset(0, 4).Value = estAtual
            Call ModProdutosServicos.AplicarProtecaoInventario
        End If
    End If
    On Error GoTo 0
End Sub


Private Function PlanVendas() As Worksheet
    Set PlanVendas = ThisWorkbook.Sheets("VENDAS")
End Function

Private Function PlanItensVenda() As Worksheet
    Set PlanItensVenda = ThisWorkbook.Sheets("ITENS_VENDA")
End Function

Private Function PlanProdutos() As Worksheet
    Set PlanProdutos = ModProdutosServicos.PlanProdutosServicos()
End Function

Private Function PlanClientes() As Worksheet
    Set PlanClientes = ThisWorkbook.Sheets("CLIENTES")
End Function

Sub IniciarNovaVenda()
    vItemIndexEmEdicao = -1
    
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        Call ModSeguranca.InicializarSessaoPadrao
    End If
    Call ModSeguranca.AtualizarBadgesFormsAbertos
    Call ModRelatorios.CarregarLogoEmImage(FormVendas.ImgLogo)
    
    With FormVendas
        .bBloqueiaEventos = True
        
        Dim linV As Long, novoID As Long
        linV = PlanVendas.Cells(PlanVendas.Rows.Count, "A").End(xlUp).Row
        If linV < 2 Then
            novoID = 1
        Else
            novoID = Val(PlanVendas.Cells(linV, "A").Value) + 1
        End If
        .TxtIDVenda.Text = Format(novoID, "0000")
        
        .TxtData.Text = Format(Date, "dd/mm/yyyy")
        
        Call FiltrarClientes("")
        .CboCliente.Text = ""
        
        Call FiltrarProdutos("")
        .CboProduto.Text = ""
        .TxtPrecoUnit.Text = ""
        .TxtQtd.Text = "1"
        .TxtSubtotalItem.Text = ""
        
        .LstItens.Clear
        .LstItens.ColumnCount = 6
        .LstItens.ColumnWidths = "30 pt;35 pt;170 pt;60 pt;40 pt;65 pt"
        .BtRemoverItem.Enabled = False
        
        .CboFormaPagto.Clear
        .CboFormaPagto.AddItem "Dinheiro"
        .CboFormaPagto.AddItem "PIX"
        .CboFormaPagto.AddItem "Cartão de Crédito"
        .CboFormaPagto.AddItem "Cartão de Débito"
        .CboFormaPagto.AddItem "Boleto Bancário"
        .CboFormaPagto.ListIndex = 0
        
        .TxtTotalBruto.Text = "R$ 0,00"
        .TxtDesconto.Text = ""
        .TxtTotalLiquido.Text = "R$ 0,00"
        
        .CboParcelas.Clear
        .CboParcelas.AddItem "1x (à vista)"
        .CboParcelas.ListIndex = 0
        .CboParcelas.Enabled = False
        
        .TxtTaxaCartao.Text = "0,00%"
        .TxtTaxaCartao.Enabled = False
        
        .TxtValorReceber.Text = "R$ 0,00"
        .TxtValorParcela.Text = "R$ 0,00 (1x)"
        .TxtTaxaRS.Text = "R$ 0,00 (0,00%)"
        
        .bBloqueiaEventos = False
        On Error Resume Next
        .CboCliente.SetFocus
        On Error GoTo 0
    End With
End Sub

Sub FormaPagtoAlterada()
    With FormVendas
        .bBloqueiaEventos = True
        
        Dim fp As String
        fp = .CboFormaPagto.Text
        
        If fp = "Cartão de Crédito" Then
            .CboParcelas.Enabled = True
            .TxtTaxaCartao.Enabled = True
            
            .CboParcelas.Clear
            Dim i As Long
            For i = 1 To 12
                If i = 1 Then
                    .CboParcelas.AddItem "1x (à vista)"
                Else
                    .CboParcelas.AddItem i & "x"
                End If
            Next i
            .CboParcelas.ListIndex = 0
            
            .TxtTaxaCartao.Text = "3,50%"
            
        ElseIf fp = "Cartão de Débito" Then
            .CboParcelas.Clear
            .CboParcelas.AddItem "1x (à vista)"
            .CboParcelas.ListIndex = 0
            .CboParcelas.Enabled = False
            
            .TxtTaxaCartao.Enabled = True
            .TxtTaxaCartao.Text = "1,50%"
            
        Else
            .CboParcelas.Clear
            .CboParcelas.AddItem "1x (à vista)"
            .CboParcelas.ListIndex = 0
            .CboParcelas.Enabled = False
            
            .TxtTaxaCartao.Enabled = False
            .TxtTaxaCartao.Text = "0,00%"
        End If
        
        .bBloqueiaEventos = False
        Call RecalcularTotaisVenda
    End With
End Sub

Sub FiltrarClientes(ByVal textoBusca As String)
    Dim ultLin As Long, r As Long
    ultLin = PlanClientes.Cells(PlanClientes.Rows.Count, "A").End(xlUp).Row
    
    With FormVendas.CboCliente
        .Clear
        If ultLin < 2 Then Exit Sub
        
        Dim busca As String
        busca = UCase(Trim(textoBusca))
        
        For r = 2 To ultLin
            Dim idCli As String, nomeCli As String, docCli As String
            idCli = PlanClientes.Cells(r, "A").Text
            nomeCli = PlanClientes.Cells(r, "B").Text
            docCli = PlanClientes.Cells(r, "C").Text
            
            If busca = "" Or InStr(1, UCase(nomeCli), busca) > 0 Or InStr(1, docCli, busca) > 0 Then
                .AddItem nomeCli & " (" & idCli & ")"
            End If
        Next r
    End With
End Sub

Sub FiltrarProdutos(ByVal textoBusca As String)
    Dim ultLin As Long, r As Long
    ultLin = PlanProdutos.Cells(PlanProdutos.Rows.Count, "A").End(xlUp).Row
    
    With FormVendas.CboProduto
        .Clear
        If ultLin < 2 Then Exit Sub
        
        Dim busca As String
        busca = UCase(Trim(textoBusca))
        
        For r = 2 To ultLin
            Dim idProd As String, descProd As String
            idProd = Trim(PlanProdutos.Cells(r, "A").Text)
            descProd = Trim(PlanProdutos.Cells(r, "C").Text)
            
            If idProd <> "" And descProd <> "" Then
                If busca = "" Or InStr(1, UCase(descProd), busca) > 0 Then
                    .AddItem descProd & " (" & idProd & ")"
                End If
            End If
        Next r
    End With
End Sub

Sub LimparCamposItem()
    With FormVendas
        .bBloqueiaEventos = True
        .CboProduto.Text = ""
        .TxtPrecoUnit.Text = ""
        .TxtQtd.Text = "1"
        .TxtSubtotalItem.Text = ""
        .bBloqueiaEventos = False
        vItemIndexEmEdicao = -1
        .BtAdicionarItem.Caption = "ADICIONAR"
        .BtAdicionarItem.BackColor = 16744576
    End With
End Sub

Sub ProdutoSelecionado()
    With FormVendas
        If .CboProduto.ListIndex = -1 And Trim(.CboProduto.Text) = "" Then
            .TxtPrecoUnit.Text = ""
            .TxtSubtotalItem.Text = ""
            Exit Sub
        End If
        
        Dim strProd As String, idProd As String
        strProd = .CboProduto.Text
        
        Dim p1 As Long, p2 As Long
        p1 = InStrRev(strProd, "(")
        p2 = InStrRev(strProd, ")")
        
        Dim linProd As Range
        If p1 > 0 And p2 > p1 Then
            idProd = Mid(strProd, p1 + 1, p2 - p1 - 1)
            Set linProd = PlanProdutos.Range("A:A").Find(What:=idProd, LookIn:=xlValues, LookAt:=xlWhole)
        Else
            Set linProd = PlanProdutos.Range("C:C").Find(What:=Trim(strProd), LookIn:=xlValues, LookAt:=xlWhole)
        End If
        
        If Not linProd Is Nothing Then
            Dim vPrecoVenda As Double
            vPrecoVenda = Val(Replace(PlanProdutos.Cells(linProd.Row, "F").Value, ",", "."))
            .TxtPrecoUnit.Text = Format(vPrecoVenda, "#,##0.00")
            If Val(.TxtQtd.Text) <= 0 Then .TxtQtd.Text = "1"
            Call AtualizarSubtotalItem
        End If
    End With
End Sub

Sub AtualizarSubtotalItem()
    With FormVendas
        Dim preco As Double, qtd As Double, subtotal As Double
        preco = 0
        qtd = 0
        
        Dim pStr As String
        pStr = Replace(.TxtPrecoUnit.Text, "R$", "")
        pStr = Replace(pStr, ".", Application.International(xlDecimalSeparator))
        pStr = Replace(pStr, ",", Application.International(xlDecimalSeparator))
        If IsNumeric(pStr) Then preco = CDbl(pStr)
        
        Dim qStr As String
        qStr = Replace(.TxtQtd.Text, ",", Application.International(xlDecimalSeparator))
        qStr = Replace(qStr, ".", Application.International(xlDecimalSeparator))
        If IsNumeric(qStr) Then qtd = CDbl(qStr)
        
        subtotal = preco * qtd
        .TxtSubtotalItem.Text = Format(subtotal, "R$ #,##0.00")
    End With
End Sub

Sub AdicionarItem()
    With FormVendas
        If Trim(.CboProduto.Text) = "" Then
            MsgBox "Selecione ou digite um produto!", vbExclamation, "Item Inválido"
            .CboProduto.SetFocus
            Exit Sub
        End If
        
        Dim preco As Double, qtd As Double
        preco = 0
        qtd = 0
        
        Dim pStr As String, qStr As String
        pStr = Replace(.TxtPrecoUnit.Text, "R$", "")
        pStr = Replace(pStr, ".", Application.International(xlDecimalSeparator))
        pStr = Replace(pStr, ",", Application.International(xlDecimalSeparator))
        If IsNumeric(pStr) Then preco = CDbl(pStr)
        
        qStr = Replace(.TxtQtd.Text, ",", Application.International(xlDecimalSeparator))
        qStr = Replace(qStr, ".", Application.International(xlDecimalSeparator))
        If IsNumeric(qStr) Then qtd = CDbl(qStr)
        
        If preco <= 0 Then
            MsgBox "Preço unitário inválido!", vbExclamation, "Preço Inválido"
            .TxtPrecoUnit.SetFocus
            Exit Sub
        End If
        
        If qtd <= 0 Then
            MsgBox "Informe uma quantidade maior que zero!", vbExclamation, "Quantidade Inválida"
            .TxtQtd.SetFocus
            Exit Sub
        End If
        
        Dim strProd As String, idProd As String, descProd As String
        strProd = Trim(.CboProduto.Text)
        
        Dim p1 As Long, p2 As Long
        p1 = InStrRev(strProd, "(")
        p2 = InStrRev(strProd, ")")
        
        If p1 > 0 And p2 > p1 Then
            idProd = Mid(strProd, p1 + 1, p2 - p1 - 1)
            descProd = Trim(Left(strProd, p1 - 1))
        Else
            descProd = strProd
            Dim celP As Range
            Set celP = PlanProdutos.Range("C:C").Find(What:=descProd, LookIn:=xlValues, LookAt:=xlWhole)
            If Not celP Is Nothing Then
                idProd = PlanProdutos.Cells(celP.Row, "A").Text
            Else
                idProd = "0"
            End If
        End If
        
        Dim subtotal As Double
        subtotal = preco * qtd
        
        Dim itemIdx As Long
        If vItemIndexEmEdicao >= 0 And vItemIndexEmEdicao < .LstItens.ListCount Then
            itemIdx = vItemIndexEmEdicao
        Else
            .LstItens.AddItem
            itemIdx = .LstItens.ListCount - 1
            .LstItens.List(itemIdx, 0) = Format(itemIdx + 1, "00")
        End If
        
        .LstItens.List(itemIdx, 1) = idProd
        .LstItens.List(itemIdx, 2) = descProd
        .LstItens.List(itemIdx, 3) = Format(preco, "R$ #,##0.00")
        .LstItens.List(itemIdx, 4) = Format(qtd, "0.00")
        .LstItens.List(itemIdx, 5) = Format(subtotal, "R$ #,##0.00")
        
        Call LimparCamposItem
        Call RecalcularTotaisVenda
        .CboProduto.SetFocus
    End With
End Sub

Sub RemoverItem()
    With FormVendas
        If .LstItens.ListIndex = -1 Then
            MsgBox "Selecione um item no carrinho para remover!", vbExclamation, "Remover Item"
            Exit Sub
        End If
        
        Dim idx As Long
        idx = .LstItens.ListIndex
        .LstItens.RemoveItem idx
        
        Dim i As Long
        For i = 0 To .LstItens.ListCount - 1
            .LstItens.List(i, 0) = Format(i + 1, "00")
        Next i
        
        If .LstItens.ListCount = 0 Then
            .BtRemoverItem.Enabled = False
        End If
        
        Call LimparCamposItem
        Call RecalcularTotaisVenda
    End With
End Sub

Sub CarregarItemParaEdicao()
    With FormVendas
        If .LstItens.ListIndex = -1 Then Exit Sub
        
        vItemIndexEmEdicao = .LstItens.ListIndex
        
        .bBloqueiaEventos = True
        Dim desc As String, precoStr As String, qtdStr As String, idProd As String
        idProd = .LstItens.List(vItemIndexEmEdicao, 1)
        desc = .LstItens.List(vItemIndexEmEdicao, 2)
        precoStr = Replace(.LstItens.List(vItemIndexEmEdicao, 3), "R$", "")
        qtdStr = .LstItens.List(vItemIndexEmEdicao, 4)
        
        .CboProduto.Text = desc & " (" & idProd & ")"
        .TxtPrecoUnit.Text = Trim(precoStr)
        .TxtQtd.Text = Trim(qtdStr)
        .bBloqueiaEventos = False
        
        Call AtualizarSubtotalItem
        
        .BtAdicionarItem.Caption = "ATUALIZAR"
        .BtAdicionarItem.BackColor = 32768
        .TxtQtd.SetFocus
    End With
End Sub

Sub RecalcularTotaisVenda()
    With FormVendas
        Dim totBruto As Double, i As Long
        totBruto = 0
        
        For i = 0 To .LstItens.ListCount - 1
            Dim subStr As String, subVal As Double
            subStr = Replace(.LstItens.List(i, 5), "R$", "")
            subStr = Replace(subStr, ".", Application.International(xlDecimalSeparator))
            subStr = Replace(subStr, ",", Application.International(xlDecimalSeparator))
            If IsNumeric(subStr) Then subVal = CDbl(subStr) Else subVal = 0
            totBruto = totBruto + subVal
        Next i
        
        .TxtTotalBruto.Text = Format(totBruto, "R$ #,##0.00")
        
        Dim descVal As Double, descStr As String
        descStr = Replace(.TxtDesconto.Text, "R$", "")
        descStr = Replace(descStr, ".", Application.International(xlDecimalSeparator))
        descStr = Replace(descStr, ",", Application.International(xlDecimalSeparator))
        If IsNumeric(descStr) Then descVal = CDbl(descStr) Else descVal = 0
        
        If descVal > totBruto Then descVal = totBruto
        
        Dim totLiquido As Double
        totLiquido = totBruto - descVal
        .TxtTotalLiquido.Text = Format(totLiquido, "R$ #,##0.00")
        
        ' Cálculo de Parcelamento e Previsão Líquida com Taxa de Cartão
        Dim numParc As Long
        numParc = 1
        If .CboParcelas.Enabled And .CboParcelas.ListIndex >= 0 Then
            Dim strP As String
            strP = .CboParcelas.Text
            If InStr(strP, "x") > 0 Then
                numParc = Val(Left(strP, InStr(strP, "x") - 1))
                If numParc < 1 Then numParc = 1
            End If
        End If
        
        Dim valParcela As Double
        valParcela = totLiquido / numParc
        .TxtValorParcela.Text = Format(valParcela, "R$ #,##0.00") & " (" & numParc & "x)"
        
        Dim taxaPct As Double, taxaStr As String
        taxaStr = Replace(.TxtTaxaCartao.Text, "%", "")
        taxaStr = Replace(taxaStr, ".", Application.International(xlDecimalSeparator))
        taxaStr = Replace(taxaStr, ",", Application.International(xlDecimalSeparator))
        If IsNumeric(taxaStr) Then taxaPct = CDbl(taxaStr) / 100 Else taxaPct = 0
        
        Dim valorTaxaRS As Double, liqReceber As Double
        valorTaxaRS = totLiquido * taxaPct
        liqReceber = totLiquido - valorTaxaRS
        
        .TxtTaxaRS.Text = Format(valorTaxaRS, "R$ #,##0.00") & " (" & Format(taxaPct, "0.00%") & ")"
        .TxtValorReceber.Text = Format(liqReceber, "R$ #,##0.00")
    End With
End Sub

Sub FinalizarVenda()
    With FormVendas
        If Trim(.CboCliente.Text) = "" Then
            MsgBox "Selecione ou digite o Cliente para finalizar a venda!", vbExclamation, "Finalizar Venda"
            .CboCliente.SetFocus
            Exit Sub
        End If
        
        If .LstItens.ListCount = 0 Then
            MsgBox "Adicione pelo menos um item ao carrinho antes de finalizar!", vbExclamation, "Carrinho Vazio"
            .CboProduto.SetFocus
            Exit Sub
        End If
        
        If .CboFormaPagto.ListIndex = -1 Then
            MsgBox "Selecione a Forma de Pagamento!", vbExclamation, "Finalizar Venda"
            .CboFormaPagto.SetFocus
            Exit Sub
        End If
        
        Dim strCliente As String, idCli As String, nomeCli As String
        strCliente = Trim(.CboCliente.Text)
        
        Dim p1 As Long, p2 As Long
        p1 = InStrRev(strCliente, "(")
        p2 = InStrRev(strCliente, ")")
        
        If p1 > 0 And p2 > p1 Then
            idCli = Trim(Mid(strCliente, p1 + 1, p2 - p1 - 1))
            nomeCli = Trim(Left(strCliente, p1 - 1))
        Else
            nomeCli = strCliente
            Dim celC As Range
            Set celC = PlanClientes.Range("B:B").Find(What:=nomeCli, LookIn:=xlValues, LookAt:=xlWhole)
            If Not celC Is Nothing Then
                idCli = celC.Offset(0, -1).Text
            Else
                idCli = "1"
            End If
        End If
        
        Dim idVenda As Long
        idVenda = Val(.TxtIDVenda.Text)
        
        Dim totBruto As Double, totDesc As Double, totLiq As Double
        totBruto = CDbl(Replace(.TxtTotalBruto.Text, "R$", ""))
        totDesc = IIf(IsNumeric(.TxtDesconto.Text), CDbl(.TxtDesconto.Text), 0)
        totLiq = CDbl(Replace(.TxtTotalLiquido.Text, "R$", ""))
        
        Dim strParcSalvar As String
        strParcSalvar = Trim(.CboParcelas.Text)
        If InStr(strParcSalvar, "(") > 0 Then
            strParcSalvar = Trim(Left(strParcSalvar, InStr(strParcSalvar, "(") - 1))
        End If
        If strParcSalvar = "" Then strParcSalvar = "1x"
        
        Dim taxaPctSalvar As Double, taxaStr As String
        taxaStr = Trim(Replace(.TxtTaxaCartao.Text, "%", ""))
        taxaStr = Replace(taxaStr, ".", Application.International(xlDecimalSeparator))
        taxaStr = Replace(taxaStr, ",", Application.International(xlDecimalSeparator))
        If IsNumeric(taxaStr) Then
            taxaPctSalvar = CDbl(taxaStr) / 100
        Else
            taxaPctSalvar = 0
        End If
        
        Dim liqReceberSalvar As Double
        liqReceberSalvar = CDbl(Replace(.TxtValorReceber.Text, "R$", ""))
        
        Dim opSalvar As String
        opSalvar = Trim(ModSeguranca.vNomeOperador)
        If opSalvar = "" Then opSalvar = "Recepção Geral"
        
        Dim msgConf As String
        If .CboFormaPagto.Text = "Cartão de Crédito" And InStr(.CboParcelas.Text, "1x") = 0 Then
            msgConf = "Confirmar finalização da Venda Nº " & .TxtIDVenda.Text & "?" & vbCrLf & vbCrLf & _
                      "• Operador: " & opSalvar & vbCrLf & _
                      "• Total Cliente: " & .TxtTotalLiquido.Text & " em " & .TxtValorParcela.Text & vbCrLf & _
                      "• Previsão Líquida a Receber: " & .TxtValorReceber.Text & " (" & .TxtTaxaRS.Text & ")"
        Else
            msgConf = "Confirmar finalização da Venda Nº " & .TxtIDVenda.Text & " no valor de " & .TxtTotalLiquido.Text & "?" & vbCrLf & _
                      "• Operador: " & opSalvar & vbCrLf & _
                      "• Forma de Pagamento: " & .CboFormaPagto.Text & vbCrLf & _
                      "• Previsão Líquida a Receber: " & .TxtValorReceber.Text
        End If
        
        If MsgBox(msgConf, vbQuestion + vbYesNo, "Confirmar Venda") = vbNo Then Exit Sub
        
        ' Gravar Cabeçalho da Venda na aba VENDAS
        Dim linV As Long
        linV = PlanVendas.Cells(PlanVendas.Rows.Count, "A").End(xlUp).Row + 1
        
        PlanVendas.Cells(linV, "A").Value = idVenda
        PlanVendas.Cells(linV, "B").Value = CDate(.TxtData.Text)
        PlanVendas.Cells(linV, "C").Value = idCli
        PlanVendas.Cells(linV, "D").Value = nomeCli
        PlanVendas.Cells(linV, "E").Value = totBruto
        PlanVendas.Cells(linV, "F").Value = totDesc
        PlanVendas.Cells(linV, "G").Value = totLiq
        PlanVendas.Cells(linV, "H").Value = .CboFormaPagto.Text
        PlanVendas.Cells(linV, "I").Value = "Concluída"
        PlanVendas.Cells(linV, "J").Value = strParcSalvar
        PlanVendas.Cells(linV, "K").Value = taxaPctSalvar
        PlanVendas.Cells(linV, "L").Value = liqReceberSalvar
        PlanVendas.Cells(linV, "M").Value = opSalvar ' COLUNA 13: OPERADOR
        
        PlanVendas.Cells(linV, "K").NumberFormat = "0.00%"
        PlanVendas.Cells(linV, "L").NumberFormat = "R$ #,##0.00"
        
        ' Gravar Itens na aba ITENS_VENDA
        Dim i As Long, linItem As Long, ultIDItem As Long
        ultIDItem = PlanItensVenda.Cells(PlanItensVenda.Rows.Count, "A").End(xlUp).Row - 1
        If ultIDItem < 0 Then ultIDItem = 0
        
        For i = 0 To .LstItens.ListCount - 1
            linItem = PlanItensVenda.Cells(PlanItensVenda.Rows.Count, "A").End(xlUp).Row + 1
            ultIDItem = ultIDItem + 1
            
            Dim idProd As String, descProd As String, prcItem As Double, qtdItem As Double, subItem As Double
            idProd = .LstItens.List(i, 1)
            descProd = .LstItens.List(i, 2)
            prcItem = CDbl(Replace(.LstItens.List(i, 3), "R$", ""))
            qtdItem = CDbl(.LstItens.List(i, 4))
            subItem = CDbl(Replace(.LstItens.List(i, 5), "R$", ""))
            
            PlanItensVenda.Cells(linItem, "A").Value = ultIDItem
            PlanItensVenda.Cells(linItem, "B").Value = idVenda
            PlanItensVenda.Cells(linItem, "C").Value = idProd
            PlanItensVenda.Cells(linItem, "D").Value = descProd
            PlanItensVenda.Cells(linItem, "E").Value = prcItem
            PlanItensVenda.Cells(linItem, "F").Value = qtdItem
            PlanItensVenda.Cells(linItem, "G").Value = subItem
            
            PlanItensVenda.Cells(linItem, "E").NumberFormat = "R$ #,##0.00"
            PlanItensVenda.Cells(linItem, "G").NumberFormat = "R$ #,##0.00"
            Call BaixarEstoqueItemVenda(descProd, CLng(qtdItem))
        Next i
        
        ' Pergunta de Emissão de Comprovante em PDF
        Dim pergCupom As VbMsgBoxResult
        pergCupom = MsgBox("Venda Nº " & Format(idVenda, "0000") & " finalizada com sucesso por " & opSalvar & "!" & vbCrLf & vbCrLf & _
                           "Deseja emitir o Comprovante de Venda em PDF agora?", vbQuestion + vbYesNo, "Comprovante de Venda")
                           
        If pergCupom = vbYes Then
            Call ModRelatorios.GerarComprovanteVendaPorID(idVenda)
        End If
        
        Call IniciarNovaVenda
    End With
End Sub






