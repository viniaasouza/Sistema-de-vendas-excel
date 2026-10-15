Attribute VB_Name = "ModVendas"
Option Explicit
Option Compare Text

Public vItemIndexEmEdicao As Long

Private Function PlanVendas() As Worksheet
    Set PlanVendas = ThisWorkbook.Sheets("VENDAS")
End Function

Private Function PlanItensVenda() As Worksheet
    Set PlanItensVenda = ThisWorkbook.Sheets("ITENS_VENDA")
End Function

Private Function PlanProdutos() As Worksheet
    Set PlanProdutos = ThisWorkbook.Sheets("PRODUTOS")
End Function

Private Function PlanClientes() As Worksheet
    Set PlanClientes = ThisWorkbook.Sheets("CLIENTES")
End Function

Sub IniciarNovaVenda()
    vItemIndexEmEdicao = -1
    
    With FormVendas
        ' 1. Gerar próximo ID de venda sequencial
        Dim linV As Long, novoID As Long
        linV = PlanVendas.Cells(PlanVendas.Rows.Count, "A").End(xlUp).Row
        If linV < 2 Then
            novoID = 1
        Else
            novoID = Val(PlanVendas.Cells(linV, "A").Value) + 1
        End If
        .TxtIDVenda.Text = Format(novoID, "0000")
        
        ' 2. Data atual do sistema
        .TxtData.Text = Format(Date, "dd/mm/yyyy")
        
        ' 3. Preencher Combo de Clientes
        .CboCliente.Clear
        Dim linC As Long, ultLinC As Long
        ultLinC = PlanClientes.Cells(PlanClientes.Rows.Count, "A").End(xlUp).Row
        For linC = 2 To ultLinC
            If Trim(PlanClientes.Cells(linC, "B").Text) <> "" Then
                .CboCliente.AddItem PlanClientes.Cells(linC, "A").Text & " - " & PlanClientes.Cells(linC, "B").Text
            End If
        Next linC
        
        ' 4. Preencher Combo de Produtos
        .CboProduto.Clear
        Dim linP As Long, ultLinP As Long
        ultLinP = PlanProdutos.Cells(PlanProdutos.Rows.Count, "A").End(xlUp).Row
        For linP = 2 To ultLinP
            If Trim(PlanProdutos.Cells(linP, "B").Text) <> "" Then
                .CboProduto.AddItem PlanProdutos.Cells(linP, "A").Text & " - " & PlanProdutos.Cells(linP, "B").Text
            End If
        Next linP
        
        ' 5. Formas de Pagamento
        .CboFormaPagto.Clear
        .CboFormaPagto.AddItem "Dinheiro"
        .CboFormaPagto.AddItem "PIX"
        .CboFormaPagto.AddItem "Cartão de Crédito"
        .CboFormaPagto.AddItem "Cartão de Débito"
        .CboFormaPagto.AddItem "Boleto Bancário"
        .CboFormaPagto.ListIndex = 1 ' PIX como padrão
        
        ' 6. Resetar campos de item e totais
        Call LimparCamposItem
        .LstItens.Clear
        .TxtTotalBruto.Text = "R$ 0,00"
        .TxtDesconto.Text = "0,00"
        .TxtTotalLiquido.Text = "R$ 0,00"
    End With
End Sub

Sub LimparCamposItem()
    vItemIndexEmEdicao = -1
    With FormVendas
        .CboProduto.ListIndex = -1
        .TxtPrecoUnit.Text = ""
        .TxtQtd.Text = "1"
        .TxtSubtotalItem.Text = ""
        .BtAdicionarItem.Caption = "ADICIONAR"
        On Error Resume Next
        .BtRemoverItem.Enabled = False
        On Error GoTo 0
    End With
End Sub

Sub ProdutoSelecionado()
    With FormVendas
        If .CboProduto.ListIndex = -1 Then Exit Sub
        
        Dim partes() As String, idProd As String
        partes = Split(.CboProduto.Text, " - ")
        idProd = Trim(partes(0))
        
        Dim cel As Range
        Set cel = PlanProdutos.Range("A:A").Find(What:=idProd, LookIn:=xlValues, LookAt:=xlWhole)
        If Not cel Is Nothing Then
            .TxtPrecoUnit.Text = Format(cel.Offset(0, 2).Value, "#,##0.00")
            Call AtualizarSubtotalItem
        End If
    End With
End Sub

Sub AtualizarSubtotalItem()
    With FormVendas
        Dim pStr As String, qStr As String
        pStr = Replace(.TxtPrecoUnit.Text, "R$", "")
        pStr = Trim(pStr)
        qStr = Trim(.TxtQtd.Text)
        
        If IsNumeric(pStr) And IsNumeric(qStr) Then
            Dim subt As Double
            subt = CDbl(pStr) * CDbl(qStr)
            .TxtSubtotalItem.Text = Format(subt, "#,##0.00")
        Else
            .TxtSubtotalItem.Text = ""
        End If
    End With
End Sub

Sub AdicionarItem()
    With FormVendas
        If .CboProduto.ListIndex = -1 Then
            MsgBox "Selecione um produto para adicionar!", vbExclamation, "Item da Venda"
            .CboProduto.SetFocus
            Exit Sub
        End If
        
        If Val(.TxtQtd.Text) <= 0 Then
            MsgBox "Informe uma quantidade válida maior que zero!", vbExclamation, "Item da Venda"
            .TxtQtd.SetFocus
            Exit Sub
        End If
        
        Dim partes() As String, idProd As String, descProd As String
        partes = Split(.CboProduto.Text, " - ")
        idProd = Trim(partes(0))
        descProd = Trim(partes(1))
        
        Dim pUnit As Double, qtd As Double, subt As Double
        pUnit = CDbl(Replace(.TxtPrecoUnit.Text, "R$", ""))
        qtd = CDbl(.TxtQtd.Text)
        subt = pUnit * qtd
        
        If vItemIndexEmEdicao >= 0 And vItemIndexEmEdicao < .LstItens.ListCount Then
            ' Atualiza item que estava em edição
            .LstItens.List(vItemIndexEmEdicao, 1) = idProd
            .LstItens.List(vItemIndexEmEdicao, 2) = descProd
            .LstItens.List(vItemIndexEmEdicao, 3) = Format(pUnit, "R$ #,##0.00")
            .LstItens.List(vItemIndexEmEdicao, 4) = qtd
            .LstItens.List(vItemIndexEmEdicao, 5) = Format(subt, "R$ #,##0.00")
        Else
            ' Novo item na lista do carrinho
            Dim idx As Long
            idx = .LstItens.ListCount
            .LstItens.AddItem
            .LstItens.List(idx, 0) = idx + 1
            .LstItens.List(idx, 1) = idProd
            .LstItens.List(idx, 2) = descProd
            .LstItens.List(idx, 3) = Format(pUnit, "R$ #,##0.00")
            .LstItens.List(idx, 4) = qtd
            .LstItens.List(idx, 5) = Format(subt, "R$ #,##0.00")
        End If
        
        Call LimparCamposItem
        Call RecalcularTotaisVenda
    End With
End Sub

Sub RemoverItem()
    With FormVendas
        If .LstItens.ListIndex = -1 Then
            MsgBox "Selecione um item no carrinho para remover!", vbExclamation, "Atenção"
            Exit Sub
        End If
        
        If MsgBox("Deseja remover este item da venda?", vbQuestion + vbYesNo, "Remover Item") = vbYes Then
            .LstItens.RemoveItem .LstItens.ListIndex
            
            ' Renumerar linhas da coluna 0
            Dim i As Long
            For i = 0 To .LstItens.ListCount - 1
                .LstItens.List(i, 0) = i + 1
            Next i
            
            Call LimparCamposItem
            Call RecalcularTotaisVenda
        End If
    End With
End Sub

Sub CarregarItemParaEdicao()
    With FormVendas
        If .LstItens.ListIndex = -1 Then Exit Sub
        vItemIndexEmEdicao = .LstItens.ListIndex
        
        Dim idProd As String, i As Long
        idProd = .LstItens.List(vItemIndexEmEdicao, 1)
        
        For i = 0 To .CboProduto.ListCount - 1
            If .CboProduto.List(i) Like idProd & " - *" Then
                .CboProduto.ListIndex = i
                Exit For
            End If
        Next i
        
        Dim pStr As String
        pStr = Replace(.LstItens.List(vItemIndexEmEdicao, 3), "R$", "")
        .TxtPrecoUnit.Text = Trim(pStr)
        .TxtQtd.Text = .LstItens.List(vItemIndexEmEdicao, 4)
        Call AtualizarSubtotalItem
        
        .BtAdicionarItem.Caption = "ALTERAR"
        On Error Resume Next
        .BtRemoverItem.Enabled = True
        On Error GoTo 0
    End With
End Sub

Sub RecalcularTotaisVenda()
    With FormVendas
        Dim i As Long, totalBruto As Double
        totalBruto = 0
        
        For i = 0 To .LstItens.ListCount - 1
            Dim vSubt As Double, subtStr As String
            subtStr = Replace(.LstItens.List(i, 5), "R$", "")
            subtStr = Trim(subtStr)
            If IsNumeric(subtStr) Then
                totalBruto = totalBruto + CDbl(subtStr)
            End If
        Next i
        
        .TxtTotalBruto.Text = Format(totalBruto, "R$ #,##0.00")
        
        Dim descVal As Double
        descVal = 0
        If IsNumeric(.TxtDesconto.Text) Then
            descVal = CDbl(.TxtDesconto.Text)
        End If
        
        Dim totalLiq As Double
        totalLiq = totalBruto - descVal
        If totalLiq < 0 Then totalLiq = 0
        
        .TxtTotalLiquido.Text = Format(totalLiq, "R$ #,##0.00")
    End With
End Sub

Sub FinalizarVenda()
    With FormVendas
        ' 1. Validações
        If .CboCliente.ListIndex = -1 Then
            MsgBox "Selecione o Cliente para finalizar a venda!", vbExclamation, "Finalizar Venda"
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
        
        ' Confirmação com usuário
        Dim msgConf As String
        msgConf = "Confirmar finalização da Venda Nº " & .TxtIDVenda.Text & " no valor de " & .TxtTotalLiquido.Text & "?"
        If MsgBox(msgConf, vbQuestion + vbYesNo, "Confirmar Venda") = vbNo Then Exit Sub
        
        ' 2. Dados do Cliente e Venda
        Dim partesC() As String, idCli As String, nomeCli As String
        partesC = Split(.CboCliente.Text, " - ")
        idCli = Trim(partesC(0))
        nomeCli = Trim(partesC(1))
        
        Dim idVenda As Long
        idVenda = Val(.TxtIDVenda.Text)
        
        Dim totBruto As Double, totDesc As Double, totLiq As Double
        totBruto = CDbl(Replace(.TxtTotalBruto.Text, "R$", ""))
        totDesc = IIf(IsNumeric(.TxtDesconto.Text), CDbl(.TxtDesconto.Text), 0)
        totLiq = CDbl(Replace(.TxtTotalLiquido.Text, "R$", ""))
        
        ' 3. Gravar Cabeçalho da Venda na aba VENDAS
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
        
        ' 4. Gravar Itens na aba ITENS_VENDA
        Dim i As Long, linItem As Long, ultIDItem As Long
        ultIDItem = PlanItensVenda.Cells(PlanItensVenda.Rows.Count, "A").End(xlUp).Row - 1
        If ultIDItem < 0 Then ultIDItem = 0
        
        For i = 0 To .LstItens.ListCount - 1
            linItem = PlanItensVenda.Cells(PlanItensVenda.Rows.Count, "A").End(xlUp).Row + 1
            ultIDItem = ultIDItem + 1
            
            Dim idProd As String, descProd As String, pUnit As Double, qtd As Double, subt As Double
            idProd = .LstItens.List(i, 1)
            descProd = .LstItens.List(i, 2)
            pUnit = CDbl(Replace(.LstItens.List(i, 3), "R$", ""))
            qtd = CDbl(.LstItens.List(i, 4))
            subt = CDbl(Replace(.LstItens.List(i, 5), "R$", ""))
            
            PlanItensVenda.Cells(linItem, "A").Value = ultIDItem
            PlanItensVenda.Cells(linItem, "B").Value = idVenda
            PlanItensVenda.Cells(linItem, "C").Value = idProd
            PlanItensVenda.Cells(linItem, "D").Value = descProd
            PlanItensVenda.Cells(linItem, "E").Value = pUnit
            PlanItensVenda.Cells(linItem, "F").Value = qtd
            PlanItensVenda.Cells(linItem, "G").Value = subt
        Next i
        
        MsgBox "Venda Nº " & Format(idVenda, "0000") & " finalizada e gravada com sucesso!", vbInformation, "Venda Concluída"
        
        ' 5. Iniciar nova venda limpa
        Call IniciarNovaVenda
    End With
End Sub

