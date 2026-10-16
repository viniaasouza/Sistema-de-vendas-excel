Attribute VB_Name = "ModProdutosServicos"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE GESTÃO DE PRODUTOS, SERVIÇOS & INVENTÁRIO (FASE 8)
' =============================================================
Public vIDProduto As Variant

Public Function PlanProdutosServicos() As Worksheet
    On Error Resume Next
    Set PlanProdutosServicos = ThisWorkbook.Sheets("PRODUTOS_SERVICOS")
    If PlanProdutosServicos Is Nothing Then
        Set PlanProdutosServicos = ThisWorkbook.Sheets("PRODUTOS")
    End If
    On Error GoTo 0
End Function

Public Function PlanInventario() As Worksheet
    On Error Resume Next
    Set PlanInventario = ThisWorkbook.Sheets("INVENTARIO")
    On Error GoTo 0
End Function

Public Sub SalvarProdutosServicos()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Salvar ou Alterar Catálogo de Produtos e Serviços") Then Exit Sub
    End If
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    If wsP Is Nothing Then Exit Sub
    
    With FormProdutos
        .TxtDescricao.BorderColor = &H80000012
        .TxtPreco.BorderColor = &H80000012
        
        If Trim(.TxtDescricao.Text) = "" Then
            .TxtDescricao.BorderColor = &HFF&
            MsgBox "Digite o campo Descrição do Produto / Serviço!", vbExclamation, "Cadastro"
            .TxtDescricao.SetFocus
            Exit Sub
        End If
        
        Dim precoFinalStr As String, precoCustoStr As String, margemStr As String
        precoFinalStr = Replace(Replace(.TxtPreco.Text, "R$", ""), " ", "")
        precoCustoStr = Replace(Replace(.TxtPrecoCusto.Text, "R$", ""), " ", "")
        margemStr = Replace(Replace(.TxtMargem.Text, "%", ""), " ", "")
        
        Dim valPrecoFinal As Double, valPrecoCusto As Double, valMargem As Double
        valPrecoFinal = IIf(IsNumeric(precoFinalStr), CDbl(precoFinalStr), 0#)
        valPrecoCusto = IIf(IsNumeric(precoCustoStr), CDbl(precoCustoStr), 0#)
        valMargem = IIf(IsNumeric(margemStr), CDbl(margemStr), 0#)
        
        If valPrecoFinal <= 0 And valPrecoCusto > 0 Then
            valPrecoFinal = Round(valPrecoCusto * (1 + (valMargem / 100)), 2)
        End If
        
        If valPrecoFinal <= 0 Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Informe o Preço Final de Venda!", vbExclamation, "Preço Inválido"
            .TxtPreco.SetFocus
            Exit Sub
        End If
        
        Dim tipoItem As String
        tipoItem = Trim(.CboTipo.Text)
        If tipoItem = "" Then tipoItem = "Produto"
        
        Dim estAtual As Long, estMin As Long
        If UCase(tipoItem) = "SERVIÇO" Or UCase(tipoItem) = "SERVICO" Then
            estAtual = 0
            estMin = 0
        Else
            estAtual = Val(.TxtEstoque.Text)
            estMin = Val(.TxtEstoqueMin.Text)
            If estMin = 0 Then estMin = 5
        End If
        
        Dim cel As Range
        If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
            Set cel = wsP.Range("C:C").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                .TxtDescricao.BorderColor = &HFF&
                MsgBox "Já existe um item cadastrado com esta Descrição!", vbExclamation, "Item Duplicado"
                .TxtDescricao.SetFocus
                Exit Sub
            End If
            
            Dim lin As Long
            lin = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row + 1
            Dim novoID As Long
            If lin = 2 Then novoID = 1 Else novoID = Val(wsP.Cells(lin - 1, "A").Value) + 1
            
            wsP.Cells(lin, "A").Value = novoID
            wsP.Cells(lin, "B").Value = tipoItem
            wsP.Cells(lin, "C").Value = Trim(.TxtDescricao.Text)
            wsP.Cells(lin, "D").Value = valPrecoCusto
            wsP.Cells(lin, "E").Value = valMargem
            wsP.Cells(lin, "F").Value = valPrecoFinal
            wsP.Cells(lin, "G").Value = estAtual
            wsP.Cells(lin, "H").Value = estMin
            
            MsgBox "Item cadastrado com sucesso!", vbInformation, "Cadastro Concluído"
        Else
            Set cel = wsP.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
            If cel Is Nothing Then
                MsgBox "Item não localizado para alteração!", vbCritical, "Erro"
                Exit Sub
            End If
            
            cel.Offset(0, 1).Value = tipoItem
            cel.Offset(0, 2).Value = Trim(.TxtDescricao.Text)
            cel.Offset(0, 3).Value = valPrecoCusto
            cel.Offset(0, 4).Value = valMargem
            cel.Offset(0, 5).Value = valPrecoFinal
            cel.Offset(0, 6).Value = estAtual
            cel.Offset(0, 7).Value = estMin
            
            MsgBox "Item alterado com sucesso!", vbInformation, "Atualização Concluída"
        End If
        
        Call LimparCampos
        Call BuscarProdutosServicos
        Call AtualizarInventario
    End With
End Sub

Public Sub ExcluirProdutoServico()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Excluir Item do Catálogo") Then Exit Sub
    End If
    
    If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
        MsgBox "Selecione um item na lista para excluir!", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    Dim cel As Range
    Set cel = wsP.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cel Is Nothing Then
        MsgBox "Item não encontrado!", vbCritical, "Erro"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir: " & cel.Offset(0, 2).Text & "?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbYes Then
        cel.EntireRow.Delete
        MsgBox "Item excluído com sucesso!", vbInformation, "Exclusão"
        Call LimparCampos
        Call BuscarProdutosServicos
        Call AtualizarInventario
    End If
End Sub

Public Sub LimparCampos()
    With FormProdutos
        .CboTipo.Text = "Produto"
        .TxtDescricao.Text = ""
        .TxtPrecoCusto.Text = ""
        .TxtMargem.Text = ""
        .TxtPreco.Text = ""
        .TxtEstoque.Text = ""
        .TxtEstoqueMin.Text = ""
        .TxtDescricao.BorderColor = &H80000012
        .TxtPreco.BorderColor = &H80000012
        .BtSalvar.Caption = "SALVAR"
        On Error Resume Next
        .BtExcluir.Enabled = False
        On Error GoTo 0
        vIDProduto = Empty
    End With
End Sub

Public Sub BuscarProdutosServicos()
    Dim lin As Long, i As Long, ultLinha As Long
    i = 0
    FormProdutos.LstDados.Clear
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    If wsP Is Nothing Then Exit Sub
    
    ultLinha = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    If ultLinha < 2 Then Exit Sub
    
    For lin = 2 To ultLinha
        If wsP.Cells(lin, "C").Text Like "*" & ModFuncoes.AcSQL(FormProdutos.TxtPesquisa.Text) & "*" Then
            FormProdutos.LstDados.AddItem
            FormProdutos.LstDados.List(i, 0) = wsP.Cells(lin, "A").Text
            FormProdutos.LstDados.List(i, 1) = wsP.Cells(lin, "B").Text
            FormProdutos.LstDados.List(i, 2) = wsP.Cells(lin, "C").Text
            FormProdutos.LstDados.List(i, 3) = Format(Val(Replace(wsP.Cells(lin, "D").Value, ",", ".")), "R$ #,##0.00")
            FormProdutos.LstDados.List(i, 4) = Format(Val(Replace(wsP.Cells(lin, "E").Value, ",", ".")), "0.0") & "%"
            FormProdutos.LstDados.List(i, 5) = Format(Val(Replace(wsP.Cells(lin, "F").Value, ",", ".")), "R$ #,##0.00")
            FormProdutos.LstDados.List(i, 6) = wsP.Cells(lin, "G").Text
            i = i + 1
        End If
    Next lin
End Sub

Public Sub AtualizarInventario()
    On Error Resume Next
    Dim wsP As Worksheet, wsInv As Worksheet
    Set wsP = PlanProdutosServicos()
    Set wsInv = PlanInventario()
    If wsP Is Nothing Or wsInv Is Nothing Then Exit Sub
    wsInv.Unprotect "VendasSeguras2026"
    wsP.Unprotect "VendasSeguras2026"
    
    Dim ultP As Long, r As Long
    ultP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    
    Dim ultInv As Long
    ultInv = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row
    If ultInv >= 12 Then
        wsInv.Range("B12:M" & ultInv + 5).Clear
    End If
    
    Dim totItensEstoque As Long, totCustoEstoque As Double, totVendaEstoque As Double
    Dim totLucroEstoque As Double, totEstoqueBaixo As Long
    totItensEstoque = 0
    totCustoEstoque = 0#
    totVendaEstoque = 0#
    totLucroEstoque = 0#
    totEstoqueBaixo = 0
    
    Dim linDest As Long
    linDest = 12
    
    If ultP >= 2 Then
        For r = 2 To ultP
            Dim idP As Long, tipoP As String, descP As String
            Dim vCusto As Double, vMargem As Double, vVenda As Double
            Dim qEstoque As Long, qMin As Long
            
            idP = Val(wsP.Cells(r, "A").Value)
            tipoP = Trim(wsP.Cells(r, "B").Text)
            descP = Trim(wsP.Cells(r, "C").Text)
            
            If idP = 0 And descP = "" Then GoTo ProximoItem
            
            vCusto = Val(Replace(wsP.Cells(r, "D").Value, ",", "."))
            vMargem = Val(Replace(wsP.Cells(r, "E").Value, ",", "."))
            vVenda = Val(Replace(wsP.Cells(r, "F").Value, ",", "."))
            qEstoque = Val(wsP.Cells(r, "G").Value)
            qMin = Val(wsP.Cells(r, "H").Value)
            If qMin = 0 Then qMin = 5
            
            If tipoP = "" Then tipoP = "Produto"
            
            Dim stEstoque As String
            If UCase(tipoP) = "SERVIÇO" Or UCase(tipoP) = "SERVICO" Then
                stEstoque = "??? SERVIÇO"
            Else
                totItensEstoque = totItensEstoque + qEstoque
                totCustoEstoque = totCustoEstoque + (qEstoque * vCusto)
                totVendaEstoque = totVendaEstoque + (qEstoque * vVenda)
                
                If qEstoque <= 0 Then
                    stEstoque = "? ZERADO"
                    totEstoqueBaixo = totEstoqueBaixo + 1
                ElseIf qEstoque <= qMin Then
                    stEstoque = "?? ESTOQUE BAIXO"
                    totEstoqueBaixo = totEstoqueBaixo + 1
                Else
                    stEstoque = "? NORMAL"
                End If
            End If
            
            Dim totCustoItem As Double, totVendaItem As Double, totLucroItem As Double
            totCustoItem = qEstoque * vCusto
            totVendaItem = qEstoque * vVenda
            totLucroItem = totVendaItem - totCustoItem
            
            wsInv.Cells(linDest, "B").Value = idP
            wsInv.Cells(linDest, "C").Value = tipoP
            wsInv.Cells(linDest, "D").Value = descP
            wsInv.Cells(linDest, "E").Value = vCusto
            wsInv.Cells(linDest, "E").NumberFormat = "R$ #,##0.00"
            wsInv.Cells(linDest, "F").Value = vMargem / 100
            wsInv.Cells(linDest, "F").NumberFormat = "0.0%"
            wsInv.Cells(linDest, "G").Value = vVenda
            wsInv.Cells(linDest, "G").NumberFormat = "R$ #,##0.00"
            
            If UCase(tipoP) = "SERVIÇO" Or UCase(tipoP) = "SERVICO" Then
                wsInv.Cells(linDest, "H").Value = "-"
                wsInv.Cells(linDest, "I").Value = "-"
                wsInv.Cells(linDest, "K").Value = "-"
                wsInv.Cells(linDest, "L").Value = "-"
                wsInv.Cells(linDest, "M").Value = "-"
            Else
                wsInv.Cells(linDest, "H").Value = qEstoque
                wsInv.Cells(linDest, "I").Value = qMin
                wsInv.Cells(linDest, "K").Value = totCustoItem
                wsInv.Cells(linDest, "K").NumberFormat = "R$ #,##0.00"
                wsInv.Cells(linDest, "L").Value = totVendaItem
                wsInv.Cells(linDest, "L").NumberFormat = "R$ #,##0.00"
                wsInv.Cells(linDest, "M").Value = totLucroItem
                wsInv.Cells(linDest, "M").NumberFormat = "R$ #,##0.00"
            End If
            
            wsInv.Cells(linDest, "J").Value = stEstoque
            
            If (linDest Mod 2) = 0 Then
                wsInv.Range("B" & linDest & ":M" & linDest).Interior.Color = RGB(247, 249, 250)
            Else
                wsInv.Range("B" & linDest & ":M" & linDest).Interior.Color = RGB(255, 255, 255)
            End If
            
            linDest = linDest + 1
ProximoItem:
        Next r
    End If
    
    totLucroEstoque = totVendaEstoque - totCustoEstoque
    
    wsInv.Range("B6").Value = totItensEstoque & " un."
    wsInv.Range("D6").Value = totCustoEstoque
    wsInv.Range("D6").NumberFormat = "R$ #,##0.00"
    wsInv.Range("F6").Value = totVendaEstoque
    wsInv.Range("F6").NumberFormat = "R$ #,##0.00"
    wsInv.Range("H6").Value = totLucroEstoque
    wsInv.Range("H6").NumberFormat = "R$ #,##0.00"
    wsInv.Range("J6").Value = totEstoqueBaixo & " alertas"
    Call AplicarProtecaoInventario
    
    On Error GoTo 0
End Sub

Public Sub SalvarProdutos()
    Call SalvarProdutosServicos
End Sub

Public Sub BuscarProdutos()
    Call BuscarProdutosServicos
End Sub

Public Sub ExcluirProduto()
    Call ExcluirProdutoServico
End Sub


Public Sub NovoItemInventario()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        MsgBox "Acesso Restrito: Recepcionistas possuem permissão apenas para consulta do inventário." & vbCrLf & vbCrLf & _
               "Qualquer alteração de estoque, custo ou cadastro exige autorização de Administrador.", vbExclamation, "Restrição de Inventário"
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Alterar ou Cadastrar Item no Inventário") Then
            Exit Sub
        End If
    End If
    FormProdutos.Show
End Sub

Public Sub AplicarProtecaoInventario()
    On Error Resume Next
    Dim wsInv As Worksheet, wsProd As Worksheet
    Set wsInv = PlanInventario()
    Set wsProd = PlanProdutosServicos()
    
    If Not wsInv Is Nothing Then
        wsInv.Unprotect "VendasSeguras2026"
        If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
            wsInv.Protect Password:="VendasSeguras2026", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
        Else
            wsInv.Protect Password:="VendasSeguras2026", UserInterfaceOnly:=True, DrawingObjects:=False, Contents:=True, Scenarios:=True
        End If
    End If
    
    If Not wsProd Is Nothing Then
        wsProd.Unprotect "VendasSeguras2026"
        If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
            wsProd.Protect Password:="VendasSeguras2026", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
    End If
    On Error GoTo 0
End Sub



