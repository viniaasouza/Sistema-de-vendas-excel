Attribute VB_Name = "ModDashboard"
Option Explicit
Option Compare Text

' =============================================================
' MOTOR DE DASHBOARD INTERATIVO, KPIS & GRÁFICOS (AULAS 22 A 27)
' =============================================================

Public gFiltroPeriodo As String ' "GERAL", "HOJE", "ESTE_MES", "MES_ANTERIOR"
Public gFiltroOperador As String ' "TODOS" ou nome/login do operador
Public gFiltroFormaPgto As String ' "TODAS" ou nome da forma

Private Function PlanVendas() As Worksheet
    Set PlanVendas = ThisWorkbook.Sheets("VENDAS")
End Function

Private Function PlanItens() As Worksheet
    Set PlanItens = ThisWorkbook.Sheets("ITENS_VENDA")
End Function

Private Function PlanInicio() As Worksheet
    Set PlanInicio = ThisWorkbook.Sheets("INÍCIO")
End Function

Private Function PlanUsuarios() As Worksheet
    Set PlanUsuarios = ThisWorkbook.Sheets("USUARIOS")
End Function

' === CONTROLADORES INTERATIVOS DE FILTRO (AULAS 26 & 27) ===

Public Sub FiltrarDashboardGeral()
    gFiltroPeriodo = "GERAL"
    gFiltroOperador = "TODOS"
    gFiltroFormaPgto = "TODAS"
    Call AtualizarDashboard
End Sub

Public Sub FiltrarDashboardHoje()
    gFiltroPeriodo = "HOJE"
    Call AtualizarDashboard
End Sub

Public Sub FiltrarDashboardEsteMes()
    gFiltroPeriodo = "ESTE_MES"
    Call AtualizarDashboard
End Sub

Public Sub FiltrarDashboardMesAnterior()
    gFiltroPeriodo = "MES_ANTERIOR"
    Call AtualizarDashboard
End Sub

Public Sub FiltrarDashboardOperador()
    Dim op As String
    op = InputBox("FILTRAR DASHBOARD POR OPERADOR:" & vbCrLf & vbCrLf & _
                  "• Digite 'TODOS' para ver todas as vendas" & vbCrLf & _
                  "• Ou informe o nome/login do operador (ex: Vitoria, Geovana, Admin):", _
                  "Filtro Interativo de Operador", IIf(Trim(gFiltroOperador) = "" Or gFiltroOperador = "TODOS", "TODOS", gFiltroOperador))
    If Trim(op) = "" Or UCase(Trim(op)) = "TODOS" Then
        gFiltroOperador = "TODOS"
    Else
        gFiltroOperador = Trim(op)
    End If
    Call AtualizarDashboard
End Sub

Public Sub LimparFiltrosDashboard()
    gFiltroPeriodo = "GERAL"
    gFiltroOperador = "TODOS"
    gFiltroFormaPgto = "TODAS"
    Call AtualizarDashboard
End Sub

Public Sub AtualizarDashboard()
    On Error GoTo TrataErro
    
    If Trim(gFiltroPeriodo) = "" Then gFiltroPeriodo = "GERAL"
    If Trim(gFiltroOperador) = "" Then gFiltroOperador = "TODOS"
    If Trim(gFiltroFormaPgto) = "" Then gFiltroFormaPgto = "TODAS"
    
    Dim wsIni As Worksheet
    Set wsIni = PlanInicio()
    
    Dim wsV As Worksheet, wsI As Worksheet
    Set wsV = PlanVendas()
    Set wsI = PlanItens()
    
    Dim ultV As Long, ultI As Long, r As Long
    ultV = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row
    ultI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    
    ' ---------------------------------------------------------
    ' 1. CÁLCULO DOS KPIS COM FILTROS INTERATIVOS (AULAS 22, 23, 26)
    ' ---------------------------------------------------------
    Dim totVendasConcluidas As Long
    Dim totFaturamentoLiq As Double
    Dim totValorReceber As Double
    Dim totItensVendidos As Long
    Dim ticketMedio As Double
    
    totVendasConcluidas = 0
    totFaturamentoLiq = 0
    totValorReceber = 0
    totItensVendidos = 0
    
    Dim vendasConcluidasIDs As Object
    Set vendasConcluidasIDs = CreateObject("Scripting.Dictionary")
    
    Dim clientesTotais As Object, clientesQtd As Object
    Set clientesTotais = CreateObject("Scripting.Dictionary")
    Set clientesQtd = CreateObject("Scripting.Dictionary")
    
    Dim faturasPorForma As Object
    Set faturasPorForma = CreateObject("Scripting.Dictionary")
    faturasPorForma("Dinheiro") = 0#
    faturasPorForma("PIX") = 0#
    faturasPorForma("Cartão de Crédito") = 0#
    faturasPorForma("Cartão de Débito") = 0#
    faturasPorForma("Boleto Bancário") = 0#
    
    Dim dtHoje As Date
    dtHoje = Date
    
    If ultV >= 2 Then
        For r = 2 To ultV
            Dim idV As Long, st As String, dtV As Date, opV As String, fp As String
            Dim vLiq As Double, vRec As Double, cliNome As String
            
            idV = Val(wsV.Cells(r, "A").Value)
            st = UCase(Trim(wsV.Cells(r, "I").Text))
            
            ' Verifica status da venda
            If st = "CONCLUÍDA" Or st = "CONCLUIDA" Or InStr(st, "CONCLU") > 0 Then
                ' 1. Filtro por Período
                Dim passaPeriodo As Boolean
                passaPeriodo = True
                
                If IsDate(wsV.Cells(r, "B").Value) Then
                    dtV = CDate(wsV.Cells(r, "B").Value)
                    Select Case UCase(Trim(gFiltroPeriodo))
                        Case "HOJE"
                            If Int(dtV) <> dtHoje Then passaPeriodo = False
                        Case "ESTE_MES"
                            If Month(dtV) <> Month(dtHoje) Or Year(dtV) <> Year(dtHoje) Then passaPeriodo = False
                        Case "MES_ANTERIOR"
                            Dim dtAnt As Date
                            dtAnt = DateAdd("m", -1, dtHoje)
                            If Month(dtV) <> Month(dtAnt) Or Year(dtV) <> Year(dtAnt) Then passaPeriodo = False
                    End Select
                End If
                
                ' 2. Filtro por Operador
                Dim passaOperador As Boolean
                passaOperador = True
                opV = Trim(wsV.Cells(r, "M").Text)
                If UCase(Trim(gFiltroOperador)) <> "TODOS" And Trim(gFiltroOperador) <> "" Then
                    If InStr(1, opV, gFiltroOperador, vbTextCompare) = 0 Then passaOperador = False
                End If
                
                ' 3. Filtro por Forma de Pagamento
                Dim passaForma As Boolean
                passaForma = True
                fp = Trim(wsV.Cells(r, "H").Text)
                If UCase(Trim(gFiltroFormaPgto)) <> "TODAS" And Trim(gFiltroFormaPgto) <> "" Then
                    If InStr(1, fp, gFiltroFormaPgto, vbTextCompare) = 0 Then passaForma = False
                End If
                
                If passaPeriodo And passaOperador And passaForma Then
                    vLiq = Val(Replace(wsV.Cells(r, "G").Value, ",", "."))
                    vRec = Val(Replace(wsV.Cells(r, "L").Value, ",", "."))
                    cliNome = Trim(wsV.Cells(r, "D").Text)
                    
                    totVendasConcluidas = totVendasConcluidas + 1
                    totFaturamentoLiq = totFaturamentoLiq + vLiq
                    totValorReceber = totValorReceber + vRec
                    vendasConcluidasIDs(idV) = True
                    
                    ' Agrupamento por Forma de Pagamento
                    If InStr(1, fp, "crédito", vbTextCompare) > 0 Or InStr(1, fp, "credito", vbTextCompare) > 0 Then
                        faturasPorForma("Cartão de Crédito") = faturasPorForma("Cartão de Crédito") + vLiq
                    ElseIf InStr(1, fp, "débito", vbTextCompare) > 0 Or InStr(1, fp, "debito", vbTextCompare) > 0 Then
                        faturasPorForma("Cartão de Débito") = faturasPorForma("Cartão de Débito") + vLiq
                    ElseIf InStr(1, fp, "pix", vbTextCompare) > 0 Then
                        faturasPorForma("PIX") = faturasPorForma("PIX") + vLiq
                    ElseIf InStr(1, fp, "dinheiro", vbTextCompare) > 0 Then
                        faturasPorForma("Dinheiro") = faturasPorForma("Dinheiro") + vLiq
                    ElseIf InStr(1, fp, "boleto", vbTextCompare) > 0 Then
                        faturasPorForma("Boleto Bancário") = faturasPorForma("Boleto Bancário") + vLiq
                    Else
                        faturasPorForma(fp) = faturasPorForma(fp) + vLiq
                    End If
                    
                    If cliNome <> "" Then
                        clientesTotais(cliNome) = clientesTotais(cliNome) + vLiq
                        clientesQtd(cliNome) = clientesQtd(cliNome) + 1
                    End If
                End If
            End If
        Next r
    End If
    
    ' Contagem de Itens e Agrupamento de Produtos e Serviços (Aula 25)
    Dim produtosTotais As Object, produtosQtd As Object
    Set produtosTotais = CreateObject("Scripting.Dictionary")
    Set produtosQtd = CreateObject("Scripting.Dictionary")
    
    If ultI >= 2 Then
        For r = 2 To ultI
            Dim idVendaItem As Long
            idVendaItem = Val(wsI.Cells(r, "B").Value)
            
            If vendasConcluidasIDs.Exists(idVendaItem) Then
                Dim qtdItem As Long, subTotItem As Double, prodNome As String
                qtdItem = Val(wsI.Cells(r, "F").Value)
                subTotItem = Val(Replace(wsI.Cells(r, "G").Value, ",", "."))
                prodNome = Trim(wsI.Cells(r, "D").Text)
                
                totItensVendidos = totItensVendidos + qtdItem
                
                If prodNome <> "" Then
                    produtosQtd(prodNome) = produtosQtd(prodNome) + qtdItem
                    produtosTotais(prodNome) = produtosTotais(prodNome) + subTotItem
                End If
            End If
        Next r
    End If
    
    If totVendasConcluidas > 0 Then
        ticketMedio = totFaturamentoLiq / totVendasConcluidas
    Else
        ticketMedio = 0
    End If
    
    ' ---------------------------------------------------------
    ' 2. BANNER DE STATUS DO FILTRO ATIVO (AULA 26)
    ' ---------------------------------------------------------
    Dim descFiltro As String
    Select Case UCase(Trim(gFiltroPeriodo))
        Case "HOJE": descFiltro = "Hoje (" & Format(Date, "dd/mm/yyyy") & ")"
        Case "ESTE_MES": descFiltro = "Este Mês (" & Format(Date, "mmmm/yyyy") & ")"
        Case "MES_ANTERIOR": descFiltro = "Mês Anterior"
        Case Else: descFiltro = "Todo o Período"
    End Select
    
    wsIni.Range("B10").Value = "?? FILTRO ATIVO: Período: [ " & descFiltro & " ]  |  Operador: [ " & gFiltroOperador & " ]  |  Pedidos no Filtro: [ " & totVendasConcluidas & " pedidos ]"
    
    ' ---------------------------------------------------------
    ' 3. ATUALIZAÇÃO DOS 4 CARDS DE KPIS
    ' ---------------------------------------------------------
    wsIni.Range("B12").Value = totFaturamentoLiq
    wsIni.Range("B12").NumberFormat = "R$ #,##0.00"
    
    wsIni.Range("D12").Value = totVendasConcluidas & " ped. (" & totItensVendidos & " itens)"
    
    wsIni.Range("G12").Value = ticketMedio
    wsIni.Range("G12").NumberFormat = "R$ #,##0.00"
    
    wsIni.Range("I12").Value = totValorReceber
    wsIni.Range("I12").NumberFormat = "R$ #,##0.00"
    
    ' ---------------------------------------------------------
    ' 4. RANKING TOP 5 CLIENTES (AULA 24)
    ' ---------------------------------------------------------
    Dim k As Long
    For k = 17 To 21
        wsIni.Range("B" & k).Value = (k - 16) & "º"
        wsIni.Range("C" & k).Value = "-"
        wsIni.Range("D" & k).Value = 0
        wsIni.Range("E" & k).Value = 0
        wsIni.Range("E" & k).NumberFormat = "R$ #,##0.00"
    Next k
    
    Dim numCli As Long, idx As Long
    numCli = clientesTotais.Count
    
    If numCli > 0 Then
        Dim cliArr() As String, valCliArr() As Double, qtdCliArr() As Long
        ReDim cliArr(1 To numCli)
        ReDim valCliArr(1 To numCli)
        ReDim qtdCliArr(1 To numCli)
        
        idx = 1
        Dim kCli As Variant
        For Each kCli In clientesTotais.Keys
            cliArr(idx) = CStr(kCli)
            valCliArr(idx) = CDbl(clientesTotais(kCli))
            qtdCliArr(idx) = CLng(clientesQtd(kCli))
            idx = idx + 1
        Next kCli
        
        Dim a As Long, b As Long, tmpC As String, tmpV As Double, tmpQ As Long
        For a = 1 To numCli - 1
            For b = a + 1 To numCli
                If valCliArr(b) > valCliArr(a) Then
                    tmpV = valCliArr(a): valCliArr(a) = valCliArr(b): valCliArr(b) = tmpV
                    tmpC = cliArr(a): cliArr(a) = cliArr(b): cliArr(b) = tmpC
                    tmpQ = qtdCliArr(a): qtdCliArr(a) = qtdCliArr(b): qtdCliArr(b) = tmpQ
                End If
            Next b
        Next a
        
        Dim limiteCli As Long
        limiteCli = IIf(numCli > 5, 5, numCli)
        For a = 1 To limiteCli
            Dim linDestCli As Long
            linDestCli = 16 + a
            wsIni.Range("C" & linDestCli).Value = cliArr(a)
            wsIni.Range("D" & linDestCli).Value = qtdCliArr(a)
            wsIni.Range("E" & linDestCli).Value = valCliArr(a)
            wsIni.Range("E" & linDestCli).NumberFormat = "R$ #,##0.00"
        Next a
    End If
    
    ' ---------------------------------------------------------
    ' 5. RANKING TOP 5 PRODUTOS E SERVIÇOS MAIS VENDIDOS (AULAS 25 & 27)
    ' ---------------------------------------------------------
    wsIni.Range("G15").Value = "?? TOP 5 PRODUTOS & SERVIÇOS MAIS VENDIDOS"
    wsIni.Range("H16").Value = "Produto / Serviço"
    
    For k = 17 To 21
        wsIni.Range("G" & k).Value = (k - 16) & "º"
        wsIni.Range("H" & k).Value = "-"
        wsIni.Range("I" & k).Value = 0
        wsIni.Range("J" & k).Value = 0
        wsIni.Range("J" & k).NumberFormat = "R$ #,##0.00"
        
        ' Células auxiliares para alimentar o gráfico
        wsIni.Range("L" & k).Value = "-"
        wsIni.Range("M" & k).Value = 0
    Next k
    
    Dim numProd As Long
    numProd = produtosTotais.Count
    
    If numProd > 0 Then
        Dim prodArr() As String, valProdArr() As Double, qtdProdArr() As Long
        ReDim prodArr(1 To numProd)
        ReDim valProdArr(1 To numProd)
        ReDim qtdProdArr(1 To numProd)
        
        idx = 1
        Dim kProd As Variant
        For Each kProd In produtosTotais.Keys
            prodArr(idx) = CStr(kProd)
            valProdArr(idx) = CDbl(produtosTotais(kProd))
            qtdProdArr(idx) = CLng(produtosQtd(kProd))
            idx = idx + 1
        Next kProd
        
        For a = 1 To numProd - 1
            For b = a + 1 To numProd
                If valProdArr(b) > valProdArr(a) Then
                    tmpV = valProdArr(a): valProdArr(a) = valProdArr(b): valProdArr(b) = tmpV
                    tmpC = prodArr(a): prodArr(a) = prodArr(b): prodArr(b) = tmpC
                    tmpQ = qtdProdArr(a): qtdProdArr(a) = qtdProdArr(b): qtdProdArr(b) = tmpQ
                End If
            Next b
        Next a
        
        Dim limiteProd As Long
        limiteProd = IIf(numProd > 5, 5, numProd)
        For a = 1 To limiteProd
            Dim linDestProd As Long
            linDestProd = 16 + a
            wsIni.Range("H" & linDestProd).Value = prodArr(a)
            wsIni.Range("I" & linDestProd).Value = qtdProdArr(a)
            wsIni.Range("J" & linDestProd).Value = valProdArr(a)
            wsIni.Range("J" & linDestProd).NumberFormat = "R$ #,##0.00"
            
            ' Auxiliar gráfico
            wsIni.Range("L" & linDestProd).Value = prodArr(a)
            wsIni.Range("M" & linDestProd).Value = valProdArr(a)
        Next a
    End If
    
    ' ---------------------------------------------------------
    ' 6. FORMAS DE PAGAMENTO (LINHAS 24 A 28)
    ' ---------------------------------------------------------
    Dim fpLista As Variant
    fpLista = Array("Dinheiro", "PIX", "Cartão de Crédito", "Cartão de Débito", "Boleto Bancário")
    
    Dim linFP As Long
    For linFP = 0 To UBound(fpLista)
        Dim nomeFP As String
        nomeFP = fpLista(linFP)
        Dim cRow As Long
        cRow = 24 + linFP
        wsIni.Range("B" & cRow).Value = nomeFP
        If faturasPorForma.Exists(nomeFP) Then
            wsIni.Range("D" & cRow).Value = faturasPorForma(nomeFP)
            wsIni.Range("C" & cRow).Value = faturasPorForma(nomeFP)
        Else
            wsIni.Range("D" & cRow).Value = 0
            wsIni.Range("C" & cRow).Value = 0
        End If
        wsIni.Range("D" & cRow).NumberFormat = "R$ #,##0.00"
        wsIni.Range("C" & cRow).NumberFormat = "R$ #,##0.00"
    Next linFP
    
    ' Rodapé com data/hora e status
    wsIni.Range("B30").Value = "Última atualização: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & " | ?? Segurança Ativa: Aba USUARIOS protegida contra visualização"
    
    ' Atualiza gráficos interativos
    Call AtualizarGraficosInterativos
    Exit Sub
    
TrataErro:
    ' Falha tratada silenciosamente
End Sub

Public Sub AtualizarGraficosInterativos()
    On Error Resume Next
    Dim wsIni As Worksheet
    Set wsIni = PlanInicio()
    If wsIni Is Nothing Then Exit Sub
    
    Dim chObj As ChartObject
    Dim chFormas As ChartObject, chTopProd As ChartObject
    
    For Each chObj In wsIni.ChartObjects
        If chObj.Name = "GraficoFormasPgto" Then Set chFormas = chObj
        If chObj.Name = "GraficoTopProdutos" Then Set chTopProd = chObj
    Next chObj
    
    ' 1. Gráfico de Rosca: Formas de Pagamento
    If chFormas Is Nothing Then
        Set chFormas = wsIni.ChartObjects.Add(wsIni.Range("B33").Left, wsIni.Range("B33").Top, wsIni.Range("B33:E33").Width, 230)
        chFormas.Name = "GraficoFormasPgto"
        chFormas.Chart.ChartType = xlDoughnut
        chFormas.Chart.SetSourceData wsIni.Range("B24:C28")
        chFormas.Chart.HasTitle = True
        chFormas.Chart.ChartTitle.Text = "Faturamento por Forma de Pagamento"
        chFormas.Chart.ChartTitle.Font.Size = 10
        chFormas.Chart.ChartTitle.Font.Bold = True
        chFormas.Chart.ApplyDataLabels xlDataLabelsShowPercent
    Else
        chFormas.Chart.SetSourceData wsIni.Range("B24:C28")
    End If
    
    ' 2. Gráfico de Barras: TOP 5 Produtos e Serviços
    If chTopProd Is Nothing Then
        Set chTopProd = wsIni.ChartObjects.Add(wsIni.Range("G33").Left, wsIni.Range("G33").Top, wsIni.Range("G33:K33").Width, 230)
        chTopProd.Name = "GraficoTopProdutos"
        chTopProd.Chart.ChartType = xlBarClustered
        chTopProd.Chart.SetSourceData wsIni.Range("L17:M21")
        chTopProd.Chart.HasTitle = True
        chTopProd.Chart.ChartTitle.Text = "TOP 5 Produtos e Serviços (Faturamento)"
        chTopProd.Chart.ChartTitle.Font.Size = 10
        chTopProd.Chart.ChartTitle.Font.Bold = True
        chTopProd.Chart.HasLegend = False
    Else
        chTopProd.Chart.SetSourceData wsIni.Range("L17:M21")
    End If
    On Error GoTo 0
End Sub

