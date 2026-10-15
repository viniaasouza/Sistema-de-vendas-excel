Attribute VB_Name = "ModDashboard"
Option Explicit
Option Compare Text

' =============================================================
' MOTOR DE INDICADORES, KPIS E RANKINGS DO DASHBOARD (AULAS 22 A 25)
' =============================================================

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

Public Sub AtualizarDashboard()
    On Error GoTo TrataErro
    
    Dim wsIni As Worksheet
    Set wsIni = PlanInicio()
    
    Dim wsV As Worksheet, wsI As Worksheet
    Set wsV = PlanVendas()
    Set wsI = PlanItens()
    
    Dim ultV As Long, ultI As Long, r As Long
    ultV = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row
    ultI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    
    ' ---------------------------------------------------------
    ' 1. CÁLCULO DOS KPIS GLOBAIS (AULAS 22 E 23)
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
    
    Dim faturasPorForma As Object
    Set faturasPorForma = CreateObject("Scripting.Dictionary")
    faturasPorForma.CompareMode = vbTextCompare
    
    Dim vendasConcluidasIDs As Object
    Set vendasConcluidasIDs = CreateObject("Scripting.Dictionary")
    
    Dim clientesTotais As Object
    Set clientesTotais = CreateObject("Scripting.Dictionary")
    Dim clientesQtd As Object
    Set clientesQtd = CreateObject("Scripting.Dictionary")
    
    If ultV >= 2 Then
        For r = 2 To ultV
            Dim st As String, idV As Long
            idV = Val(wsV.Cells(r, "A").Value)
            st = Trim(wsV.Cells(r, "I").Text)
            
            If LCase(st) = "concluída" Or LCase(st) = "concluida" Then
                totVendasConcluidas = totVendasConcluidas + 1
                vendasConcluidasIDs(idV) = True
                
                Dim vLiq As Double, vRec As Double, fp As String, cliNome As String
                vLiq = Val(Replace(wsV.Cells(r, "G").Value, ",", "."))
                vRec = Val(Replace(wsV.Cells(r, "L").Value, ",", "."))
                If vRec = 0 Then vRec = vLiq
                fp = Trim(wsV.Cells(r, "H").Text)
                cliNome = Trim(wsV.Cells(r, "D").Text)
                
                totFaturamentoLiq = totFaturamentoLiq + vLiq
                totValorReceber = totValorReceber + vRec
                
                If fp <> "" Then
                    ' Normalização de acentos
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
                End If
                
                If cliNome <> "" Then
                    clientesTotais(cliNome) = clientesTotais(cliNome) + vLiq
                    clientesQtd(cliNome) = clientesQtd(cliNome) + 1
                End If
            End If
        Next r
    End If
    
    ' Contagem de Itens e Agrupamento de Produtos (Aula 25)
    Dim produtosTotais As Object, produtosQtd As Object
    Set produtosTotais = CreateObject("Scripting.Dictionary")
    Set produtosQtd = CreateObject("Scripting.Dictionary")
    
    If ultI >= 2 Then
        For r = 2 To ultI
            Dim idVendaItem As Long
            idVendaItem = Val(wsI.Cells(r, "B").Value)
            
            ' Só contabiliza itens de vendas concluídas
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
    ' 2. ATUALIZAÇÃO DOS 4 CARDS DE KPIS
    ' ---------------------------------------------------------
    wsIni.Range("B12").Value = totFaturamentoLiq
    wsIni.Range("B12").NumberFormat = "R$ #,##0.00"
    
    wsIni.Range("D12").Value = totVendasConcluidas & " ped. (" & totItensVendidos & " itens)"
    
    wsIni.Range("G12").Value = ticketMedio
    wsIni.Range("G12").NumberFormat = "R$ #,##0.00"
    
    wsIni.Range("I12").Value = totValorReceber
    wsIni.Range("I12").NumberFormat = "R$ #,##0.00"
    
    ' ---------------------------------------------------------
    ' 3. RANKING TOP 5 CLIENTES (AULA 24)
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
    ' 4. RANKING TOP 5 PRODUTOS MAIS VENDIDOS (AULA 25)
    ' ---------------------------------------------------------
    For k = 17 To 21
        wsIni.Range("G" & k).Value = (k - 16) & "º"
        wsIni.Range("H" & k).Value = "-"
        wsIni.Range("I" & k).Value = 0
        wsIni.Range("J" & k).Value = 0
        wsIni.Range("J" & k).NumberFormat = "R$ #,##0.00"
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
        Next a
    End If
    
    ' ---------------------------------------------------------
    ' 5. FORMAS DE PAGAMENTO (LINHAS 24 A 28)
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
        Else
            wsIni.Range("D" & cRow).Value = 0
        End If
        wsIni.Range("D" & cRow).NumberFormat = "R$ #,##0.00"
    Next linFP
    
    ' Rodapé com data/hora e status de segurança
    wsIni.Range("B30").Value = "Última atualização: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & " | ?? Segurança Ativa: Aba USUARIOS protegida por perfil"
    Exit Sub
    
TrataErro:
    ' Falha tratada silenciosamente
End Sub

