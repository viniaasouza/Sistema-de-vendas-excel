Attribute VB_Name = "ModRelatorios"
Option Explicit
Option Compare Text

Public bBloqueiaFiltro As Boolean

Private Function PlanVendas() As Worksheet
    Set PlanVendas = ThisWorkbook.Sheets("VENDAS")
End Function

Private Function PlanItens() As Worksheet
    Set PlanItens = ThisWorkbook.Sheets("ITENS_VENDA")
End Function

Private Function PlanRelVendas() As Worksheet
    Set PlanRelVendas = ThisWorkbook.Sheets("R_Vendas")
End Function

Private Function PlanConfig() As Worksheet
    Set PlanConfig = ThisWorkbook.Sheets("CONFIG")
End Function

Private Function LimparMoeda(ByVal valStr As String) As Double
    Dim s As String
    s = Trim(Replace(valStr, "R$", ""))
    s = Replace(s, " ", "")
    If Application.International(xlDecimalSeparator) = "," Then
        s = Replace(s, ".", "")
    Else
        s = Replace(s, ",", "")
    End If
    If IsNumeric(s) Then
        LimparMoeda = CDbl(s)
    Else
        LimparMoeda = 0
    End If
End Function

Public Function ObterPastaLocalRelatorios() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim caminhoBase As String
    caminhoBase = ThisWorkbook.Path
    
    ' Tratamento para caminhos do OneDrive (URL http / https)
    If LCase(Left(caminhoBase, 4)) = "http" Then
        Dim userProf As String
        userProf = Environ("USERPROFILE")
        
        If fso.FolderExists(userProf & "\OneDrive\Área de Trabalho\SISTEMA DE VENDAS") Then
            caminhoBase = userProf & "\OneDrive\Área de Trabalho\SISTEMA DE VENDAS"
        ElseIf fso.FolderExists(userProf & "\OneDrive\Desktop\SISTEMA DE VENDAS") Then
            caminhoBase = userProf & "\OneDrive\Desktop\SISTEMA DE VENDAS"
        ElseIf fso.FolderExists(userProf & "\Área de Trabalho\SISTEMA DE VENDAS") Then
            caminhoBase = userProf & "\Área de Trabalho\SISTEMA DE VENDAS"
        ElseIf fso.FolderExists(userProf & "\Desktop\SISTEMA DE VENDAS") Then
            caminhoBase = userProf & "\Desktop\SISTEMA DE VENDAS"
        Else
            caminhoBase = userProf & "\Desktop"
        End If
    End If
    
    Dim pastaDestino As String
    pastaDestino = caminhoBase & "\Relatorios"
    
    If Not fso.FolderExists(pastaDestino) Then
        On Error Resume Next
        fso.CreateFolder pastaDestino
        On Error GoTo 0
    End If
    
    If fso.FolderExists(pastaDestino) Then
        ObterPastaLocalRelatorios = pastaDestino
    Else
        ObterPastaLocalRelatorios = caminhoBase
    End If
End Function

Sub AbrirRelatorios()
    FormRelatorios.Show
End Sub

Sub InicializarFormRelatorios()
    bBloqueiaFiltro = True
    
    With FormRelatorios
        .LstVendas.Clear
        .LstVendas.ColumnCount = 9
        .LstVendas.ColumnWidths = "45 pt;60 pt;135 pt;80 pt;35 pt;65 pt;55 pt;65 pt;55 pt"
        
        Dim dtPrimeiroDia As Date
        dtPrimeiroDia = DateSerial(Year(Date), Month(Date), 1)
        .TxtDataIni.Text = Format(dtPrimeiroDia, "dd/mm/yyyy")
        .TxtDataFim.Text = Format(Date, "dd/mm/yyyy")
        
        .CboStatus.Clear
        .CboStatus.AddItem "Todos"
        .CboStatus.AddItem "Concluída"
        .CboStatus.AddItem "Cancelada"
        .CboStatus.ListIndex = 0
        
        .CboFormaPagto.Clear
        .CboFormaPagto.AddItem "Todas"
        .CboFormaPagto.AddItem "Dinheiro"
        .CboFormaPagto.AddItem "PIX"
        .CboFormaPagto.AddItem "Cartão de Crédito"
        .CboFormaPagto.AddItem "Cartão de Débito"
        .CboFormaPagto.AddItem "Boleto Bancário"
        .CboFormaPagto.ListIndex = 0
        
        .TxtCliente.Text = ""
    End With
    
    bBloqueiaFiltro = False
    Call FiltrarVendas
End Sub

Sub FiltrarVendas()
    If bBloqueiaFiltro Then Exit Sub
    
    Dim ultLin As Long, r As Long
    ultLin = PlanVendas.Cells(PlanVendas.Rows.Count, "A").End(xlUp).Row
    
    With FormRelatorios
        .LstVendas.Clear
        
        If ultLin < 2 Then
            .TxtQtdVendas.Text = "0"
            .TxtTotalVendido.Text = "R$ 0,00"
            .TxtTotalTaxas.Text = "R$ 0,00"
            .TxtTotalReceber.Text = "R$ 0,00"
            Exit Sub
        End If
        
        Dim fCli As String, fStatus As String, fForma As String
        Dim dtIni As Date, dtFim As Date
        Dim temDtIni As Boolean, temDtFim As Boolean
        
        fCli = Trim(.TxtCliente.Text)
        fStatus = Trim(.CboStatus.Text)
        fForma = Trim(.CboFormaPagto.Text)
        
        If IsDate(.TxtDataIni.Text) Then
            dtIni = CDate(.TxtDataIni.Text)
            temDtIni = True
        End If
        
        If IsDate(.TxtDataFim.Text) Then
            dtFim = CDate(.TxtDataFim.Text)
            temDtFim = True
        End If
        
        Dim totQtd As Long, somaVendido As Double, somaTaxas As Double, somaReceber As Double
        totQtd = 0
        somaVendido = 0
        somaTaxas = 0
        somaReceber = 0
        
        Dim linDest As Long
        linDest = 0
        
        For r = 2 To ultLin
            Dim idV As String, dtV As Date, cliV As String, fpV As String
            Dim parcV As String, totLiqV As Double, taxaPctV As Double, liqRecV As Double, statusV As String
            
            idV = Format(PlanVendas.Cells(r, "A").Value, "0000")
            
            If IsDate(PlanVendas.Cells(r, "B").Value) Then
                dtV = CDate(PlanVendas.Cells(r, "B").Value)
            ElseIf IsDate(PlanVendas.Cells(r, "B").Text) Then
                dtV = CDate(PlanVendas.Cells(r, "B").Text)
            Else
                dtV = Date
            End If
            
            cliV = PlanVendas.Cells(r, "D").Text
            totLiqV = Val(Replace(CStr(PlanVendas.Cells(r, "G").Value), ",", "."))
            fpV = PlanVendas.Cells(r, "H").Text
            statusV = PlanVendas.Cells(r, "I").Text
            parcV = PlanVendas.Cells(r, "J").Text
            If parcV = "" Then parcV = "1x"
            
            taxaPctV = Val(Replace(CStr(PlanVendas.Cells(r, "K").Value), ",", "."))
            liqRecV = Val(Replace(CStr(PlanVendas.Cells(r, "L").Value), ",", "."))
            If liqRecV = 0 Then liqRecV = totLiqV
            
            Dim valorTaxaRS As Double
            valorTaxaRS = totLiqV - liqRecV
            If valorTaxaRS < 0 Then valorTaxaRS = 0
            
            Dim passa As Boolean
            passa = True
            
            If fCli <> "" Then
                If InStr(1, cliV, fCli, vbTextCompare) = 0 Then passa = False
            End If
            
            If passa And temDtIni Then
                If dtV < dtIni Then passa = False
            End If
            
            If passa And temDtFim Then
                If dtV > dtFim Then passa = False
            End If
            
            If passa And fStatus <> "Todos" And fStatus <> "" Then
                If StrComp(statusV, fStatus, vbTextCompare) <> 0 Then passa = False
            End If
            
            If passa And fForma <> "Todas" And fForma <> "" Then
                If StrComp(fpV, fForma, vbTextCompare) <> 0 Then passa = False
            End If
            
            If passa Then
                .LstVendas.AddItem
                .LstVendas.List(linDest, 0) = idV
                .LstVendas.List(linDest, 1) = Format(dtV, "dd/mm/yyyy")
                .LstVendas.List(linDest, 2) = cliV
                .LstVendas.List(linDest, 3) = fpV
                .LstVendas.List(linDest, 4) = parcV
                .LstVendas.List(linDest, 5) = Format(totLiqV, "R$ #,##0.00")
                .LstVendas.List(linDest, 6) = Format(valorTaxaRS, "R$ #,##0.00")
                .LstVendas.List(linDest, 7) = Format(liqRecV, "R$ #,##0.00")
                .LstVendas.List(linDest, 8) = statusV
                linDest = linDest + 1
                
                If statusV = "Concluída" Then
                    totQtd = totQtd + 1
                    somaVendido = somaVendido + totLiqV
                    somaTaxas = somaTaxas + valorTaxaRS
                    somaReceber = somaReceber + liqRecV
                End If
            End If
        Next r
        
        .TxtQtdVendas.Text = Format(totQtd, "#,##0") & " vendas"
        .TxtTotalVendido.Text = Format(somaVendido, "R$ #,##0.00")
        .TxtTotalTaxas.Text = Format(somaTaxas, "R$ #,##0.00")
        .TxtTotalReceber.Text = Format(somaReceber, "R$ #,##0.00")
    End With
End Sub

Sub AlternarStatusVenda()
    With FormRelatorios
        If .LstVendas.ListIndex = -1 Then
            MsgBox "Selecione uma venda na lista para alterar o status!", vbExclamation, "Atenção"
            Exit Sub
        End If
        
        Dim idSel As Long, statusAtual As String
        idSel = Val(.LstVendas.List(.LstVendas.ListIndex, 0))
        statusAtual = .LstVendas.List(.LstVendas.ListIndex, 8)
        
        Dim lin As Range
        Set lin = PlanVendas.Range("A:A").Find(What:=idSel, LookIn:=xlValues, LookAt:=xlWhole)
        If lin Is Nothing Then
            MsgBox "Venda Nº " & Format(idSel, "0000") & " não encontrada na planilha!", vbCritical, "Erro"
            Exit Sub
        End If
        
        Dim novoStatus As String, msgAviso As String
        If statusAtual = "Concluída" Then
            novoStatus = "Cancelada"
            msgAviso = "Deseja realmente CANCELAR a Venda Nº " & Format(idSel, "0000") & "?"
        Else
            novoStatus = "Concluída"
            msgAviso = "Deseja REATIVAR a Venda Nº " & Format(idSel, "0000") & " como Concluída?"
        End If
        
        If MsgBox(msgAviso, vbQuestion + vbYesNo, "Alterar Status da Venda") = vbYes Then
            lin.Offset(0, 8).Value = novoStatus
            MsgBox "Status da Venda Nº " & Format(idSel, "0000") & " alterado para: " & novoStatus, vbInformation, "Sucesso"
            Call FiltrarVendas
        End If
    End With
End Sub

Sub GerarRelatorioPDF()
    Dim wsR As Worksheet, wsCfg As Worksheet
    Set wsR = PlanRelVendas
    Set wsCfg = PlanConfig
    
    With FormRelatorios
        If .LstVendas.ListCount = 0 Then
            MsgBox "Não há dados filtrados para exportar o relatório!", vbExclamation, "Relatório Vazio"
            Exit Sub
        End If
        
        Dim ultLinR As Long
        ultLinR = wsR.Cells(wsR.Rows.Count, "A").End(xlUp).Row
        If ultLinR > 5 Then
            wsR.Range("A6:I" & ultLinR).Clear
        End If
        
        wsR.Range("A1").Value = wsCfg.Range("B3").Value
        wsR.Range("A2").Value = "RELATÓRIO GERENCIAL DE VENDAS"
        wsR.Range("A3").Value = "Período: " & .TxtDataIni.Text & " até " & .TxtDataFim.Text & _
                                " | Status: " & .CboStatus.Text & _
                                " | Emissão: " & Format(Now, "dd/mm/yyyy hh:nn")
                                
        Dim i As Long, linPlan As Long
        Dim somaV As Double, somaT As Double, somaR As Double
        linPlan = 6
        somaV = 0
        somaT = 0
        somaR = 0
        
        For i = 0 To .LstVendas.ListCount - 1
            Dim v1 As Double, v2 As Double, v3 As Double
            v1 = LimparMoeda(.LstVendas.List(i, 5))
            v2 = LimparMoeda(.LstVendas.List(i, 6))
            v3 = LimparMoeda(.LstVendas.List(i, 7))
            
            wsR.Cells(linPlan, 1).Value = .LstVendas.List(i, 0)
            wsR.Cells(linPlan, 2).Value = CDate(.LstVendas.List(i, 1))
            wsR.Cells(linPlan, 3).Value = .LstVendas.List(i, 2)
            wsR.Cells(linPlan, 4).Value = .LstVendas.List(i, 3)
            wsR.Cells(linPlan, 5).Value = .LstVendas.List(i, 4)
            wsR.Cells(linPlan, 6).Value = v1
            wsR.Cells(linPlan, 7).Value = v2
            wsR.Cells(linPlan, 8).Value = v3
            wsR.Cells(linPlan, 9).Value = .LstVendas.List(i, 8)
            
            If .LstVendas.List(i, 8) = "Concluída" Then
                somaV = somaV + v1
                somaT = somaT + v2
                somaR = somaR + v3
            End If
            
            wsR.Cells(linPlan, 1).HorizontalAlignment = -4108
            wsR.Cells(linPlan, 2).NumberFormat = "dd/mm/yyyy"
            wsR.Cells(linPlan, 2).HorizontalAlignment = -4108
            wsR.Cells(linPlan, 5).HorizontalAlignment = -4108
            wsR.Cells(linPlan, 6).NumberFormat = "R$ #,##0.00"
            wsR.Cells(linPlan, 7).NumberFormat = "R$ #,##0.00"
            wsR.Cells(linPlan, 8).NumberFormat = "R$ #,##0.00"
            wsR.Cells(linPlan, 9).HorizontalAlignment = -4108
            
            If .LstVendas.List(i, 8) = "Cancelada" Then
                wsR.Range(wsR.Cells(linPlan, 1), wsR.Cells(linPlan, 9)).Font.Color = 255
            End If
            
            linPlan = linPlan + 1
        Next i
        
        wsR.Cells(linPlan, 1).Value = "TOTAIS (Concluídas):"
        wsR.Cells(linPlan, 6).Value = somaV
        wsR.Cells(linPlan, 7).Value = somaT
        wsR.Cells(linPlan, 8).Value = somaR
        
        With wsR.Range(wsR.Cells(linPlan, 1), wsR.Cells(linPlan, 9))
            .Font.Bold = True
            .Interior.Color = 15790320
            .NumberFormat = "R$ #,##0.00"
        End With
        wsR.Cells(linPlan, 1).NumberFormat = "@"
        
        With wsR.Range("A5:I" & linPlan).Borders
            .LineStyle = 1
            .Color = 12632256
            .Weight = 2
        End With
        
        Dim pastaRel As String, nomeArq As String, caminhoPDF As String
        pastaRel = ObterPastaLocalRelatorios()
        
        nomeArq = "Relatorio_Vendas_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
        caminhoPDF = pastaRel & "\" & nomeArq
        
        On Error Resume Next
        wsR.ExportAsFixedFormat Type:=0, _
                                Filename:=caminhoPDF, _
                                Quality:=0, _
                                IncludeDocProperties:=True, _
                                IgnorePrintAreas:=False, _
                                OpenAfterPublish:=True
                                
        If Err.Number <> 0 Then
            MsgBox "Não foi possível gerar o PDF: " & Err.Description, vbCritical, "Erro de Exportação"
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
                                
        MsgBox "Relatório em PDF gerado com sucesso!" & vbCrLf & vbCrLf & _
               "Salvo em: " & caminhoPDF, vbInformation, "Relatório PDF Gerado"
    End With
End Sub

Sub GerarComprovanteVendaPDF()
    With FormRelatorios
        If .LstVendas.ListIndex = -1 Then
            MsgBox "Selecione uma venda na lista para gerar o comprovante!", vbExclamation, "Atenção"
            Exit Sub
        End If
        
        Dim idSel As Long
        idSel = Val(.LstVendas.List(.LstVendas.ListIndex, 0))
        
        Dim linV As Range
        Set linV = PlanVendas.Range("A:A").Find(What:=idSel, LookIn:=xlValues, LookAt:=xlWhole)
        If linV Is Nothing Then Exit Sub
        
        Dim wsC As Worksheet
        On Error Resume Next
        Set wsC = ThisWorkbook.Sheets("CUPOM_TEMP")
        If wsC Is Nothing Then
            Set wsC = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wsC.Name = "CUPOM_TEMP"
        Else
            wsC.Cells.Clear
        End If
        On Error GoTo 0
        
        wsC.PageSetup.Orientation = 1
        wsC.PageSetup.PaperSize = 9
        wsC.PageSetup.Zoom = False
        wsC.PageSetup.FitToPagesWide = 1
        wsC.PageSetup.FitToPagesTall = 1
        
        Dim wsCfg As Worksheet
        Set wsCfg = PlanConfig
        
        wsC.Range("A1:E1").Merge
        wsC.Range("A1").Value = wsCfg.Range("B3").Value
        wsC.Range("A1").Font.Size = 13
        wsC.Range("A1").Font.Bold = True
        wsC.Range("A1").HorizontalAlignment = -4108
        
        wsC.Range("A2:E2").Merge
        wsC.Range("A2").Value = "CNPJ: " & wsCfg.Range("B4").Value & " - Tel: " & wsCfg.Range("B5").Value
        wsC.Range("A2").Font.Size = 9
        wsC.Range("A2").HorizontalAlignment = -4108
        
        wsC.Range("A3:E3").Merge
        wsC.Range("A3").Value = wsCfg.Range("B7").Value & " - " & wsCfg.Range("B8").Value
        wsC.Range("A3").Font.Size = 9
        wsC.Range("A3").HorizontalAlignment = -4108
        
        wsC.Range("A4:E4").Merge
        wsC.Range("A4").Value = "--------------------------------------------------------------------------------------------------"
        wsC.Range("A4").HorizontalAlignment = -4108
        
        wsC.Range("A5:E5").Merge
        wsC.Range("A5").Value = "COMPROVANTE DE VENDA Nº " & Format(idSel, "0000")
        wsC.Range("A5").Font.Size = 11
        wsC.Range("A5").Font.Bold = True
        wsC.Range("A5").HorizontalAlignment = -4108
        
        wsC.Cells(6, 1).Value = "Data: " & Format(linV.Offset(0, 1).Value, "dd/mm/yyyy")
        wsC.Cells(6, 4).Value = "Status: " & linV.Offset(0, 8).Value
        wsC.Range("A7:E7").Merge
        wsC.Cells(7, 1).Value = "Cliente: " & linV.Offset(0, 3).Value & " (Cód: " & linV.Offset(0, 2).Value & ")"
        
        wsC.Cells(9, 1).Value = "Item"
        wsC.Cells(9, 2).Value = "Descrição do Produto"
        wsC.Cells(9, 3).Value = "Preço Unit."
        wsC.Cells(9, 4).Value = "Qtd"
        wsC.Cells(9, 5).Value = "Subtotal"
        wsC.Range("A9:E9").Font.Bold = True
        wsC.Range("A9:E9").Interior.Color = 3355443
        wsC.Range("A9:E9").Font.Color = 16777215
        
        Dim ultLinI As Long, rI As Long, linItemC As Long, itemNum As Long
        ultLinI = PlanItens.Cells(PlanItens.Rows.Count, "A").End(xlUp).Row
        linItemC = 10
        itemNum = 1
        
        For rI = 2 To ultLinI
            If PlanItens.Cells(rI, "B").Value = idSel Then
                wsC.Cells(linItemC, 1).Value = itemNum
                wsC.Cells(linItemC, 2).Value = PlanItens.Cells(rI, "D").Value
                wsC.Cells(linItemC, 3).Value = PlanItens.Cells(rI, "E").Value
                wsC.Cells(linItemC, 4).Value = PlanItens.Cells(rI, "F").Value
                wsC.Cells(linItemC, 5).Value = PlanItens.Cells(rI, "G").Value
                
                wsC.Cells(linItemC, 3).NumberFormat = "R$ #,##0.00"
                wsC.Cells(linItemC, 5).NumberFormat = "R$ #,##0.00"
                itemNum = itemNum + 1
                linItemC = linItemC + 1
            End If
        Next rI
        
        linItemC = linItemC + 1
        wsC.Cells(linItemC, 4).Value = "TOTAL BRUTO:"
        wsC.Cells(linItemC, 5).Value = linV.Offset(0, 4).Value
        wsC.Cells(linItemC, 5).NumberFormat = "R$ #,##0.00"
        
        linItemC = linItemC + 1
        wsC.Cells(linItemC, 4).Value = "DESCONTO:"
        wsC.Cells(linItemC, 5).Value = linV.Offset(0, 5).Value
        wsC.Cells(linItemC, 5).NumberFormat = "R$ #,##0.00"
        
        linItemC = linItemC + 1
        wsC.Cells(linItemC, 4).Value = "TOTAL PAGO:"
        wsC.Cells(linItemC, 5).Value = linV.Offset(0, 6).Value
        wsC.Cells(linItemC, 5).NumberFormat = "R$ #,##0.00"
        wsC.Range(wsC.Cells(linItemC, 4), wsC.Cells(linItemC, 5)).Font.Bold = True
        
        linItemC = linItemC + 1
        wsC.Range(wsC.Cells(linItemC, 1), wsC.Cells(linItemC, 5)).Merge
        wsC.Cells(linItemC, 1).Value = "Forma de Pagamento: " & linV.Offset(0, 7).Value & " (" & linV.Offset(0, 9).Value & ")"
        wsC.Cells(linItemC, 1).Font.Italic = True
        
        linItemC = linItemC + 2
        wsC.Range(wsC.Cells(linItemC, 1), wsC.Cells(linItemC, 5)).Merge
        wsC.Cells(linItemC, 1).Value = wsCfg.Range("B9").Value
        wsC.Cells(linItemC, 1).HorizontalAlignment = -4108
        wsC.Cells(linItemC, 1).Font.Bold = True
        
        wsC.Columns.AutoFit
        
        Dim pastaRel As String, caminhoPDF As String
        pastaRel = ObterPastaLocalRelatorios()
        
        caminhoPDF = pastaRel & "\" & "Comprovante_Venda_" & Format(idSel, "0000") & ".pdf"
        
        On Error Resume Next
        wsC.ExportAsFixedFormat Type:=0, _
                                Filename:=caminhoPDF, _
                                Quality:=0, _
                                IncludeDocProperties:=True, _
                                IgnorePrintAreas:=False, _
                                OpenAfterPublish:=True
                                
        If Err.Number <> 0 Then
            MsgBox "Não foi possível gerar o comprovante em PDF: " & Err.Description, vbCritical, "Erro de Exportação"
            Err.Clear
        Else
            MsgBox "Comprovante de Venda gerado com sucesso em PDF!" & vbCrLf & vbCrLf & _
                   "Arquivo: " & caminhoPDF, vbInformation, "Comprovante Emitido"
        End If
        On Error GoTo 0
                                
        Application.DisplayAlerts = False
        wsC.Delete
        Application.DisplayAlerts = True
    End With
End Sub



