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
    LimparMoeda = ConverterParaDouble(valStr)
End Function

Public Function ObterPastaLocalRelatorios() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim caminhoBase As String
    caminhoBase = ThisWorkbook.Path
    
    If InStr(1, caminhoBase, "http://", vbTextCompare) > 0 Or InStr(1, caminhoBase, "https://", vbTextCompare) > 0 Then
        Dim userProf As String
        userProf = Environ("USERPROFILE")
        If userProf <> "" Then
            Dim candidata As String
            candidata = userProf & "\OneDrive\Área de Trabalho\SISTEMA DE VENDAS"
            If fso.FolderExists(candidata) Then
                caminhoBase = candidata
            Else
                candidata = userProf & "\Desktop\SISTEMA DE VENDAS"
                If fso.FolderExists(candidata) Then
                    caminhoBase = candidata
                Else
                    caminhoBase = userProf & "\Documents"
                End If
            End If
        Else
            caminhoBase = "C:\Temp"
        End If
    End If
    
    Dim pastaRel As String
    pastaRel = caminhoBase & "\Relatorios"
    If Not fso.FolderExists(pastaRel) Then
        On Error Resume Next
        fso.CreateFolder pastaRel
        On Error GoTo 0
    End If
    
    ObterPastaLocalRelatorios = pastaRel
End Function

Public Function ObterCaminhoLogo() As String
    On Error Resume Next
    Dim wsCfg As Worksheet
    Set wsCfg = PlanConfig()
    
    Dim caminho As String
    caminho = ""
    
    If Not wsCfg Is Nothing Then
        Dim cel As Range
        Set cel = wsCfg.Range("A:A").Find(What:="LOGO_PATH", LookIn:=xlValues, LookAt:=xlWhole)
        If Not cel Is Nothing Then
            caminho = Trim(cel.Offset(0, 1).Text)
        End If
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If caminho <> "" Then
        If Not fso.FileExists(caminho) Then
            Dim tentativa As String
            tentativa = ThisWorkbook.Path & "\" & caminho
            If fso.FileExists(tentativa) Then
                caminho = tentativa
            Else
                caminho = ""
            End If
        End If
    End If
    
    If caminho = "" Then
        Dim extArr As Variant, e As Variant
        extArr = Array("logo_empresa.png", "logo_empresa.bmp", "logo.png", "logo.bmp", "logo.jpg")
        For Each e In extArr
            If fso.FileExists(ThisWorkbook.Path & "\" & e) Then
                caminho = ThisWorkbook.Path & "\" & e
                Exit For
            End If
        Next e
    End If
    
    ObterCaminhoLogo = caminho
    On Error GoTo 0
End Function

Public Sub SelecionarLogoEmpresa()
    If Not ModSeguranca.VerificarAutorizacaoAdmin("Configurar Logotipo da Empresa") Then Exit Sub
    
    Dim arq As Variant
    arq = Application.GetOpenFilename("Imagens de Logotipo (*.png; *.jpg; *.jpeg; *.bmp), *.png; *.jpg; *.jpeg; *.bmp", , "Selecione o Logotipo da Empresa")
    
    If arq = False Or Trim(arq) = "" Then Exit Sub
    
    Dim wsCfg As Worksheet
    Set wsCfg = PlanConfig()
    If wsCfg Is Nothing Then Exit Sub
    
    Dim cel As Range
    Set cel = wsCfg.Range("A:A").Find(What:="LOGO_PATH", LookIn:=xlValues, LookAt:=xlWhole)
    If cel Is Nothing Then
        Dim ultLin As Long
        ultLin = wsCfg.Cells(wsCfg.Rows.Count, "A").End(xlUp).Row + 1
        wsCfg.Cells(ultLin, "A").Value = "LOGO_PATH"
        wsCfg.Cells(ultLin, "B").Value = CStr(arq)
    Else
        cel.Offset(0, 1).Value = CStr(arq)
    End If
    
    MsgBox "Logotipo configurado com sucesso!" & vbCrLf & "Arquivo: " & arq, vbInformation, "Logotipo Atualizado"
    
    On Error Resume Next
    Call CarregarLogoEmImage(FormVendas.ImgLogo)
    Call CarregarLogoEmImage(FormRelatorios.ImgLogo)
    On Error GoTo 0
End Sub

Public Sub CarregarLogoEmImage(ByVal imgCtrl As Object)
    On Error Resume Next
    If imgCtrl Is Nothing Then Exit Sub
    
    Dim caminho As String
    caminho = ObterCaminhoLogo()
    
    If caminho <> "" Then
        Dim ext As String
        ext = LCase(Right(caminho, 4))
        If ext = ".bmp" Or ext = ".jpg" Or ext = "jpeg" Or ext = ".ico" Or ext = ".gif" Then
            Set imgCtrl.Picture = LoadPicture(caminho)
            imgCtrl.PictureSizeMode = 3 ' fmPictureSizeModeZoom
            imgCtrl.Visible = True
        ElseIf ext = ".png" Then
            ' Se arquivo for png e existir cópia bmp no mesmo diretório, carrega o bmp
            Dim bmpAlt As String
            bmpAlt = Left(caminho, Len(caminho) - 4) & ".bmp"
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(bmpAlt) Then
                Set imgCtrl.Picture = LoadPicture(bmpAlt)
                imgCtrl.PictureSizeMode = 3
                imgCtrl.Visible = True
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub InserirLogoEmPlanilha(ByVal ws As Worksheet, ByVal pLeft As Double, ByVal pTop As Double, ByVal pWidth As Double, ByVal pHeight As Double)
    On Error Resume Next
    Dim caminho As String
    caminho = ObterCaminhoLogo()
    If caminho = "" Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(caminho) Then Exit Sub
    
    Dim shp As Shape
    Set shp = ws.Shapes.AddPicture(caminho, msoFalse, msoTrue, pLeft, pTop, pWidth, pHeight)
    If Not shp Is Nothing Then
        shp.LockAspectRatio = msoTrue
    End If
    On Error GoTo 0
End Sub

Sub AbrirRelatorios()
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        Call ModSeguranca.InicializarSessaoPadrao
    End If
    FormRelatorios.Show
End Sub

Sub InicializarFormRelatorios()
    bBloqueiaFiltro = True
    
    If Trim(ModSeguranca.vUsuarioAtivo) = "" Then
        Call ModSeguranca.InicializarSessaoPadrao
    End If
    Call ModSeguranca.AtualizarBadgesFormsAbertos
    Call CarregarLogoEmImage(FormRelatorios.ImgLogo)
    Call ModSeguranca.ConfigurarPermissoesRelatorios
    
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
            dtV = CDate(PlanVendas.Cells(r, "B").Value)
            cliV = PlanVendas.Cells(r, "D").Text
            totLiqV = Val(Replace(PlanVendas.Cells(r, "G").Value, ",", "."))
            fpV = PlanVendas.Cells(r, "H").Text
            statusV = PlanVendas.Cells(r, "I").Text
            parcV = PlanVendas.Cells(r, "J").Text
            If parcV = "" Then parcV = "1x"
            
            taxaPctV = Val(Replace(PlanVendas.Cells(r, "K").Value, ",", "."))
            liqRecV = Val(Replace(PlanVendas.Cells(r, "L").Value, ",", "."))
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
                If statusV <> fStatus Then passa = False
            End If
            
            If passa And fForma <> "Todas" And fForma <> "" Then
                If fpV <> fForma Then passa = False
            End If
            
            If passa Then
                .LstVendas.AddItem
                .LstVendas.List(linDest, 0) = idV
                .LstVendas.List(linDest, 1) = Format(dtV, "dd/mm/yyyy")
                .LstVendas.List(linDest, 2) = cliV
                .LstVendas.List(linDest, 3) = fpV
                .LstVendas.List(linDest, 4) = parcV
                .LstVendas.List(linDest, 5) = Format(totLiqV, "R$ #,##0.00")
                .LstVendas.List(linDest, 6) = Format(taxaPctV, "0.00%")
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
        
        .TxtQtdVendas.Text = Format(totQtd, "#,##0")
        
        If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
            .TxtTotalVendido.Text = "R$ ***"
            .TxtTotalTaxas.Text = "R$ ***"
            .TxtTotalReceber.Text = "R$ ***"
        Else
            .TxtTotalVendido.Text = Format(somaVendido, "R$ #,##0.00")
            .TxtTotalTaxas.Text = Format(somaTaxas, "R$ #,##0.00")
            .TxtTotalReceber.Text = Format(somaReceber, "R$ #,##0.00")
        End If
    End With
End Sub

Sub AlternarStatusVenda()
    With FormRelatorios
        If .LstVendas.ListIndex = -1 Then
            MsgBox "Selecione uma venda na lista para alterar o status!", vbExclamation, "Atenção"
            Exit Sub
        End If
        
        Dim idSel As Long, statusAtual As String, novoStatus As String
        idSel = Val(.LstVendas.List(.LstVendas.ListIndex, 0))
        statusAtual = .LstVendas.List(.LstVendas.ListIndex, 8)
        
        If statusAtual = "Concluída" Then
            novoStatus = "Cancelada"
        Else
            novoStatus = "Concluída"
        End If
        
        If MsgBox("Deseja realmente alterar o status da Venda Nº " & Format(idSel, "0000") & _
                  " de [" & statusAtual & "] para [" & novoStatus & "]?", _
                  vbQuestion + vbYesNo, "Confirmar Alteração de Status") = vbNo Then
            Exit Sub
        End If
        
        Dim linV As Range
        Set linV = PlanVendas.Range("A:A").Find(What:=idSel, LookIn:=xlValues, LookAt:=xlWhole)
        
        If linV Is Nothing Then
            MsgBox "Venda não encontrada na base de dados!", vbCritical, "Erro"
            Exit Sub
        End If
        
        linV.Offset(0, 8).Value = novoStatus
        
        MsgBox "Status da Venda Nº " & Format(idSel, "0000") & " atualizado para [" & novoStatus & "] com sucesso!", _
               vbInformation, "Status Atualizado"
               
        Call FiltrarVendas
    End With
End Sub

Sub GerarRelatorioPDF()
    Dim wsRel As Worksheet
    Set wsRel = PlanRelVendas()
    
    Dim wsCfg As Worksheet
    Set wsCfg = PlanConfig()
    
    wsRel.Cells.Clear
    
    ' Configurações de Página
    wsRel.PageSetup.Orientation = 2 ' Paisagem
    wsRel.PageSetup.PaperSize = 9   ' A4
    wsRel.PageSetup.Zoom = False
    wsRel.PageSetup.FitToPagesWide = 1
    wsRel.PageSetup.FitToPagesTall = False
    wsRel.PageSetup.PrintTitleRows = "$1:$5"
    wsRel.PageSetup.LeftMargin = Application.InchesToPoints(0.4)
    wsRel.PageSetup.RightMargin = Application.InchesToPoints(0.4)
    wsRel.PageSetup.TopMargin = Application.InchesToPoints(0.5)
    wsRel.PageSetup.BottomMargin = Application.InchesToPoints(0.5)
    
    ' Cabeçalho
    wsRel.Range("A1:I1").Merge
    wsRel.Range("A1").Value = wsCfg.Range("B3").Value
    wsRel.Range("A1").Font.Size = 14
    wsRel.Range("A1").Font.Bold = True
    wsRel.Range("A1").Font.Name = "Segoe UI"
    wsRel.Range("A1").Interior.Color = 3877150
    wsRel.Range("A1").Font.Color = 16777215
    wsRel.Range("A1").HorizontalAlignment = -4108
    
    wsRel.Range("A2:I2").Merge
    wsRel.Range("A2").Value = "RELATÓRIO GERENCIAL DE VENDAS"
    wsRel.Range("A2").Font.Size = 11
    wsRel.Range("A2").Font.Bold = True
    wsRel.Range("A2").Font.Name = "Segoe UI"
    wsRel.Range("A2").Interior.Color = 5263440
    wsRel.Range("A2").Font.Color = 16777215
    wsRel.Range("A2").HorizontalAlignment = -4108
    
    Dim strFiltros As String
    strFiltros = "Período: " & FormRelatorios.TxtDataIni.Text & " até " & FormRelatorios.TxtDataFim.Text & _
                 " | Status: " & FormRelatorios.CboStatus.Text & _
                 " | Forma: " & FormRelatorios.CboFormaPagto.Text & _
                 " | Emitido por: " & ModSeguranca.vNomeOperador & " em " & Format(Now, "dd/mm/yyyy hh:mm")
                 
    wsRel.Range("A3:I3").Merge
    wsRel.Range("A3").Value = strFiltros
    wsRel.Range("A3").Font.Size = 8.5
    wsRel.Range("A3").Font.Name = "Segoe UI"
    wsRel.Range("A3").HorizontalAlignment = -4108
    
    ' Inserir logotipo no cabeçalho se houver
    Call InserirLogoEmPlanilha(wsRel, 15, 6, 75, 40)
    
    ' Tabela de Dados
    Dim headers_rel As Variant
    headers_rel = Array("Nº Venda", "Data", "Cliente", "Forma Pagto", "Parc.", "Total Venda", "Taxa Cartão", "Líq. Receber", "Status")
    Dim col_i As Long
    For col_i = 0 To UBound(headers_rel)
        Dim cellH As Range
        Set cellH = wsRel.Cells(5, col_i + 1)
        cellH.Value = headers_rel(col_i)
        cellH.Font.Bold = True
        cellH.Font.Name = "Segoe UI"
        cellH.Font.Size = 9.5
        cellH.Interior.Color = 3877150
        cellH.Font.Color = 16777215
        cellH.HorizontalAlignment = -4108
    Next col_i
    
    Dim col_widths As Variant
    col_widths = Array(12, 14, 30, 18, 10, 16, 14, 16, 14)
    For col_i = 0 To UBound(col_widths)
        wsRel.Columns(col_i + 1).ColumnWidth = col_widths(col_i)
    Next col_i
    
    Dim linRel As Long, i As Long
    linRel = 6
    
    With FormRelatorios.LstVendas
        If .ListCount = 0 Then
            wsRel.Range("A6:I6").Merge
            wsRel.Range("A6").Value = "Nenhuma venda localizada com os filtros selecionados."
            wsRel.Range("A6").HorizontalAlignment = -4108
            wsRel.Range("A6").Font.Italic = True
            linRel = 7
        Else
            For i = 0 To .ListCount - 1
                wsRel.Cells(linRel, 1).Value = .List(i, 0)
                wsRel.Cells(linRel, 2).Value = .List(i, 1)
                wsRel.Cells(linRel, 3).Value = .List(i, 2)
                wsRel.Cells(linRel, 4).Value = .List(i, 3)
                wsRel.Cells(linRel, 5).Value = .List(i, 4)
                wsRel.Cells(linRel, 6).Value = LimparMoeda(.List(i, 5))
                
                Dim txStr As String
                txStr = Replace(.List(i, 6), "%", "")
                txStr = Replace(txStr, ",", Application.International(xlDecimalSeparator))
                txStr = Replace(txStr, ".", Application.International(xlDecimalSeparator))
                If IsNumeric(txStr) Then
                    wsRel.Cells(linRel, 7).Value = CDbl(txStr) / 100
                Else
                    wsRel.Cells(linRel, 7).Value = 0
                End If
                
                wsRel.Cells(linRel, 8).Value = LimparMoeda(.List(i, 7))
                wsRel.Cells(linRel, 9).Value = .List(i, 8)
                
                wsRel.Cells(linRel, 1).HorizontalAlignment = -4108
                wsRel.Cells(linRel, 2).HorizontalAlignment = -4108
                wsRel.Cells(linRel, 4).HorizontalAlignment = -4108
                wsRel.Cells(linRel, 5).HorizontalAlignment = -4108
                wsRel.Cells(linRel, 9).HorizontalAlignment = -4108
                
                wsRel.Cells(linRel, 6).NumberFormat = "R$ #,##0.00"
                wsRel.Cells(linRel, 7).NumberFormat = "0.00%"
                wsRel.Cells(linRel, 8).NumberFormat = "R$ #,##0.00"
                
                If i Mod 2 = 1 Then
                    wsRel.Range(wsRel.Cells(linRel, 1), wsRel.Cells(linRel, 9)).Interior.Color = 16119285
                End If
                
                linRel = linRel + 1
            Next i
        End If
    End With
    
    ' Linha de Totais
    wsRel.Cells(linRel, 1).Value = "TOTAIS:"
    wsRel.Range(wsRel.Cells(linRel, 1), wsRel.Cells(linRel, 5)).Merge
    wsRel.Range(wsRel.Cells(linRel, 1), wsRel.Cells(linRel, 5)).HorizontalAlignment = -4152
    wsRel.Cells(linRel, 1).Font.Bold = True
    
    wsRel.Cells(linRel, 6).Formula = "=SUM(F6:F" & (linRel - 1) & ")"
    wsRel.Cells(linRel, 6).NumberFormat = "R$ #,##0.00"
    wsRel.Cells(linRel, 6).Font.Bold = True
    
    wsRel.Cells(linRel, 8).Formula = "=SUM(H6:H" & (linRel - 1) & ")"
    wsRel.Cells(linRel, 8).NumberFormat = "R$ #,##0.00"
    wsRel.Cells(linRel, 8).Font.Bold = True
    
    wsRel.Range(wsRel.Cells(linRel, 1), wsRel.Cells(linRel, 9)).Interior.Color = 15132390
    
    ' Bordas
    Dim rngTabela As Range
    Set rngTabela = wsRel.Range(wsRel.Cells(5, 1), wsRel.Cells(linRel, 9))
    rngTabela.Borders.LineStyle = 1
    rngTabela.Borders.Color = 12632256
    
    Dim pastaRel As String, caminhoPDF As String
    pastaRel = ObterPastaLocalRelatorios()
    caminhoPDF = pastaRel & "\Relatorio_Vendas_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    
    On Error Resume Next
    wsRel.ExportAsFixedFormat Type:=0, _
                             Filename:=caminhoPDF, _
                             Quality:=0, _
                             IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, _
                             OpenAfterPublish:=True
                             
    If Err.Number <> 0 Then
        MsgBox "Não foi possível gerar o PDF: " & Err.Description, vbCritical, "Erro de Exportação"
        Err.Clear
    Else
        MsgBox "Relatório Gerencial gerado com sucesso em PDF!" & vbCrLf & vbCrLf & _
               "Arquivo: " & caminhoPDF, vbInformation, "Relatório Emitido"
    End If
    On Error GoTo 0
End Sub

Sub GerarComprovanteVendaPDF()
    With FormRelatorios
        If .LstVendas.ListIndex = -1 Then
            MsgBox "Selecione uma venda na lista para gerar o comprovante!", vbExclamation, "Atenção"
            Exit Sub
        End If
        
        Dim idSel As Long
        idSel = Val(.LstVendas.List(.LstVendas.ListIndex, 0))
        Call GerarComprovanteVendaPorID(idSel)
    End With
End Sub

Public Sub GerarComprovanteVendaPorID(ByVal idSel As Long)
    Dim linV As Range
    Set linV = PlanVendas.Range("A:A").Find(What:=idSel, LookIn:=xlValues, LookAt:=xlWhole)
    If linV Is Nothing Then
        MsgBox "Venda Nº " & Format(idSel, "0000") & " não encontrada!", vbCritical, "Erro"
        Exit Sub
    End If
    
    Dim wsC As Worksheet
    On Error Resume Next
    Set wsC = ThisWorkbook.Sheets("CUPOM_TEMP")
    If wsC Is Nothing Then
        Set wsC = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsC.Name = "CUPOM_TEMP"
    Else
        wsC.Cells.Clear
        Dim shpItem As Shape
        For Each shpItem In wsC.Shapes
            shpItem.Delete
        Next shpItem
    End If
    On Error GoTo 0
    
    wsC.PageSetup.Orientation = 1
    wsC.PageSetup.PaperSize = 9
    wsC.PageSetup.Zoom = False
    wsC.PageSetup.FitToPagesWide = 1
    wsC.PageSetup.FitToPagesTall = 1
    
    Dim wsCfg As Worksheet
    Set wsCfg = PlanConfig()
    
    ' Logotipo no comprovante
    Call InserirLogoEmPlanilha(wsC, 18, 8, 65, 36)
    
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
    
    ' OPERADOR RESPONSÁVEL (Coluna M da aba VENDAS)
    Dim opNome As String
    opNome = Trim(linV.Offset(0, 12).Text)
    If opNome = "" Then opNome = "Atendimento Geral"
    wsC.Range("A8:E8").Merge
    wsC.Cells(8, 1).Value = "Atendido por: " & opNome
    wsC.Cells(8, 1).Font.Italic = True
    wsC.Cells(8, 1).Font.Bold = True
    wsC.Cells(8, 1).Font.Size = 9
    
    wsC.Cells(10, 1).Value = "Item"
    wsC.Cells(10, 2).Value = "Descrição do Produto"
    wsC.Cells(10, 3).Value = "Preço Unit."
    wsC.Cells(10, 4).Value = "Qtd"
    wsC.Cells(10, 5).Value = "Subtotal"
    wsC.Range("A10:E10").Font.Bold = True
    wsC.Range("A10:E10").Interior.Color = 3877150
    wsC.Range("A10:E10").Font.Color = 16777215
    
    Dim ultLinI As Long, rI As Long, linItemC As Long, itemNum As Long
    ultLinI = PlanItens.Cells(PlanItens.Rows.Count, "A").End(xlUp).Row
    linItemC = 11
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
    caminhoPDF = pastaRel & "\Comprovante_Venda_" & Format(idSel, "0000") & ".pdf"
    
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
               "• Operador: " & opNome & vbCrLf & _
               "• Arquivo: " & caminhoPDF, vbInformation, "Comprovante Emitido"
    End If
    On Error GoTo 0
                            
    Application.DisplayAlerts = False
    wsC.Delete
    Application.DisplayAlerts = True
End Sub


