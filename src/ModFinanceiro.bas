Attribute VB_Name = "ModFinanceiro"
Option Explicit
Option Compare Text

' =============================================================================
' MÓDULO FINANCEIRO: GESTÃO DE DESPESAS, DRE E FLUXO DE CAIXA (FASES 2 & 3)
' =============================================================================

Public Function PlanDespesas() As Worksheet
    On Error Resume Next
    Set PlanDespesas = ThisWorkbook.Sheets("DESPESAS")
    On Error GoTo 0
End Function

Public Function PlanDRE() As Worksheet
    On Error Resume Next
    Set PlanDRE = ThisWorkbook.Sheets("APURACAO_RESULTADOS")
    On Error GoTo 0
End Function

Public Function PlanFluxoCaixa() As Worksheet
    On Error Resume Next
    Set PlanFluxoCaixa = ThisWorkbook.Sheets("FLUXO_CAIXA")
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' NAVEGAÇÃO COM CONTROLE DE ACESSO (ADMIN / PIN MASCARADO)
' -----------------------------------------------------------------------------
Public Sub AbrirPainelDespesas()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Acessar Painel de Despesas e Contas a Pagar") Then
            Exit Sub
        End If
    End If
    FormDespesas.Show
End Sub

Public Sub IrParaDRE()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Visualizar Apuração de Resultados (DRE)") Then
            Exit Sub
        End If
        ThisWorkbook.Sheets("APURACAO_RESULTADOS").Visible = -1
    End If
    On Error Resume Next
    Call AtualizarDemonstrativos
    ThisWorkbook.Sheets("APURACAO_RESULTADOS").Visible = -1
    ThisWorkbook.Sheets("APURACAO_RESULTADOS").Activate
    On Error GoTo 0
End Sub

Public Sub IrParaFluxoCaixa()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Visualizar Demonstrativo de Fluxo de Caixa") Then
            Exit Sub
        End If
        ThisWorkbook.Sheets("FLUXO_CAIXA").Visible = -1
    End If
    On Error Resume Next
    Call AtualizarDemonstrativos
    ThisWorkbook.Sheets("FLUXO_CAIXA").Visible = -1
    ThisWorkbook.Sheets("FLUXO_CAIXA").Activate
    On Error GoTo 0
End Sub

Public Sub IrParaAbaDespesas()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Visualizar Base de Dados de Despesas") Then
            Exit Sub
        End If
        ThisWorkbook.Sheets("DESPESAS").Visible = -1
    End If
    On Error Resume Next
    ThisWorkbook.Sheets("DESPESAS").Visible = -1
    ThisWorkbook.Sheets("DESPESAS").Activate
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' MAPEAMENTO AUTOMÁTICO: CATEGORIA DE FLUXO -> GRUPO DRE
' -----------------------------------------------------------------------------
Public Function MapearGrupoDREPorCategoria(ByVal catFluxo As String) As String
    Dim c As String
    c = Trim(catFluxo)
    Select Case c
        Case "Gastos com produtos para venda"
            MapearGrupoDREPorCategoria = "4.1 (-) Custo das Mercadorias Vendidas"
        Case "Gastos com serviços prestados", "Gastos com serviços prestados ***"
            MapearGrupoDREPorCategoria = "4.2 (-) Custo dos Serviços Prestados"
        Case "Carnê de arrecadação - MEI", "DAS - Documento de Arrecadação Simples", "Outros tributos (impostos, taxas, etc.)"
            MapearGrupoDREPorCategoria = "2.2 (-) Impostos e Contribuições sobre vendas"
        Case "Gastos com juros e taxas bancos"
            MapearGrupoDREPorCategoria = "8.1 Despesas Financeiras"
        Case "Serviços contábeis"
            MapearGrupoDREPorCategoria = "11.1 Serviços de contabilidade"
        Case "Manutenção do veículo da empresa"
            MapearGrupoDREPorCategoria = "11.2 Manutenção de veículos"
        Case "Manutenção da empresa"
            MapearGrupoDREPorCategoria = "11.3 Manutenção da empresa"
        Case Else
            ' Aluguel, Energia, Água, Telefone, Internet, Folha, Retirada, Combustível, Segurança, Diversas
            MapearGrupoDREPorCategoria = "6.2 Despesas Administrativas"
    End Select
End Function

' -----------------------------------------------------------------------------
' MOTOR DE RECÁLCULO AUTOMÁTICO DA DRE E FLUXO DE CAIXA
' -----------------------------------------------------------------------------
Public Sub AtualizarDemonstrativos()
    On Error GoTo FimAtualizacao
    Application.ScreenUpdating = False
    
    Dim wsD As Worksheet, wsDRE As Worksheet, wsFC As Worksheet
    Dim wsV As Worksheet, wsI As Worksheet, wsP As Worksheet
    Set wsD = PlanDespesas()
    Set wsDRE = PlanDRE()
    Set wsFC = PlanFluxoCaixa()
    Set wsV = ThisWorkbook.Sheets("VENDAS")
    Set wsI = ThisWorkbook.Sheets("ITENS_VENDA")
    Set wsP = ThisWorkbook.Sheets("PRODUTOS_SERVICOS")
    
    If wsD Is Nothing Or wsDRE Is Nothing Or wsFC Is Nothing Then GoTo FimAtualizacao
    
    ' 1. Determinar Ano Base (padrão 2026)
    Dim anoBase As Long
    anoBase = Val(wsDRE.Range("E5").Value)
    If anoBase < 2000 Or anoBase > 2100 Then
        anoBase = Year(Date)
        wsDRE.Range("E5").Value = anoBase
    End If
    wsFC.Range("E5").Value = anoBase
    
    ' 2. Sincronizar colunas auxiliares (MES_COMP, ANO_COMP, MES_PAGTO, ANO_PAGTO) em DESPESAS
    Dim ultD As Long, r As Long
    ultD = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    For r = 2 To ultD
        If Trim(wsD.Cells(r, "A").Text) <> "" Then
            If IsDate(wsD.Cells(r, "B").Value) Then
                wsD.Cells(r, "L").Value = Month(CDate(wsD.Cells(r, "B").Value))
                wsD.Cells(r, "M").Value = Year(CDate(wsD.Cells(r, "B").Value))
            End If
            If IsDate(wsD.Cells(r, "D").Value) And UCase(Trim(wsD.Cells(r, "J").Text)) = "PAGA" Then
                wsD.Cells(r, "N").Value = Month(CDate(wsD.Cells(r, "D").Value))
                wsD.Cells(r, "O").Value = Year(CDate(wsD.Cells(r, "D").Value))
            Else
                wsD.Cells(r, "N").Value = ""
                wsD.Cells(r, "O").Value = ""
            End If
        End If
    Next r
    
    ' 3. Mapear Produtos/Serviços -> Tipo e Preço de Custo
    Dim dictTipo As Object, dictCusto As Object
    Set dictTipo = CreateObject("Scripting.Dictionary")
    Set dictCusto = CreateObject("Scripting.Dictionary")
    
    Dim ultP As Long, idP As Long
    ultP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    For r = 2 To ultP
        idP = Val(wsP.Cells(r, "A").Value)
        If idP > 0 Then
            dictTipo(idP) = Trim(wsP.Cells(r, "B").Text)
            dictCusto(idP) = ModFuncoes.ConverterParaDouble(wsP.Cells(r, "D").Value)
        End If
    Next r
    
    ' 4. Mapear Vendas Concluídas no Ano Base
    Dim dictVendaMes As Object
    Set dictVendaMes = CreateObject("Scripting.Dictionary")
    
    Dim vProd(1 To 12) As Double, vServ(1 To 12) As Double
    Dim vDesc(1 To 12) As Double, vTaxa(1 To 12) As Double
    Dim vCMV(1 To 12) As Double, vCSP(1 To 12) As Double
    
    Dim ultV As Long, idV As Long, dtV As Date, mesV As Long, anoV As Long
    Dim stV As String, totLiq As Double, valRec As Double, descV As Double, taxaRS As Double
    ultV = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To ultV
        idV = Val(wsV.Cells(r, "A").Value)
        stV = Trim(wsV.Cells(r, "I").Text)
        If idV > 0 And Left(UCase(stV), 6) = "CONCLU" Then
            If IsDate(wsV.Cells(r, "B").Value) Then
                dtV = CDate(wsV.Cells(r, "B").Value)
                mesV = Month(dtV)
                anoV = Year(dtV)
            Else
                mesV = Month(Date)
                anoV = anoBase
            End If
            
            If anoV = anoBase And mesV >= 1 And mesV <= 12 Then
                dictVendaMes(idV) = mesV
                descV = ModFuncoes.ConverterParaDouble(wsV.Cells(r, "F").Value)
                totLiq = ModFuncoes.ConverterParaDouble(wsV.Cells(r, "G").Value)
                valRec = ModFuncoes.ConverterParaDouble(wsV.Cells(r, "L").Value)
                If valRec <= 0 Then valRec = totLiq
                taxaRS = totLiq - valRec
                If taxaRS < 0 Then taxaRS = 0#
                
                vDesc(mesV) = vDesc(mesV) + descV
                vTaxa(mesV) = vTaxa(mesV) + Round(taxaRS, 2)
            End If
        End If
    Next r
    
    ' 5. Apurar Receita Bruta (Produtos vs Serviços) e CMV/CSP em ITENS_VENDA
    Dim ultI As Long, idVendaItem As Long, idProdItem As Long
    Dim qtdItem As Double, subItem As Double, custoUnit As Double, tipoItem As String
    ultI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To ultI
        idVendaItem = Val(wsI.Cells(r, "B").Value)
        If dictVendaMes.Exists(idVendaItem) Then
            mesV = dictVendaMes(idVendaItem)
            idProdItem = Val(wsI.Cells(r, "C").Value)
            qtdItem = ModFuncoes.ConverterParaDouble(wsI.Cells(r, "F").Value)
            subItem = ModFuncoes.ConverterParaDouble(wsI.Cells(r, "G").Value)
            
            If dictTipo.Exists(idProdItem) Then
                tipoItem = dictTipo(idProdItem)
                custoUnit = dictCusto(idProdItem)
            Else
                tipoItem = "Produto"
                If qtdItem > 0 Then custoUnit = Round((subItem / qtdItem) * 0.6, 2) Else custoUnit = 0#
            End If
            
            If Left(UCase(tipoItem), 4) = "SERV" Then
                vServ(mesV) = vServ(mesV) + subItem
                vCSP(mesV) = vCSP(mesV) + Round(qtdItem * custoUnit, 2)
            Else
                vProd(mesV) = vProd(mesV) + subItem
                vCMV(mesV) = vCMV(mesV) + Round(qtdItem * custoUnit, 2)
            End If
        End If
    Next r
    
    ' 6. Atualizar linhas dinâmicas na aba APURACAO_RESULTADOS e FLUXO_CAIXA
    Dim colsMes As Variant
    colsMes = Array("D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O")
    
    Dim m As Long, colL As String
    Dim sCMV As String, sCSP As String, sTax As String, sTaxFC As String
    
    For m = 1 To 12
        colL = colsMes(m - 1)
        wsDRE.Range(colL & "13").Value = Round(vProd(m), 2)
        wsDRE.Range(colL & "14").Value = Round(vServ(m), 2)
        wsDRE.Range(colL & "18").Value = Round(vDesc(m), 2)
        
        sCMV = Replace(CStr(Round(vCMV(m), 2)), ",", ".")
        sCSP = Replace(CStr(Round(vCSP(m), 2)), ",", ".")
        sTax = Replace(CStr(Round(vTaxa(m), 2)), ",", ".")
        
        wsDRE.Range(colL & "21").Formula = "=" & sCMV & "+SUMIFS(DESPESAS!$H:$H,DESPESAS!$E:$E,""4.1 (-) Custo das Mercadorias Vendidas"",DESPESAS!$L:$L," & m & ",DESPESAS!$M:$M,$E$5,DESPESAS!$J:$J,""<>Cancelada"")"
        wsDRE.Range(colL & "22").Formula = "=" & sCSP & "+SUMIFS(DESPESAS!$H:$H,DESPESAS!$E:$E,""4.2 (-) Custo dos Serviços Prestados"",DESPESAS!$L:$L," & m & ",DESPESAS!$M:$M,$E$5,DESPESAS!$J:$J,""<>Cancelada"")"
        wsDRE.Range(colL & "31").Formula = "=" & sTax & "+SUMIFS(DESPESAS!$H:$H,DESPESAS!$E:$E,""8.1 Despesas Financeiras"",DESPESAS!$L:$L," & m & ",DESPESAS!$M:$M,$E$5,DESPESAS!$J:$J,""<>Cancelada"")"
        
        wsFC.Range(colL & "27").Formula = "=" & sTax & "+SUMIFS(DESPESAS!$H:$H,DESPESAS!$F:$F,""Gastos com juros e taxas bancos"",DESPESAS!$N:$N," & m & ",DESPESAS!$O:$O,$E$5,DESPESAS!$J:$J,""Paga"")"
    Next m
    
    Application.Calculate
    
FimAtualizacao:
    Application.ScreenUpdating = True
End Sub

' -----------------------------------------------------------------------------
' ROTINAS DO PAINEL VBA: FormDespesas
' -----------------------------------------------------------------------------
Public Sub InicializarFormDespesas(ByVal frm As Object)
    On Error Resume Next
    With frm
        ' Preencher Categorias do Fluxo de Caixa (19 categorias oficiais)
        .CboCategoriaFluxo.Clear
        .CboCategoriaFluxo.AddItem "Aluguel da empresa"
        .CboCategoriaFluxo.AddItem "Energia elétrica da empresa"
        .CboCategoriaFluxo.AddItem "Água da empresa"
        .CboCategoriaFluxo.AddItem "Telefone da empresa"
        .CboCategoriaFluxo.AddItem "Internet da empresa"
        .CboCategoriaFluxo.AddItem "Folha de pagamento"
        .CboCategoriaFluxo.AddItem "Retirada do(s) proprietário(s)"
        .CboCategoriaFluxo.AddItem "DAS - Documento de Arrecadação Simples"
        .CboCategoriaFluxo.AddItem "Carnê de arrecadação - MEI"
        .CboCategoriaFluxo.AddItem "Outros tributos (impostos, taxas, etc.)"
        .CboCategoriaFluxo.AddItem "Gastos com produtos para venda"
        .CboCategoriaFluxo.AddItem "Gastos com serviços prestados"
        .CboCategoriaFluxo.AddItem "Gastos com juros e taxas bancos"
        .CboCategoriaFluxo.AddItem "Serviços contábeis"
        .CboCategoriaFluxo.AddItem "Manutenção da empresa"
        .CboCategoriaFluxo.AddItem "Manutenção do veículo da empresa"
        .CboCategoriaFluxo.AddItem "Combustível da empresa"
        .CboCategoriaFluxo.AddItem "Segurança da empresa"
        .CboCategoriaFluxo.AddItem "Despesas diversas da empresa"
        
        ' Preencher Grupos da DRE
        .CboGrupoDRE.Clear
        .CboGrupoDRE.AddItem "6.2 Despesas Administrativas"
        .CboGrupoDRE.AddItem "6.1 Despesas com Vendas"
        .CboGrupoDRE.AddItem "2.2 (-) Impostos e Contribuições sobre vendas"
        .CboGrupoDRE.AddItem "2.1 (-) Devoluções"
        .CboGrupoDRE.AddItem "4.1 (-) Custo das Mercadorias Vendidas"
        .CboGrupoDRE.AddItem "4.2 (-) Custo dos Serviços Prestados"
        .CboGrupoDRE.AddItem "8.1 Despesas Financeiras"
        .CboGrupoDRE.AddItem "11.1 Serviços de contabilidade"
        .CboGrupoDRE.AddItem "11.2 Manutenção de veículos"
        .CboGrupoDRE.AddItem "11.3 Manutenção da empresa"
        .CboGrupoDRE.AddItem "11.4 Seguro predial e instalação"
        .CboGrupoDRE.AddItem "11.5 Seguro produtos"
        .CboGrupoDRE.AddItem "11.6 Depreciação"
        .CboGrupoDRE.AddItem "13.1 Imposto de Renda"
        .CboGrupoDRE.AddItem "13.2 Contribuição Social sobre o Lucro Líquido"
        
        ' Formas de Pagamento
        .CboFormaPagto.Clear
        .CboFormaPagto.AddItem "PIX"
        .CboFormaPagto.AddItem "Boleto"
        .CboFormaPagto.AddItem "Dinheiro"
        .CboFormaPagto.AddItem "Transferência"
        .CboFormaPagto.AddItem "Cartão de Débito"
        .CboFormaPagto.AddItem "Cartão de Crédito"
        .CboFormaPagto.AddItem "Débito Automático"
        
        ' Status
        .CboStatus.Clear
        .CboStatus.AddItem "Paga"
        .CboStatus.AddItem "Pendente"
        .CboStatus.AddItem "Cancelada"
        
        ' Filtros de Mês, Ano e Status
        .CboFiltroMes.Clear
        .CboFiltroMes.AddItem "TODOS"
        .CboFiltroMes.AddItem "01 - Janeiro"
        .CboFiltroMes.AddItem "02 - Fevereiro"
        .CboFiltroMes.AddItem "03 - Março"
        .CboFiltroMes.AddItem "04 - Abril"
        .CboFiltroMes.AddItem "05 - Maio"
        .CboFiltroMes.AddItem "06 - Junho"
        .CboFiltroMes.AddItem "07 - Julho"
        .CboFiltroMes.AddItem "08 - Agosto"
        .CboFiltroMes.AddItem "09 - Setembro"
        .CboFiltroMes.AddItem "10 - Outubro"
        .CboFiltroMes.AddItem "11 - Novembro"
        .CboFiltroMes.AddItem "12 - Dezembro"
        .CboFiltroMes.Text = "TODOS"
        
        .CboFiltroAno.Clear
        .CboFiltroAno.AddItem "TODOS"
        .CboFiltroAno.AddItem "2025"
        .CboFiltroAno.AddItem "2026"
        .CboFiltroAno.AddItem "2027"
        .CboFiltroAno.Text = "2026"
        
        .CboFiltroStatus.Clear
        .CboFiltroStatus.AddItem "TODOS"
        .CboFiltroStatus.AddItem "Paga"
        .CboFiltroStatus.AddItem "Pendente"
        .CboFiltroStatus.AddItem "Cancelada"
        .CboFiltroStatus.Text = "TODOS"
        
        ' Configurar ListBox (9 colunas)
        .LstDespesas.ColumnCount = 9
        .LstDespesas.ColumnWidths = "35 pt;62 pt;62 pt;62 pt;125 pt;135 pt;68 pt;62 pt;52 pt"
    End With
    
    Call LimparFormDespesas(frm)
    Call CarregarListaDespesas(frm)
    On Error GoTo 0
End Sub

Public Sub LimparFormDespesas(ByVal frm As Object)
    On Error Resume Next
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    Dim proxID As Long
    proxID = 1
    If Not wsD Is Nothing Then
        Dim ult As Long, r As Long, idAt As Long
        ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
        For r = 2 To ult
            idAt = Val(wsD.Cells(r, "A").Value)
            If idAt >= proxID Then proxID = idAt + 1
        Next r
    End If
    
    With frm
        .TxtID.Text = CStr(proxID)
        .TxtDataComp.Text = Format(Date, "dd/mm/yyyy")
        .TxtDataVenc.Text = Format(Date, "dd/mm/yyyy")
        .TxtDataPagto.Text = Format(Date, "dd/mm/yyyy")
        .CboStatus.Text = "Paga"
        .CboFormaPagto.Text = "PIX"
        .CboCategoriaFluxo.Text = "Aluguel da empresa"
        .CboGrupoDRE.Text = "6.2 Despesas Administrativas"
        .TxtDescricao.Text = ""
        .TxtValor.Text = "0,00"
        .LstDespesas.ListIndex = -1
    End With
    On Error GoTo 0
End Sub

Public Sub AoMudarCategoriaFluxo(ByVal frm As Object)
    On Error Resume Next
    Dim cat As String
    cat = Trim(frm.CboCategoriaFluxo.Text)
    If cat <> "" Then
        frm.CboGrupoDRE.Text = MapearGrupoDREPorCategoria(cat)
    End If
    On Error GoTo 0
End Sub

Public Sub AoMudarStatusDespesa(ByVal frm As Object)
    On Error Resume Next
    If UCase(Trim(frm.CboStatus.Text)) = "PAGA" Then
        If Trim(frm.TxtDataPagto.Text) = "" Then
            frm.TxtDataPagto.Text = Format(Date, "dd/mm/yyyy")
        End If
    ElseIf UCase(Trim(frm.CboStatus.Text)) = "PENDENTE" Then
        frm.TxtDataPagto.Text = ""
    End If
    On Error GoTo 0
End Sub

Public Sub CarregarListaDespesas(ByVal frm As Object)
    On Error GoTo TrataErroLista
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    If wsD Is Nothing Then Exit Sub
    
    frm.LstDespesas.Clear
    
    Dim fMes As Long, fAno As Long, fStatus As String
    fMes = 0
    If Left(Trim(frm.CboFiltroMes.Text), 5) <> "TODOS" And Trim(frm.CboFiltroMes.Text) <> "" Then
        fMes = Val(Left(Trim(frm.CboFiltroMes.Text), 2))
    End If
    
    fAno = 0
    If UCase(Trim(frm.CboFiltroAno.Text)) <> "TODOS" And Trim(frm.CboFiltroAno.Text) <> "" Then
        fAno = Val(frm.CboFiltroAno.Text)
    End If
    
    fStatus = UCase(Trim(frm.CboFiltroStatus.Text))
    If fStatus = "" Then fStatus = "TODOS"
    
    Dim totMes As Double, totPagas As Double, totPend As Double
    totMes = 0#: totPagas = 0#: totPend = 0#
    
    Dim ult As Long, r As Long, idx As Long
    ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    idx = 0
    
    For r = 2 To ult
        If Trim(wsD.Cells(r, "A").Text) <> "" Then
            Dim mReg As Long, aReg As Long, stReg As String, vlrReg As Double
            If IsDate(wsD.Cells(r, "B").Value) Then
                mReg = Month(CDate(wsD.Cells(r, "B").Value))
                aReg = Year(CDate(wsD.Cells(r, "B").Value))
            Else
                mReg = Val(wsD.Cells(r, "L").Value)
                aReg = Val(wsD.Cells(r, "M").Value)
            End If
            stReg = Trim(wsD.Cells(r, "J").Text)
            vlrReg = ModFuncoes.ConverterParaDouble(wsD.Cells(r, "H").Value)
            
            Dim passaFiltro As Boolean
            passaFiltro = True
            If fMes > 0 And mReg <> fMes Then passaFiltro = False
            If fAno > 0 And aReg <> fAno Then passaFiltro = False
            If fStatus <> "TODOS" And UCase(stReg) <> fStatus Then passaFiltro = False
            
            If passaFiltro Then
                frm.LstDespesas.AddItem CStr(wsD.Cells(r, "A").Value)
                frm.LstDespesas.List(idx, 1) = Format(wsD.Cells(r, "B").Value, "dd/mm/yyyy")
                frm.LstDespesas.List(idx, 2) = Format(wsD.Cells(r, "C").Value, "dd/mm/yyyy")
                If IsDate(wsD.Cells(r, "D").Value) Then
                    frm.LstDespesas.List(idx, 3) = Format(wsD.Cells(r, "D").Value, "dd/mm/yyyy")
                Else
                    frm.LstDespesas.List(idx, 3) = "-"
                End If
                frm.LstDespesas.List(idx, 4) = Trim(wsD.Cells(r, "F").Text)
                frm.LstDespesas.List(idx, 5) = Trim(wsD.Cells(r, "G").Text)
                frm.LstDespesas.List(idx, 6) = Format(vlrReg, "#,##0.00")
                frm.LstDespesas.List(idx, 7) = Trim(wsD.Cells(r, "I").Text)
                frm.LstDespesas.List(idx, 8) = stReg
                idx = idx + 1
                
                If UCase(stReg) <> "CANCELADA" Then
                    totMes = totMes + vlrReg
                    If UCase(stReg) = "PAGA" Then
                        totPagas = totPagas + vlrReg
                    ElseIf UCase(stReg) = "PENDENTE" Then
                        totPend = totPend + vlrReg
                    End If
                End If
            End If
        End If
    Next r
    
    frm.LblResumoTotal.Caption = "TOTAL: R$ " & Format(totMes, "#,##0.00")
    frm.LblResumoPagas.Caption = "PAGAS: R$ " & Format(totPagas, "#,##0.00")
    frm.LblResumoPendentes.Caption = "PENDENTES: R$ " & Format(totPend, "#,##0.00")
    Exit Sub
    
TrataErroLista:
End Sub

Public Sub SelecionarItemListaDespesas(ByVal frm As Object)
    On Error Resume Next
    Dim idx As Long
    idx = frm.LstDespesas.ListIndex
    If idx < 0 Then Exit Sub
    
    Dim idSel As Long
    idSel = Val(frm.LstDespesas.List(idx, 0))
    If idSel <= 0 Then Exit Sub
    
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    Dim ult As Long, r As Long
    ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To ult
        If Val(wsD.Cells(r, "A").Value) = idSel Then
            frm.TxtID.Text = CStr(idSel)
            frm.TxtDataComp.Text = Format(wsD.Cells(r, "B").Value, "dd/mm/yyyy")
            frm.TxtDataVenc.Text = Format(wsD.Cells(r, "C").Value, "dd/mm/yyyy")
            If IsDate(wsD.Cells(r, "D").Value) Then
                frm.TxtDataPagto.Text = Format(wsD.Cells(r, "D").Value, "dd/mm/yyyy")
            Else
                frm.TxtDataPagto.Text = ""
            End If
            frm.CboGrupoDRE.Text = Trim(wsD.Cells(r, "E").Text)
            frm.CboCategoriaFluxo.Text = Trim(wsD.Cells(r, "F").Text)
            frm.TxtDescricao.Text = Trim(wsD.Cells(r, "G").Text)
            frm.TxtValor.Text = Format(ModFuncoes.ConverterParaDouble(wsD.Cells(r, "H").Value), "#,##0.00")
            frm.CboFormaPagto.Text = Trim(wsD.Cells(r, "I").Text)
            frm.CboStatus.Text = Trim(wsD.Cells(r, "J").Text)
            Exit For
        End If
    Next r
    On Error GoTo 0
End Sub

Public Sub SalvarDespesaForm(ByVal frm As Object)
    On Error GoTo TrataErroSalvar
    Dim vlr As Double
    vlr = ModFuncoes.ConverterParaDouble(frm.TxtValor.Text)
    If vlr <= 0 Then
        MsgBox "Informe um valor de despesa válido maior que zero!", vbExclamation, "Valor Inválido"
        frm.TxtValor.SetFocus
        Exit Sub
    End If
    
    If Trim(frm.CboCategoriaFluxo.Text) = "" Then
        MsgBox "Selecione a Categoria da Despesa (Fluxo de Caixa)!", vbExclamation, "Categoria Obrigatória"
        Exit Sub
    End If
    
    If Trim(frm.CboGrupoDRE.Text) = "" Then
        frm.CboGrupoDRE.Text = MapearGrupoDREPorCategoria(frm.CboCategoriaFluxo.Text)
    End If
    
    Dim dtComp As Date, dtVenc As Date
    If IsDate(frm.TxtDataComp.Text) Then dtComp = CDate(frm.TxtDataComp.Text) Else dtComp = Date
    If IsDate(frm.TxtDataVenc.Text) Then dtVenc = CDate(frm.TxtDataVenc.Text) Else dtVenc = dtComp
    
    Dim st As String
    st = Trim(frm.CboStatus.Text)
    If st = "" Then st = "Paga"
    
    Dim temPagto As Boolean, dtPagto As Date
    temPagto = False
    If UCase(st) = "PAGA" Then
        temPagto = True
        If IsDate(frm.TxtDataPagto.Text) Then dtPagto = CDate(frm.TxtDataPagto.Text) Else dtPagto = Date
    ElseIf IsDate(frm.TxtDataPagto.Text) And UCase(st) <> "PENDENTE" Then
        temPagto = True
        dtPagto = CDate(frm.TxtDataPagto.Text)
    End If
    
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    Dim idDesp As Long
    idDesp = Val(frm.TxtID.Text)
    If idDesp <= 0 Then idDesp = 1
    
    Dim ult As Long, r As Long, linGravar As Long
    ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    linGravar = 0
    For r = 2 To ult
        If Val(wsD.Cells(r, "A").Value) = idDesp Then
            linGravar = r
            Exit For
        End If
    Next r
    If linGravar = 0 Then linGravar = ult + 1
    
    Dim opNome As String
    opNome = Trim(ModSeguranca.vNomeOperador)
    If opNome = "" Then opNome = "Administrador"
    
    wsD.Cells(linGravar, "A").Value = idDesp
    wsD.Cells(linGravar, "B").Value = dtComp
    wsD.Cells(linGravar, "C").Value = dtVenc
    If temPagto Then
        wsD.Cells(linGravar, "D").Value = dtPagto
        wsD.Cells(linGravar, "N").Value = Month(dtPagto)
        wsD.Cells(linGravar, "O").Value = Year(dtPagto)
    Else
        wsD.Cells(linGravar, "D").Value = ""
        wsD.Cells(linGravar, "N").Value = ""
        wsD.Cells(linGravar, "O").Value = ""
    End If
    wsD.Cells(linGravar, "E").Value = Trim(frm.CboGrupoDRE.Text)
    wsD.Cells(linGravar, "F").Value = Trim(frm.CboCategoriaFluxo.Text)
    wsD.Cells(linGravar, "G").Value = IIf(Trim(frm.TxtDescricao.Text) = "", Trim(frm.CboCategoriaFluxo.Text), Trim(frm.TxtDescricao.Text))
    wsD.Cells(linGravar, "H").Value = Round(vlr, 2)
    wsD.Cells(linGravar, "I").Value = IIf(Trim(frm.CboFormaPagto.Text) = "", "PIX", Trim(frm.CboFormaPagto.Text))
    wsD.Cells(linGravar, "J").Value = st
    wsD.Cells(linGravar, "K").Value = opNome
    wsD.Cells(linGravar, "L").Value = Month(dtComp)
    wsD.Cells(linGravar, "M").Value = Year(dtComp)
    
    Call AtualizarDemonstrativos
    Call CarregarListaDespesas(frm)
    Call LimparFormDespesas(frm)
    
    MsgBox "Despesa Nº " & idDesp & " salva e integrada à DRE e ao Fluxo de Caixa com sucesso!", vbInformation, "Despesa Registrada"
    Exit Sub
    
TrataErroSalvar:
    MsgBox "Erro ao salvar despesa: " & Err.Description, vbCritical, "Erro"
End Sub

Public Sub DarBaixaDespesaForm(ByVal frm As Object)
    On Error GoTo TrataErroBaixa
    Dim idx As Long
    idx = frm.LstDespesas.ListIndex
    If idx < 0 Then
        MsgBox "Selecione uma despesa na lista abaixo para dar baixa (marcar como Paga)!", vbExclamation, "Seleção Necessária"
        Exit Sub
    End If
    
    Dim idSel As Long
    idSel = Val(frm.LstDespesas.List(idx, 0))
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    Dim ult As Long, r As Long
    ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To ult
        If Val(wsD.Cells(r, "A").Value) = idSel Then
            wsD.Cells(r, "D").Value = Date
            wsD.Cells(r, "J").Value = "Paga"
            wsD.Cells(r, "N").Value = Month(Date)
            wsD.Cells(r, "O").Value = Year(Date)
            Exit For
        End If
    Next r
    
    Call AtualizarDemonstrativos
    Call CarregarListaDespesas(frm)
    Call LimparFormDespesas(frm)
    
    MsgBox "Baixa realizada com sucesso! A despesa Nº " & idSel & " foi marcada como PAGA e lançada no Fluxo de Caixa.", vbInformation, "Baixa Confirmada"
    Exit Sub
    
TrataErroBaixa:
    MsgBox "Erro ao dar baixa: " & Err.Description, vbCritical, "Erro"
End Sub

Public Sub ExcluirDespesaForm(ByVal frm As Object)
    On Error GoTo TrataErroExcluir
    Dim idx As Long
    idx = frm.LstDespesas.ListIndex
    If idx < 0 Then
        MsgBox "Selecione uma despesa na lista para excluir!", vbExclamation, "Seleção Necessária"
        Exit Sub
    End If
    
    Dim idSel As Long
    idSel = Val(frm.LstDespesas.List(idx, 0))
    If MsgBox("Confirma a exclusão definitiva da Despesa Nº " & idSel & "?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbNo Then
        Exit Sub
    End If
    
    Dim wsD As Worksheet
    Set wsD = PlanDespesas()
    Dim ult As Long, r As Long
    ult = wsD.Cells(wsD.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To ult
        If Val(wsD.Cells(r, "A").Value) = idSel Then
            wsD.Rows(r).Delete
            Exit For
        End If
    Next r
    
    Call AtualizarDemonstrativos
    Call CarregarListaDespesas(frm)
    Call LimparFormDespesas(frm)
    
    MsgBox "Despesa Nº " & idSel & " excluída e demonstrativos recalculados!", vbInformation, "Exclusão Concluída"
    Exit Sub
    
TrataErroExcluir:
    MsgBox "Erro ao excluir despesa: " & Err.Description, vbCritical, "Erro"
End Sub

