Attribute VB_Name = "ModProdutosServicos"
Option Explicit
Option Compare Text

Public vIDProduto As Variant
Public bSincronizandoPlanilha As Boolean

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
    On Error Resume Next
    wsP.Unprotect "VendasSeguras2026"
    On Error GoTo 0
    
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
        precoFinalStr = Replace(Replace(Replace(.TxtPreco.Text, "R$", ""), " ", ""), ",", ".")
        precoCustoStr = Replace(Replace(Replace(.TxtPrecoCusto.Text, "R$", ""), " ", ""), ",", ".")
        margemStr = Replace(Replace(Replace(.TxtMargem.Text, "%", ""), " ", ""), ",", ".")
        
        Dim valPrecoFinal As Double, valPrecoCusto As Double, valMargem As Double
        valPrecoFinal = Val(precoFinalStr)
        valPrecoCusto = Val(precoCustoStr)
        valMargem = Val(margemStr)
        
        If valPrecoFinal <= 0 And valPrecoCusto > 0 Then
            valPrecoFinal = Round(valPrecoCusto * (1 + (valMargem / 100)), 2)
        ElseIf valPrecoCusto > 0 And valMargem = 0 And valPrecoFinal > 0 Then
            valMargem = Round(((valPrecoFinal / valPrecoCusto) - 1) * 100, 1)
        End If
        
        If valPrecoFinal <= 0 Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Informe o Preço de Custo/Margem ou Preço Final de Venda!", vbExclamation, "Preço Inválido"
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
        
        bSincronizandoPlanilha = True
        Dim cel As Range
        If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
            Set cel = wsP.Range("C:C").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                bSincronizandoPlanilha = False
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
            wsP.Cells(lin, "D").NumberFormat = "R$ #,##0.00"
            wsP.Cells(lin, "E").Value = valMargem
            wsP.Cells(lin, "E").NumberFormat = "0.0""%"""
            If valPrecoCusto > 0 Then
                wsP.Cells(lin, "F").Formula = "=ROUND(D" & lin & "*(1+(E" & lin & "/100)),2)"
            Else
                wsP.Cells(lin, "F").Value = valPrecoFinal
            End If
            wsP.Cells(lin, "F").NumberFormat = "R$ #,##0.00"
            wsP.Cells(lin, "G").Value = estAtual
            wsP.Cells(lin, "H").Value = estMin
            
            bSincronizandoPlanilha = False
            MsgBox "Item cadastrado com sucesso!", vbInformation, "Cadastro Concluído"
        Else
            Set cel = wsP.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
            If cel Is Nothing Then
                bSincronizandoPlanilha = False
                MsgBox "Item não localizado para alteração!", vbCritical, "Erro"
                Exit Sub
            End If
            
            Dim rEd As Long
            rEd = cel.Row
            wsP.Cells(rEd, "B").Value = tipoItem
            wsP.Cells(rEd, "C").Value = Trim(.TxtDescricao.Text)
            wsP.Cells(rEd, "D").Value = valPrecoCusto
            wsP.Cells(rEd, "D").NumberFormat = "R$ #,##0.00"
            wsP.Cells(rEd, "E").Value = valMargem
            wsP.Cells(rEd, "E").NumberFormat = "0.0""%"""
            If valPrecoCusto > 0 Then
                wsP.Cells(rEd, "F").Formula = "=ROUND(D" & rEd & "*(1+(E" & rEd & "/100)),2)"
            Else
                wsP.Cells(rEd, "F").Value = valPrecoFinal
            End If
            wsP.Cells(rEd, "F").NumberFormat = "R$ #,##0.00"
            wsP.Cells(rEd, "G").Value = estAtual
            wsP.Cells(rEd, "H").Value = estMin
            
            bSincronizandoPlanilha = False
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
        On Error Resume Next
        wsP.Unprotect "VendasSeguras2026"
        On Error GoTo 0
        bSincronizandoPlanilha = True
        cel.EntireRow.Delete
        bSincronizandoPlanilha = False
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
    
    bSincronizandoPlanilha = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    wsInv.Unprotect "VendasSeguras2026"
    wsP.Unprotect "VendasSeguras2026"
    
    Dim ultP As Long
    ultP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    
    Dim ultInv As Long
    ultInv = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row
    If ultInv >= 12 Then
        wsInv.Range("B12:M" & ultInv + 5).Clear
    End If
    
    If ultP >= 2 Then
        Dim qtdItens As Long, linFim As Long
        qtdItens = ultP - 1
        linFim = 11 + qtdItens
        
        ' Copiar dados base (ID, Tipo, Descricao, Custo, Margem) em bloco
        wsInv.Range("B12:F" & linFim).Value = wsP.Range("A2:E" & ultP).Value
        ' Copiar Estoque Atual e Estoque Minimo
        wsInv.Range("H12:I" & linFim).Value = wsP.Range("G2:H" & ultP).Value
        
        ' Garantir formulas dinamicas em PRODUTOS_SERVICOS (Coluna F = Preco Venda)
        wsP.Range("F2:F" & ultP).Formula = "=ROUND(D2*(1+(E2/100)),2)"
        
        ' Inserir FORMULAS DINAMICAS AUTOMATIVAS em INVENTARIO:
        ' Col G (Preco Venda): aceita Margem tanto em formato 50 (50%) quanto 0.5 (50%)
        wsInv.Range("G12:G" & linFim).Formula = "=ROUND(E12*(1+IF(ABS(F12)<=1,F12,F12/100)),2)"
        
        ' Col J (Status Automatico do Estoque)
        wsInv.Range("J12:J" & linFim).Formula = "=IF(OR(UPPER(C12)=""SERVIÇO"",UPPER(C12)=""SERVICO""),""SERVIÇO"",IF(H12<=0,""ZERADO"",IF(H12<=I12,""ESTOQUE BAIXO"",""NORMAL"")))"
        
        ' Col K (Custo Total = Estoque * Custo Unit)
        wsInv.Range("K12:K" & linFim).Formula = "=IF(OR(UPPER(C12)=""SERVIÇO"",UPPER(C12)=""SERVICO""),0,N(H12)*N(E12))"
        
        ' Col L (Venda Total = Estoque * Preco Venda)
        wsInv.Range("L12:L" & linFim).Formula = "=IF(OR(UPPER(C12)=""SERVIÇO"",UPPER(C12)=""SERVICO""),0,N(H12)*N(G12))"
        
        ' Col M (Lucro Projetado = Venda Total - Custo Total)
        wsInv.Range("M12:M" & linFim).Formula = "=IF(OR(UPPER(C12)=""SERVIÇO"",UPPER(C12)=""SERVICO""),0,L12-K12)"
        
        ' Formatacoes numericas elegantes
        wsInv.Range("E12:E" & linFim).NumberFormat = "R$ #,##0.00"
        wsInv.Range("F12:F" & linFim).NumberFormat = "0.0""%"""
        wsInv.Range("G12:G" & linFim).NumberFormat = "R$ #,##0.00"
        wsInv.Range("K12:M" & linFim).NumberFormat = "R$ #,##0.00"
        wsInv.Range("B12:C" & linFim).HorizontalAlignment = xlCenter
        wsInv.Range("H12:J" & linFim).HorizontalAlignment = xlCenter
        wsInv.Range("B12:M" & linFim).Font.Name = "Segoe UI"
        wsInv.Range("B12:M" & linFim).Font.Size = 9.5
        
        ' Cards de KPIs no Topo (Linha 6) com FORMULAS DINAMICAS EM TEMPO REAL
        wsInv.Range("B6").Formula = "=SUM(H12:H" & linFim & ")"
        wsInv.Range("B6").NumberFormat = "#,##0"" un."""
        
        wsInv.Range("D6").Formula = "=SUM(K12:K" & linFim & ")"
        wsInv.Range("D6").NumberFormat = "R$ #,##0.00"
        
        wsInv.Range("F6").Formula = "=SUM(L12:L" & linFim & ")"
        wsInv.Range("F6").NumberFormat = "R$ #,##0.00"
        
        wsInv.Range("H6").Formula = "=SUM(M12:M" & linFim & ")"
        wsInv.Range("H6").NumberFormat = "R$ #,##0.00"
        
        wsInv.Range("J6").Formula = "=COUNTIF(J12:J" & linFim & ",""ZERADO"")+COUNTIF(J12:J" & linFim & ",""ESTOQUE BAIXO"")"
        wsInv.Range("J6").NumberFormat = "#,##0"" alertas"""
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    bSincronizandoPlanilha = False
    
    Call AplicarProtecaoInventario
    On Error GoTo 0
End Sub

Public Sub ProcessarEdicaoPlanilha(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo FimEdicao
    If bSincronizandoPlanilha Then Exit Sub
    If Target Is Nothing Then Exit Sub
    
    Dim nomeAba As String
    nomeAba = UCase(Trim(Sh.Name))
    If nomeAba <> "INVENTARIO" And nomeAba <> "PRODUTOS_SERVICOS" Then Exit Sub
    
    ' Se for RECEPCAO, bloquear qualquer edicao manual
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        Call AplicarProtecaoInventario
        Exit Sub
    End If
    
    bSincronizandoPlanilha = True
    Dim wsP As Worksheet, wsInv As Worksheet
    Set wsP = PlanProdutosServicos()
    Set wsInv = PlanInventario()
    wsP.Unprotect "VendasSeguras2026"
    wsInv.Unprotect "VendasSeguras2026"
    
    Dim cel As Range
    For Each cel In Target.Cells
        Dim r As Long, col As Long
        r = cel.Row
        col = cel.Column
        
        If nomeAba = "INVENTARIO" And r >= 12 Then
            Dim idInv As Long
            idInv = Val(wsInv.Cells(r, "B").Value)
            If idInv > 0 Then
                Dim vCusto As Double, vMargem As Double, vVenda As Double
                vCusto = Val(Replace(wsInv.Cells(r, "E").Value, ",", "."))
                
                ' Se o usuario digitou diretamente um novo PRECO DE VENDA (Coluna G = 7)
                If col = 7 Then
                    vVenda = Val(Replace(wsInv.Cells(r, "G").Value, ",", "."))
                    If vCusto > 0 And vVenda > 0 Then
                        vMargem = Round(((vVenda / vCusto) - 1) * 100, 1)
                        wsInv.Cells(r, "F").Value = vMargem
                        wsInv.Cells(r, "F").NumberFormat = "0.0""%"""
                    End If
                    wsInv.Cells(r, "G").Formula = "=ROUND(E" & r & "*(1+IF(ABS(F" & r & ")<=1,F" & r & ",F" & r & "/100)),2)"
                ElseIf col = 5 Or col = 6 Then
                    ' Alterou Custo (E=5) ou Margem (F=6) -> garantir formula em G
                    If col = 6 Then
                        Dim mRaw As Double
                        mRaw = Val(Replace(wsInv.Cells(r, "F").Value, ",", "."))
                        If Abs(mRaw) > 0 And Abs(mRaw) <= 1 Then
                            wsInv.Cells(r, "F").Value = Round(mRaw * 100, 1)
                        End If
                        wsInv.Cells(r, "F").NumberFormat = "0.0""%"""
                    End If
                    wsInv.Cells(r, "G").Formula = "=ROUND(E" & r & "*(1+IF(ABS(F" & r & ")<=1,F" & r & ",F" & r & "/100)),2)"
                End If
                
                ' Sincronizar linha editada em INVENTARIO de volta para PRODUTOS_SERVICOS
                Dim celP As Range
                Set celP = wsP.Range("A:A").Find(What:=idInv, LookIn:=xlValues, LookAt:=xlWhole)
                If Not celP Is Nothing Then
                    Dim rP As Long
                    rP = celP.Row
                    wsP.Cells(rP, "B").Value = wsInv.Cells(r, "C").Value
                    wsP.Cells(rP, "C").Value = wsInv.Cells(r, "D").Value
                    wsP.Cells(rP, "D").Value = Val(Replace(wsInv.Cells(r, "E").Value, ",", "."))
                    Dim mSync As Double
                    mSync = Val(Replace(wsInv.Cells(r, "F").Value, ",", "."))
                    If Abs(mSync) > 0 And Abs(mSync) <= 1 Then mSync = Round(mSync * 100, 1)
                    wsP.Cells(rP, "E").Value = mSync
                    wsP.Cells(rP, "F").Formula = "=ROUND(D" & rP & "*(1+(E" & rP & "/100)),2)"
                    wsP.Cells(rP, "G").Value = Val(wsInv.Cells(r, "H").Value)
                    wsP.Cells(rP, "H").Value = Val(wsInv.Cells(r, "I").Value)
                End If
            End If
            
        ElseIf nomeAba = "PRODUTOS_SERVICOS" And r >= 2 Then
            Dim idP As Long
            idP = Val(wsP.Cells(r, "A").Value)
            If idP > 0 Then
                Dim cP As Double, mP As Double, vP As Double
                cP = Val(Replace(wsP.Cells(r, "D").Value, ",", "."))
                
                ' Se o usuario digitou diretamente o PRECO DE VENDA (Coluna F = 6)
                If col = 6 Then
                    vP = Val(Replace(wsP.Cells(r, "F").Value, ",", "."))
                    If cP > 0 And vP > 0 Then
                        mP = Round(((vP / cP) - 1) * 100, 1)
                        wsP.Cells(r, "E").Value = mP
                        wsP.Cells(r, "E").NumberFormat = "0.0""%"""
                    End If
                    wsP.Cells(r, "F").Formula = "=ROUND(D" & r & "*(1+(E" & r & "/100)),2)"
                ElseIf col = 4 Or col = 5 Then
                    ' Alterou Custo (D=4) ou Margem (E=5) -> garantir formula em F
                    If col = 5 Then
                        Dim mPRaw As Double
                        mPRaw = Val(Replace(wsP.Cells(r, "E").Value, ",", "."))
                        If Abs(mPRaw) > 0 And Abs(mPRaw) <= 1 Then
                            wsP.Cells(r, "E").Value = Round(mPRaw * 100, 1)
                        End If
                        wsP.Cells(r, "E").NumberFormat = "0.0""%"""
                    End If
                    wsP.Cells(r, "F").Formula = "=ROUND(D" & r & "*(1+(E" & r & "/100)),2)"
                End If
                
                ' Sincronizar com a linha correspondente em INVENTARIO
                Dim celInv As Range
                Set celInv = wsInv.Range("B:B").Find(What:=idP, LookIn:=xlValues, LookAt:=xlWhole)
                If Not celInv Is Nothing Then
                    Dim rI As Long
                    rI = celInv.Row
                    wsInv.Cells(rI, "C").Value = wsP.Cells(r, "B").Value
                    wsInv.Cells(rI, "D").Value = wsP.Cells(r, "C").Value
                    wsInv.Cells(rI, "E").Value = Val(Replace(wsP.Cells(r, "D").Value, ",", "."))
                    wsInv.Cells(rI, "F").Value = Val(Replace(wsP.Cells(r, "E").Value, ",", "."))
                    wsInv.Cells(rI, "F").NumberFormat = "0.0""%"""
                    wsInv.Cells(rI, "G").Formula = "=ROUND(E" & rI & "*(1+IF(ABS(F" & rI & ")<=1,F" & rI & ",F" & rI & "/100)),2)"
                    wsInv.Cells(rI, "H").Value = Val(wsP.Cells(r, "G").Value)
                    wsInv.Cells(rI, "I").Value = Val(wsP.Cells(r, "H").Value)
                End If
            End If
        End If
    Next cel

FimEdicao:
    bSincronizandoPlanilha = False
    Call AplicarProtecaoInventario
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
        ' Bloqueia celulas apenas quando o perfil ativo for RECEPCAO!
        ' Quando for ADMIN, deixa desbloqueado para edicao dinamica direta na planilha!
        If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
            wsInv.Protect Password:="VendasSeguras2026", UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, Scenarios:=True
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

Public Sub SalvarProdutos()
    Call SalvarProdutosServicos
End Sub

Public Sub BuscarProdutos()
    Call BuscarProdutosServicos
End Sub

Public Sub ExcluirProduto()
    Call ExcluirProdutoServico
End Sub


