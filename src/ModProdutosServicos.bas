Attribute VB_Name = "ModProdutosServicos"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE GESTÃO DE PRODUTOS E SERVIÇOS (FASE 7)
' =============================================================
Public vIDProduto As Variant

Public Function PlanProdutosServicos() As Worksheet
    On Error Resume Next
    Set PlanProdutosServicos = ThisWorkbook.Sheets("PRODUTOS_SERVICOS")
    If PlanProdutosServicos Is Nothing Then
        Set PlanProdutosServicos = ThisWorkbook.Sheets("PRODUTOS E SERVIÇOS")
    End If
    If PlanProdutosServicos Is Nothing Then
        Set PlanProdutosServicos = ThisWorkbook.Sheets("PRODUTOS")
    End If
    On Error GoTo 0
End Function

Public Function PlanProdutos() As Worksheet
    Set PlanProdutos = PlanProdutosServicos()
End Function

Public Sub SalvarProdutosServicos()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Salvar ou Alterar Preço de Produto/Serviço") Then Exit Sub
    End If
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    If wsP Is Nothing Then
        MsgBox "Planilha de Produtos e Serviços não encontrada!", vbCritical, "Erro de Banco de Dados"
        Exit Sub
    End If
    
    With FormProdutos
        .TxtDescricao.BorderColor = &H80000012
        .TxtPreco.BorderColor = &H80000012
        
        If Trim(.TxtDescricao.Text) = "" Then
            .TxtDescricao.BorderColor = &HFF&
            MsgBox "Digite o campo Descrição do Produto / Serviço!", vbExclamation, "Cadastro de Produtos e Serviços"
            .TxtDescricao.SetFocus
            Exit Sub
        End If
        
        If Trim(.TxtPreco.Text) = "" Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Digite o campo Preço!", vbExclamation, "Cadastro de Produtos e Serviços"
            .TxtPreco.SetFocus
            Exit Sub
        End If
        
        Dim precoFormatado As String
        precoFormatado = Replace(.TxtPreco.Text, "R$", "")
        precoFormatado = Trim(precoFormatado)
        
        If Not IsNumeric(precoFormatado) Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Informe um valor numérico válido para o preço!", vbExclamation, "Cadastro de Produtos e Serviços"
            .TxtPreco.SetFocus
            Exit Sub
        End If
        
        Dim valPreco As Double
        valPreco = VBA.CDbl(precoFormatado)
        
        Dim cel As Range
        
        If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
            Set cel = wsP.Range("B:B").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                .TxtDescricao.BorderColor = &HFF&
                MsgBox "Já existe um produto/serviço cadastrado com esta Descrição!", vbExclamation, "Item Duplicado"
                .TxtDescricao.SetFocus
                Exit Sub
            End If
            
            Dim lin As Long
            lin = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row + 1
            
            Dim novoID As Long
            If lin = 2 Then
                novoID = 1
            Else
                novoID = wsP.Cells(lin - 1, "A").Value + 1
            End If
            
            wsP.Cells(lin, "A").Value = novoID
            wsP.Cells(lin, "B").Value = Trim(.TxtDescricao.Text)
            wsP.Cells(lin, "C").Value = valPreco
            wsP.Cells(lin, "C").NumberFormat = "R$ #,##0.00"
            
            MsgBox "Produto / Serviço cadastrado com sucesso!", vbInformation, "Cadastro de Produtos e Serviços"
        Else
            Set cel = wsP.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
            
            If cel Is Nothing Then
                MsgBox "Produto/Serviço não localizado para alteração!", vbCritical, "Erro"
                Exit Sub
            End If
            
            Dim celDup As Range
            Set celDup = wsP.Range("B:B").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not celDup Is Nothing Then
                If celDup.Row <> cel.Row Then
                    .TxtDescricao.BorderColor = &HFF&
                    MsgBox "Já existe outro produto/serviço cadastrado com esta Descrição!", vbExclamation, "Item Duplicado"
                    .TxtDescricao.SetFocus
                    Exit Sub
                End If
            End If
            
            cel.Offset(0, 1).Value = Trim(.TxtDescricao.Text)
            cel.Offset(0, 2).Value = valPreco
            cel.Offset(0, 2).NumberFormat = "R$ #,##0.00"
            
            MsgBox "Produto / Serviço alterado com sucesso!", vbInformation, "Cadastro de Produtos e Serviços"
        End If
        
        Call LimparCampos
        Call BuscarProdutosServicos
    End With
End Sub

Public Sub ExcluirProdutoServico()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Excluir Produto/Serviço do Catálogo") Then Exit Sub
    End If
    
    If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
        MsgBox "Selecione um produto/serviço na lista para excluir!", vbExclamation, "Excluir Item"
        Exit Sub
    End If
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    
    Dim cel As Range
    Set cel = wsP.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cel Is Nothing Then
        MsgBox "Produto/Serviço não encontrado!", vbCritical, "Erro"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir o item: " & cel.Offset(0, 1).Text & "?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbYes Then
        cel.EntireRow.Delete
        MsgBox "Produto / Serviço excluído com sucesso!", vbInformation, "Exclusão Concluída"
        Call LimparCampos
        Call BuscarProdutosServicos
    End If
End Sub

Public Sub LimparCampos()
    With FormProdutos
        .TxtDescricao.Text = ""
        .TxtPreco.Text = ""
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
    Dim lin As Long, i As Long
    Dim ultLinha As Long
    i = 0
    
    FormProdutos.LstDados.Clear
    
    Dim wsP As Worksheet
    Set wsP = PlanProdutosServicos()
    If wsP Is Nothing Then Exit Sub
    
    ultLinha = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    If ultLinha < 2 Then Exit Sub
    
    For lin = 2 To ultLinha
        If wsP.Cells(lin, "B").Text Like "*" & ModFuncoes.AcSQL(FormProdutos.TxtPesquisa.Text) & "*" Then
            FormProdutos.LstDados.AddItem
            FormProdutos.LstDados.List(i, 0) = wsP.Cells(lin, "A").Text
            FormProdutos.LstDados.List(i, 1) = "-"
            FormProdutos.LstDados.List(i, 2) = wsP.Cells(lin, "B").Text
            FormProdutos.LstDados.List(i, 3) = Format(wsP.Cells(lin, "C").Value, "R$ #,##0.00")
            i = i + 1
        End If
    Next lin
End Sub

' === ALIASES PARA COMPATIBILIDADE TOTAL ===
Public Sub SalvarProdutos()
    Call SalvarProdutosServicos
End Sub

Public Sub BuscarProdutos()
    Call BuscarProdutosServicos
End Sub

Public Sub ExcluirProduto()
    Call ExcluirProdutoServico
End Sub
