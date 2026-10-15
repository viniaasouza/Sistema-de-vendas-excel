Attribute VB_Name = "ModProdutos"
Option Explicit
Option Compare Text

Public vIDProduto As Variant

Sub SalvarProdutos()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Salvar ou Alterar Preço de Produto") Then Exit Sub
    End If
    
    With FormProdutos
        .TxtDescricao.BorderColor = &H80000012
        .TxtPreco.BorderColor = &H80000012
        
        If Trim(.TxtDescricao.Text) = "" Then
            .TxtDescricao.BorderColor = &HFF&
            MsgBox "Digite o campo Descrição!", vbExclamation, "Cadastro de Produtos"
            .TxtDescricao.SetFocus
            Exit Sub
        End If
        
        If Trim(.TxtPreco.Text) = "" Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Digite o campo Preço!", vbExclamation, "Cadastro de Produtos"
            .TxtPreco.SetFocus
            Exit Sub
        End If
        
        Dim precoFormatado As String
        precoFormatado = Replace(.TxtPreco.Text, "R$", "")
        precoFormatado = Trim(precoFormatado)
        
        If Not IsNumeric(precoFormatado) Then
            .TxtPreco.BorderColor = &HFF&
            MsgBox "Informe um valor numérico válido para o preço!", vbExclamation, "Cadastro de Produtos"
            .TxtPreco.SetFocus
            Exit Sub
        End If
        
        Dim valPreco As Double
        valPreco = VBA.CDbl(precoFormatado)
        
        Dim cel As Range
        
        If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
            Set cel = PlanProdutos.Range("B:B").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                .TxtDescricao.BorderColor = &HFF&
                MsgBox "Já existe um produto cadastrado com esta Descrição!", vbExclamation, "Produto Duplicado"
                .TxtDescricao.SetFocus
                Exit Sub
            End If
            
            Dim lin As Long
            lin = PlanProdutos.Cells(PlanProdutos.Rows.Count, "A").End(xlUp).Row + 1
            
            Dim novoID As Long
            If lin = 2 Then
                novoID = 1
            Else
                novoID = PlanProdutos.Cells(lin - 1, "A").Value + 1
            End If
            
            PlanProdutos.Cells(lin, "A").Value = novoID
            PlanProdutos.Cells(lin, "B").Value = Trim(.TxtDescricao.Text)
            PlanProdutos.Cells(lin, "C").Value = valPreco
            
            MsgBox "Produto cadastrado com sucesso!", vbInformation, "Cadastro de Produtos"
        Else
            Set cel = PlanProdutos.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
            
            If cel Is Nothing Then
                MsgBox "Produto não localizado para alteração!", vbCritical, "Erro"
                Exit Sub
            End If
            
            Dim celDup As Range
            Set celDup = PlanProdutos.Range("B:B").Find(What:=Trim(.TxtDescricao.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not celDup Is Nothing Then
                If celDup.Row <> cel.Row Then
                    .TxtDescricao.BorderColor = &HFF&
                    MsgBox "Já existe outro produto cadastrado com esta Descrição!", vbExclamation, "Produto Duplicado"
                    .TxtDescricao.SetFocus
                    Exit Sub
                End If
            End If
            
            cel.Offset(0, 1).Value = Trim(.TxtDescricao.Text)
            cel.Offset(0, 2).Value = valPreco
            
            MsgBox "Produto alterado com sucesso!", vbInformation, "Cadastro de Produtos"
        End If
        
        Call LimparCampos
        Call BuscarProdutos
    End With
End Sub

Sub ExcluirProduto()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        If Not ModSeguranca.VerificarAutorizacaoAdmin("Excluir Produto do Catálogo") Then Exit Sub
    End If
    
    If IsEmpty(vIDProduto) Or vIDProduto = "" Or vIDProduto = 0 Then
        MsgBox "Selecione um produto na lista para excluir!", vbExclamation, "Excluir Produto"
        Exit Sub
    End If
    
    Dim cel As Range
    Set cel = PlanProdutos.Range("A:A").Find(What:=vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
    
    If cel Is Nothing Then
        MsgBox "Produto não encontrado!", vbCritical, "Erro"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir o produto: " & cel.Offset(0, 1).Text & "?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbYes Then
        cel.EntireRow.Delete
        MsgBox "Produto excluído com sucesso!", vbInformation, "Exclusão"
        Call LimparCampos
        Call BuscarProdutos
    End If
End Sub

Sub LimparCampos()
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

Sub BuscarProdutos()
    Dim lin As Long, i As Long
    Dim ultLinha As Long
    i = 0
    
    FormProdutos.LstDados.Clear
    
    ultLinha = PlanProdutos.Cells(PlanProdutos.Rows.Count, "A").End(xlUp).Row
    If ultLinha < 2 Then Exit Sub
    
    For lin = 2 To ultLinha
        If PlanProdutos.Cells(lin, "B").Text Like "*" & AcSQL(FormProdutos.TxtPesquisa.Text) & "*" Then
            FormProdutos.LstDados.AddItem
            FormProdutos.LstDados.List(i, 0) = PlanProdutos.Cells(lin, "A").Text
            FormProdutos.LstDados.List(i, 1) = "-"
            FormProdutos.LstDados.List(i, 2) = PlanProdutos.Cells(lin, "B").Text
            FormProdutos.LstDados.List(i, 3) = Format(PlanProdutos.Cells(lin, "C").Value, "R$ #,##0.00")
            i = i + 1
        End If
    Next lin
End Sub

