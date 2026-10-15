Attribute VB_Name = "ModClientes"
Option Explicit
Option Compare Text

Public vIDCliente As Variant

Private Function PlanClientes() As Worksheet
    Set PlanClientes = ThisWorkbook.Sheets("CLIENTES")
End Function

Sub SalvarClientes()
    With FormClientes
        .TxtNome.BorderColor = &H80000012
        
        ' 1. Validação de campo obrigatório (Nome)
        If Trim(.TxtNome.Text) = "" Then
            .TxtNome.BorderColor = &HFF&
            MsgBox "Digite o Nome ou Razão Social do cliente!", vbExclamation, "Cadastro de Clientes"
            .TxtNome.SetFocus
            Exit Sub
        End If
        
        Dim cel As Range
        
        ' 2. Decisão: Novo Cliente ou Alteração
        If IsEmpty(vIDCliente) Or vIDCliente = "" Or vIDCliente = 0 Then
            ' --- NOVO CLIENTE ---
            ' Evitar duplicidade de nome
            Set cel = PlanClientes.Range("B:B").Find(What:=Trim(.TxtNome.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                .TxtNome.BorderColor = &HFF&
                MsgBox "Já existe um cliente cadastrado com este Nome!", vbExclamation, "Cliente Duplicado"
                .TxtNome.SetFocus
                Exit Sub
            End If
            
            Dim lin As Long
            lin = PlanClientes.Cells(PlanClientes.Rows.Count, "A").End(xlUp).Row + 1
            
            Dim novoID As Long
            If lin = 2 Then
                novoID = 1
            Else
                novoID = PlanClientes.Cells(lin - 1, "A").Value + 1
            End If
            
            PlanClientes.Cells(lin, "A").Value = novoID
            PlanClientes.Cells(lin, "B").Value = Trim(.TxtNome.Text)
            PlanClientes.Cells(lin, "C").Value = Trim(.TxtCPF.Text)
            PlanClientes.Cells(lin, "D").Value = Trim(.TxtTelefone.Text)
            PlanClientes.Cells(lin, "E").Value = Trim(.TxtEmail.Text)
            PlanClientes.Cells(lin, "F").Value = Trim(.TxtEndereco.Text)
            
            MsgBox "Cliente cadastrado com sucesso!", vbInformation, "Sistema de Vendas"
        Else
            ' --- ALTERAR CLIENTE EXISTENTE ---
            Set cel = PlanClientes.Range("A:A").Find(What:=vIDCliente, LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                cel.Offset(0, 1).Value = Trim(.TxtNome.Text)
                cel.Offset(0, 2).Value = Trim(.TxtCPF.Text)
                cel.Offset(0, 3).Value = Trim(.TxtTelefone.Text)
                cel.Offset(0, 4).Value = Trim(.TxtEmail.Text)
                cel.Offset(0, 5).Value = Trim(.TxtEndereco.Text)
                MsgBox "Cliente alterado com sucesso!", vbInformation, "Sistema de Vendas"
            Else
                MsgBox "Código do cliente não localizado na planilha!", vbCritical, "Erro"
            End If
        End If
        
        Call LimparCampos
        Call BuscarClientes
    End With
End Sub

Sub ExcluirCliente()
    If IsEmpty(vIDCliente) Or vIDCliente = "" Or vIDCliente = 0 Then
        MsgBox "Selecione um cliente na lista para excluir!", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir o cliente selecionado?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbYes Then
        Dim cel As Range
        Set cel = PlanClientes.Range("A:A").Find(What:=vIDCliente, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not cel Is Nothing Then
            cel.EntireRow.Delete
            MsgBox "Cliente excluído com sucesso!", vbInformation, "Sistema de Vendas"
        Else
            MsgBox "Cliente não localizado na planilha!", vbCritical, "Erro"
        End If
        
        Call LimparCampos
        Call BuscarClientes
    End If
End Sub

Sub LimparCampos()
    vIDCliente = Empty
    With FormClientes
        .TxtNome.Text = ""
        .TxtCPF.Text = ""
        .TxtTelefone.Text = ""
        .TxtEmail.Text = ""
        .TxtEndereco.Text = ""
        .TxtNome.BorderColor = &H80000012
        .BtSalvar.Caption = "SALVAR"
        On Error Resume Next
        .BtExcluir.Enabled = False
        On Error GoTo 0
        .TxtNome.SetFocus
    End With
End Sub

Sub BuscarClientes()
    Dim lin As Long, i As Long
    Dim ultLinha As Long
    i = 0
    
    FormClientes.LstDados.Clear
    
    ultLinha = PlanClientes.Cells(PlanClientes.Rows.Count, "A").End(xlUp).Row
    If ultLinha < 2 Then Exit Sub
    
    For lin = 2 To ultLinha
        If PlanClientes.Cells(lin, "B").Text Like "*" & AcSQL(FormClientes.TxtPesquisa.Text) & "*" Then
            FormClientes.LstDados.AddItem
            FormClientes.LstDados.List(i, 0) = PlanClientes.Cells(lin, "A").Text
            FormClientes.LstDados.List(i, 1) = PlanClientes.Cells(lin, "B").Text
            FormClientes.LstDados.List(i, 2) = PlanClientes.Cells(lin, "C").Text
            FormClientes.LstDados.List(i, 3) = PlanClientes.Cells(lin, "D").Text
            FormClientes.LstDados.List(i, 4) = PlanClientes.Cells(lin, "E").Text
            FormClientes.LstDados.List(i, 5) = PlanClientes.Cells(lin, "F").Text
            i = i + 1
        End If
    Next lin
End Sub

