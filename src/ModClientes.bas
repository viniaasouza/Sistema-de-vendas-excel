Attribute VB_Name = "ModClientes"
Option Explicit
Option Compare Text

' =============================================================
' MÓDULO DE GESTÃO DE CLIENTES COM BAIRRO E IDADE (FASE 8)
' =============================================================
Public vIDCliente As Variant

Private Function PlanClientes() As Worksheet
    Set PlanClientes = ThisWorkbook.Sheets("CLIENTES")
End Function

Public Function CalcularIdade(ByVal dataNasc As Variant) As Long
    On Error Resume Next
    If IsDate(dataNasc) Then
        Dim dt As Date, anos As Long
        dt = CDate(dataNasc)
        anos = DateDiff("yyyy", dt, Date)
        If DateSerial(Year(Date), Month(dt), Day(dt)) > Date Then
            anos = anos - 1
        End If
        If anos < 0 Then anos = 0
        CalcularIdade = anos
    Else
        CalcularIdade = 0
    End If
    On Error GoTo 0
End Function

Public Sub SalvarClientes()
    With FormClientes
        .TxtNome.BorderColor = &H80000012
        
        If Trim(.TxtNome.Text) = "" Then
            .TxtNome.BorderColor = &HFF&
            MsgBox "Digite o Nome do cliente!", vbExclamation, "Cadastro de Clientes"
            .TxtNome.SetFocus
            Exit Sub
        End If
        
        Dim cel As Range
        Dim wsC As Worksheet
        Set wsC = PlanClientes()
        
        Dim idadeCalc As Long
        idadeCalc = Val(Replace(.TxtIdade.Text, " anos", ""))
        If idadeCalc <= 0 And IsDate(.TxtDataNasc.Text) Then
            idadeCalc = CalcularIdade(.TxtDataNasc.Text)
        End If
        
        If IsEmpty(vIDCliente) Or vIDCliente = "" Or vIDCliente = 0 Then
            Set cel = wsC.Range("B:B").Find(What:=Trim(.TxtNome.Text), LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                .TxtNome.BorderColor = &HFF&
                MsgBox "Já existe um cliente cadastrado com este Nome!", vbExclamation, "Cliente Duplicado"
                .TxtNome.SetFocus
                Exit Sub
            End If
            
            Dim lin As Long
            lin = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row + 1
            
            Dim novoID As Long
            If lin = 2 Then
                novoID = 1
            Else
                novoID = wsC.Cells(lin - 1, "A").Value + 1
            End If
            
            wsC.Cells(lin, "A").Value = novoID
            wsC.Cells(lin, "B").Value = Trim(.TxtNome.Text)
            wsC.Cells(lin, "C").Value = Trim(.TxtCPF.Text)
            wsC.Cells(lin, "D").Value = Trim(.TxtTelefone.Text)
            wsC.Cells(lin, "E").Value = Trim(.TxtEmail.Text)
            wsC.Cells(lin, "F").Value = Trim(.TxtEndereco.Text)
            wsC.Cells(lin, "G").Value = Trim(.TxtBairro.Text)
            wsC.Cells(lin, "H").Value = Trim(.TxtDataNasc.Text)
            wsC.Cells(lin, "I").Value = idadeCalc
            
            MsgBox "Cliente cadastrado com sucesso!", vbInformation, "Sistema de Vendas"
        Else
            Set cel = wsC.Range("A:A").Find(What:=vIDCliente, LookIn:=xlValues, LookAt:=xlWhole)
            If Not cel Is Nothing Then
                cel.Offset(0, 1).Value = Trim(.TxtNome.Text)
                cel.Offset(0, 2).Value = Trim(.TxtCPF.Text)
                cel.Offset(0, 3).Value = Trim(.TxtTelefone.Text)
                cel.Offset(0, 4).Value = Trim(.TxtEmail.Text)
                cel.Offset(0, 5).Value = Trim(.TxtEndereco.Text)
                cel.Offset(0, 6).Value = Trim(.TxtBairro.Text)
                cel.Offset(0, 7).Value = Trim(.TxtDataNasc.Text)
                cel.Offset(0, 8).Value = idadeCalc
                MsgBox "Cliente alterado com sucesso!", vbInformation, "Sistema de Vendas"
            Else
                MsgBox "Código do cliente não localizado!", vbCritical, "Erro"
            End If
        End If
        
        Call LimparCampos
        Call BuscarClientes
    End With
End Sub

Public Sub ExcluirCliente()
    If IsEmpty(vIDCliente) Or vIDCliente = "" Or vIDCliente = 0 Then
        MsgBox "Selecione um cliente na lista para excluir!", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir o cliente selecionado?", vbQuestion + vbYesNo, "Confirmar Exclusão") = vbYes Then
        Dim cel As Range
        Dim wsC As Worksheet
        Set wsC = PlanClientes()
        Set cel = wsC.Range("A:A").Find(What:=vIDCliente, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not cel Is Nothing Then
            cel.EntireRow.Delete
            MsgBox "Cliente excluído com sucesso!", vbInformation, "Sistema de Vendas"
        Else
            MsgBox "Cliente não localizado!", vbCritical, "Erro"
        End If
        
        Call LimparCampos
        Call BuscarClientes
    End If
End Sub

Public Sub LimparCampos()
    vIDCliente = Empty
    With FormClientes
        .TxtNome.Text = ""
        .TxtCPF.Text = ""
        .TxtTelefone.Text = ""
        .TxtEmail.Text = ""
        .TxtEndereco.Text = ""
        .TxtBairro.Text = ""
        .TxtDataNasc.Text = ""
        .TxtIdade.Text = ""
        .TxtNome.BorderColor = &H80000012
        .BtSalvar.Caption = "SALVAR"
        On Error Resume Next
        .BtExcluir.Enabled = False
        .TxtNome.SetFocus
        On Error GoTo 0
    End With
End Sub

Public Sub BuscarClientes()
    Dim lin As Long, i As Long, ultLinha As Long
    i = 0
    FormClientes.LstDados.Clear
    
    Dim wsC As Worksheet
    Set wsC = PlanClientes()
    
    ultLinha = wsC.Cells(wsC.Rows.Count, "A").End(xlUp).Row
    If ultLinha < 2 Then Exit Sub
    
    Dim qry As String
    qry = Trim(FormClientes.TxtPesquisa.Text)
    
    For lin = 2 To ultLinha
        Dim match As Boolean
        match = False
        
        If qry = "" Then
            match = True
        Else
            Dim nomeC As String, docC As String, bairroC As String
            nomeC = wsC.Cells(lin, "B").Text
            docC = wsC.Cells(lin, "C").Text
            bairroC = wsC.Cells(lin, "G").Text
            If nomeC Like "*" & ModFuncoes.AcSQL(qry) & "*" Or docC Like "*" & qry & "*" Or bairroC Like "*" & ModFuncoes.AcSQL(qry) & "*" Then
                match = True
            End If
        End If
        
        If match Then
            FormClientes.LstDados.AddItem
            FormClientes.LstDados.List(i, 0) = wsC.Cells(lin, "A").Text
            FormClientes.LstDados.List(i, 1) = wsC.Cells(lin, "B").Text
            FormClientes.LstDados.List(i, 2) = wsC.Cells(lin, "C").Text
            FormClientes.LstDados.List(i, 3) = wsC.Cells(lin, "D").Text
            FormClientes.LstDados.List(i, 4) = wsC.Cells(lin, "G").Text
            FormClientes.LstDados.List(i, 5) = IIf(wsC.Cells(lin, "I").Value <> "", wsC.Cells(lin, "I").Text & " anos", "-")
            FormClientes.LstDados.List(i, 6) = wsC.Cells(lin, "F").Text
            i = i + 1
        End If
    Next lin
End Sub

