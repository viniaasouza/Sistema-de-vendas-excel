VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProdutos 
   Caption         =   "CadastroDeProdutos"
   ClientHeight    =   7693
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6419
   OleObjectBlob   =   "FormProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    CboTipo.Clear
    CboTipo.AddItem "Produto"
    CboTipo.AddItem "Serviço"
    CboTipo.Text = "Produto"
    
    Call ModProdutosServicos.LimparCampos
    Call ModProdutosServicos.BuscarProdutosServicos
    
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        TxtPreco.Locked = True
        TxtPrecoCusto.Locked = True
        TxtMargem.Locked = True
        TxtPreco.BackColor = 15790320
        TxtPrecoCusto.BackColor = 15790320
        TxtMargem.BackColor = 15790320
        Me.Caption = "Consulta de Produtos e Serviços (Modo Consulta - Recepção)"
    Else
        TxtPreco.Locked = False
        TxtPrecoCusto.Locked = False
        TxtMargem.Locked = False
        TxtPreco.BackColor = &H80000005
        TxtPrecoCusto.BackColor = &H80000005
        TxtMargem.BackColor = &H80000005
        Me.Caption = "Cadastro de Produtos e Serviços"
    End If
End Sub

Private Sub CboTipo_Change()
    If UCase(Trim(CboTipo.Text)) = "SERVIÇO" Or UCase(Trim(CboTipo.Text)) = "SERVICO" Then
        TxtEstoque.Text = "N/A"
        TxtEstoqueMin.Text = "N/A"
        TxtEstoque.Locked = True
        TxtEstoqueMin.Locked = True
        TxtEstoque.BackColor = 15790320
        TxtEstoqueMin.BackColor = 15790320
    Else
        If TxtEstoque.Text = "N/A" Then TxtEstoque.Text = "10"
        If TxtEstoqueMin.Text = "N/A" Then TxtEstoqueMin.Text = "5"
        TxtEstoque.Locked = False
        TxtEstoqueMin.Locked = False
        TxtEstoque.BackColor = &H80000005
        TxtEstoqueMin.BackColor = &H80000005
    End If
End Sub

Private Sub TxtPrecoCusto_AfterUpdate()
    Call CalcularPrecoFinalPorMargem
End Sub

Private Sub TxtMargem_AfterUpdate()
    Call CalcularPrecoFinalPorMargem
End Sub

Private Sub CalcularPrecoFinalPorMargem()
    Dim cStr As String, mStr As String
    cStr = Replace(Replace(TxtPrecoCusto.Text, "R$", ""), " ", "")
    mStr = Replace(Replace(TxtMargem.Text, "%", ""), " ", "")
    
    If IsNumeric(cStr) And IsNumeric(mStr) Then
        Dim cVal As Double, mVal As Double, fVal As Double
        cVal = CDbl(cStr)
        mVal = CDbl(mStr)
        fVal = Round(cVal * (1 + (mVal / 100)), 2)
        TxtPreco.Text = Format(fVal, "#,##0.00")
        TxtPrecoCusto.Text = Format(cVal, "#,##0.00")
        TxtMargem.Text = Format(mVal, "0.0") & "%"
    End If
End Sub

Private Sub TxtPreco_AfterUpdate()
    Dim pStr As String, cStr As String
    pStr = Replace(Replace(TxtPreco.Text, "R$", ""), " ", "")
    cStr = Replace(Replace(TxtPrecoCusto.Text, "R$", ""), " ", "")
    
    If IsNumeric(pStr) Then
        Dim pVal As Double
        pVal = CDbl(pStr)
        TxtPreco.Text = Format(pVal, "#,##0.00")
        
        If IsNumeric(cStr) Then
            Dim cVal As Double
            cVal = CDbl(cStr)
            If cVal > 0 Then
                Dim mCalc As Double
                mCalc = Round(((pVal / cVal) - 1) * 100, 1)
                TxtMargem.Text = Format(mCalc, "0.0") & "%"
            End If
        End If
    End If
End Sub

Private Sub BtSalvar_Click()
    Call ModProdutosServicos.SalvarProdutosServicos
End Sub

Private Sub BtNovo_Click()
    Call ModProdutosServicos.LimparCampos
End Sub

Private Sub BtExcluir_Click()
    Call ModProdutosServicos.ExcluirProdutoServico
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    ModProdutosServicos.vIDProduto = LstDados.List(LstDados.ListIndex, 0)
    CboTipo.Text = LstDados.List(LstDados.ListIndex, 1)
    TxtDescricao.Text = LstDados.List(LstDados.ListIndex, 2)
    
    Dim cStr As String, pStr As String
    cStr = Replace(LstDados.List(LstDados.ListIndex, 3), "R$", "")
    TxtPrecoCusto.Text = Trim(cStr)
    TxtMargem.Text = LstDados.List(LstDados.ListIndex, 4)
    pStr = Replace(LstDados.List(LstDados.ListIndex, 5), "R$", "")
    TxtPreco.Text = Trim(pStr)
    
    On Error Resume Next
    TxtEstoque.Text = LstDados.List(LstDados.ListIndex, 6)
    
    Dim wsP As Worksheet
    Set wsP = ModProdutosServicos.PlanProdutosServicos()
    Dim cel As Range
    Set cel = wsP.Range("A:A").Find(What:=ModProdutosServicos.vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cel Is Nothing Then
        TxtEstoqueMin.Text = cel.Offset(0, 7).Text
    End If
    On Error GoTo 0
    
    BtSalvar.Caption = "ALTERAR"
    On Error Resume Next
    BtExcluir.Enabled = True
    On Error GoTo 0
End Sub

Private Sub TxtPesquisa_Change()
    If TxtPesquisa.Text <> Empty Then
        LblPesquisa.Visible = False
    Else
        LblPesquisa.Visible = True
    End If
    Call ModProdutosServicos.BuscarProdutosServicos
End Sub

Private Sub LblPesquisa_Click()
    TxtPesquisa.SetFocus
End Sub

Private Sub CommandButton2_Click()
    TxtPesquisa.Text = ""
    LblPesquisa.Visible = True
    TxtPesquisa.SetFocus
    Call ModProdutosServicos.BuscarProdutosServicos
End Sub

