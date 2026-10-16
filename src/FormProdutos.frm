VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProdutos 
   Caption         =   "Cadastro de Produtos, Serviços & Inventário"
   ClientHeight    =   8743.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9338.001
   OleObjectBlob   =   "FormProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCalculandoPreco As Boolean

Private Sub ConfigurarLayoutPadrao()
    On Error Resume Next
    Me.Width = 476
    Me.Height = 462
    
    Label1.Left = 16: Label1.Top = 6: Label1.Width = 436: Label1.Height = 24
    Label1.Caption = "PRODUTOS, SERVIÇOS & ESTOQUE"
    Label1.ForeColor = 16744576: Label1.BackStyle = 0
    
    LblTipo.Left = 16: LblTipo.Top = 36: LblTipo.Width = 95: LblTipo.Height = 14
    LblTipo.Caption = "Tipo do Item:": LblTipo.ForeColor = 0: LblTipo.BackStyle = 0: LblTipo.Font.Bold = True: LblTipo.Visible = True
    CboTipo.Left = 16: CboTipo.Top = 50: CboTipo.Width = 95: CboTipo.Height = 20
    
    Label2.Left = 120: Label2.Top = 36: Label2.Width = 250: Label2.Height = 14
    Label2.Caption = "Descrição do Produto / Serviço:": Label2.ForeColor = 0: Label2.BackStyle = 0: Label2.Font.Bold = True: Label2.Visible = True
    TxtDescricao.Left = 120: TxtDescricao.Top = 50: TxtDescricao.Width = 332: TxtDescricao.Height = 20
    
    LblPrecoCusto.Left = 16: LblPrecoCusto.Top = 76: LblPrecoCusto.Width = 95: LblPrecoCusto.Height = 14
    LblPrecoCusto.Caption = "Preço Custo (R$):": LblPrecoCusto.ForeColor = 0: LblPrecoCusto.BackStyle = 0: LblPrecoCusto.Font.Bold = True: LblPrecoCusto.Visible = True
    TxtPrecoCusto.Left = 16: TxtPrecoCusto.Top = 90: TxtPrecoCusto.Width = 95: TxtPrecoCusto.Height = 20
    
    LblMargem.Left = 120: LblMargem.Top = 76: LblMargem.Width = 100: LblMargem.Height = 14
    LblMargem.Caption = "Margem Lucro (%):": LblMargem.ForeColor = 0: LblMargem.BackStyle = 0: LblMargem.Font.Bold = True: LblMargem.Visible = True
    TxtMargem.Left = 120: TxtMargem.Top = 90: TxtMargem.Width = 95: TxtMargem.Height = 20
    
    Label3.Left = 224: Label3.Top = 76: Label3.Width = 105: Label3.Height = 14
    Label3.Caption = "Preço Venda (R$):": Label3.ForeColor = 0: Label3.BackStyle = 0: Label3.Font.Bold = True: Label3.Visible = True
    TxtPreco.Left = 224: TxtPreco.Top = 90: TxtPreco.Width = 100: TxtPreco.Height = 20
    
    BtSalvar.Left = 336: BtSalvar.Top = 86: BtSalvar.Width = 116: BtSalvar.Height = 26
    
    LblEstoque.Left = 16: LblEstoque.Top = 114: LblEstoque.Width = 95: LblEstoque.Height = 14
    LblEstoque.Caption = "Estoque Atual:": LblEstoque.ForeColor = 0: LblEstoque.BackStyle = 0: LblEstoque.Font.Bold = True: LblEstoque.Visible = True
    TxtEstoque.Left = 16: TxtEstoque.Top = 128: TxtEstoque.Width = 95: TxtEstoque.Height = 20
    
    LblEstoqueMin.Left = 120: LblEstoqueMin.Top = 114: LblEstoqueMin.Width = 100: LblEstoqueMin.Height = 14
    LblEstoqueMin.Caption = "Estoque Mínimo:": LblEstoqueMin.ForeColor = 0: LblEstoqueMin.BackStyle = 0: LblEstoqueMin.Font.Bold = True: LblEstoqueMin.Visible = True
    TxtEstoqueMin.Left = 120: TxtEstoqueMin.Top = 128: TxtEstoqueMin.Width = 95: TxtEstoqueMin.Height = 20
    
    TxtPesquisa.Left = 16: TxtPesquisa.Top = 158: TxtPesquisa.Width = 404: TxtPesquisa.Height = 20
    LblPesquisa.Left = 22: LblPesquisa.Top = 161: LblPesquisa.Width = 240: LblPesquisa.Height = 14
    LblPesquisa.ForeColor = &H80000010: LblPesquisa.BackStyle = 0
    CommandButton2.Left = 424: CommandButton2.Top = 158: CommandButton2.Width = 28: CommandButton2.Height = 20
    
    Label6.Left = 16: Label6.Top = 184: Label6.Width = 436: Label6.Height = 2
    Label4.Left = 18: Label4.Top = 188: Label4.Width = 434: Label4.Height = 14
    Label4.ForeColor = 0: Label4.BackStyle = 0: Label4.Font.Bold = True
    Label4.Caption = "Cód   Tipo         Descrição do Item                                 Custo       Margem   Preço Venda   Est."
    Label5.Visible = False
    
    LstDados.Left = 16: LstDados.Top = 202: LstDados.Width = 436: LstDados.Height = 175
    LstDados.ColumnCount = 7
    LstDados.ColumnWidths = "28 pt;52 pt;148 pt;52 pt;44 pt;56 pt;38 pt"
    
    Label7.Left = 16: Label7.Top = 382: Label7.Width = 436: Label7.Height = 2
    BtNovo.Left = 16: BtNovo.Top = 390: BtNovo.Width = 80: BtNovo.Height = 26
    BtExcluir.Left = 104: BtExcluir.Top = 390: BtExcluir.Width = 80: BtExcluir.Height = 26
    On Error GoTo 0
End Sub

Public Sub AplicarRestricoesPerfil()
    Dim isRecepcao As Boolean
    isRecepcao = (UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO")
    
    If isRecepcao Then
        CboTipo.Locked = True
        TxtDescricao.Locked = True
        TxtPreco.Locked = True
        TxtPrecoCusto.Locked = True
        TxtMargem.Locked = True
        TxtEstoque.Locked = True
        TxtEstoqueMin.Locked = True
        
        CboTipo.BackColor = 15790320
        TxtDescricao.BackColor = 15790320
        TxtPreco.BackColor = 15790320
        TxtPrecoCusto.BackColor = 15790320
        TxtMargem.BackColor = 15790320
        TxtEstoque.BackColor = 15790320
        TxtEstoqueMin.BackColor = 15790320
        
        BtSalvar.Enabled = False
        BtExcluir.Enabled = False
        BtNovo.Enabled = False
        Me.Caption = "Consulta de Produtos, Serviços & Inventário (Modo Leitura - Recepção)"
    Else
        CboTipo.Locked = False
        TxtDescricao.Locked = False
        TxtPreco.Locked = False
        TxtPrecoCusto.Locked = False
        TxtMargem.Locked = False
        
        CboTipo.BackColor = &H80000005
        TxtDescricao.BackColor = &H80000005
        TxtPreco.BackColor = &H80000005
        TxtPrecoCusto.BackColor = &H80000005
        TxtMargem.BackColor = &H80000005
        
        If UCase(Trim(CboTipo.Text)) = "SERVIÇO" Or UCase(Trim(CboTipo.Text)) = "SERVICO" Then
            TxtEstoque.Locked = True
            TxtEstoqueMin.Locked = True
            TxtEstoque.BackColor = 15790320
            TxtEstoqueMin.BackColor = 15790320
        Else
            TxtEstoque.Locked = False
            TxtEstoqueMin.Locked = False
            TxtEstoque.BackColor = &H80000005
            TxtEstoqueMin.BackColor = &H80000005
        End If
        
        BtSalvar.Enabled = True
        BtNovo.Enabled = True
        Me.Caption = "Cadastro de Produtos, Serviços & Inventário"
    End If
End Sub

Private Sub UserForm_Initialize()
    Call ConfigurarLayoutPadrao
    
    CboTipo.Clear
    CboTipo.AddItem "Produto"
    CboTipo.AddItem "Serviço"
    CboTipo.Text = "Produto"
    
    Call ModProdutosServicos.LimparCampos
    Call ModProdutosServicos.BuscarProdutosServicos
    Call AplicarRestricoesPerfil
End Sub

Private Sub CboTipo_Change()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        TxtEstoque.Locked = True
        TxtEstoqueMin.Locked = True
        TxtEstoque.BackColor = 15790320
        TxtEstoqueMin.BackColor = 15790320
        Exit Sub
    End If
    
    If UCase(Trim(CboTipo.Text)) = "SERVIÇO" Or UCase(Trim(CboTipo.Text)) = "SERVICO" Then
        TxtEstoque.Text = "N/A"
        TxtEstoqueMin.Text = "N/A"
        TxtEstoque.Locked = True
        TxtEstoqueMin.Locked = True
        TxtEstoque.BackColor = 15790320
        TxtEstoqueMin.BackColor = 15790320
    Else
        If TxtEstoque.Text = "N/A" Or Trim(TxtEstoque.Text) = "" Then TxtEstoque.Text = "10"
        If TxtEstoqueMin.Text = "N/A" Or Trim(TxtEstoqueMin.Text) = "" Then TxtEstoqueMin.Text = "5"
        TxtEstoque.Locked = False
        TxtEstoqueMin.Locked = False
        TxtEstoque.BackColor = &H80000005
        TxtEstoqueMin.BackColor = &H80000005
    End If
End Sub

Private Sub CalcularPrecoFinalAutomatico()
    On Error Resume Next
    If bCalculandoPreco Then Exit Sub
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then Exit Sub
    Dim c As Double, m As Double, pf As Double
    c = Val(Replace(Replace(TxtPrecoCusto.Text, "R$", ""), ",", "."))
    m = Val(Replace(Replace(TxtMargem.Text, "%", ""), ",", "."))
    If c > 0 Then
        bCalculandoPreco = True
        pf = Round(c * (1 + (m / 100)), 2)
        TxtPreco.Text = Format(pf, "0.00")
        bCalculandoPreco = False
    End If
    On Error GoTo 0
End Sub

Private Sub CalcularMargemReversaAutomatica()
    On Error Resume Next
    If bCalculandoPreco Then Exit Sub
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then Exit Sub
    Dim c As Double, pf As Double, m As Double
    c = Val(Replace(Replace(TxtPrecoCusto.Text, "R$", ""), ",", "."))
    pf = Val(Replace(Replace(TxtPreco.Text, "R$", ""), ",", "."))
    If c > 0 And pf > 0 Then
        bCalculandoPreco = True
        m = Round(((pf / c) - 1) * 100, 1)
        TxtMargem.Text = Format(m, "0.0")
        bCalculandoPreco = False
    End If
    On Error GoTo 0
End Sub

Private Sub TxtPrecoCusto_Change()
    Call CalcularPrecoFinalAutomatico
End Sub

Private Sub TxtMargem_Change()
    Call CalcularPrecoFinalAutomatico
End Sub

Private Sub TxtPreco_Change()
    Call CalcularMargemReversaAutomatica
End Sub

Private Sub TxtPrecoCusto_AfterUpdate()
    Call CalcularPrecoFinalAutomatico
End Sub

Private Sub TxtMargem_AfterUpdate()
    Call CalcularPrecoFinalAutomatico
End Sub

Private Sub TxtPreco_AfterUpdate()
    Call CalcularMargemReversaAutomatica
End Sub

Private Sub BtSalvar_Click()
    Call ModProdutosServicos.SalvarProdutosServicos
End Sub

Private Sub BtNovo_Click()
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        MsgBox "Acesso Restrito: Recepcionistas não possuem permissão para alterar o inventário ou catálogo.", vbExclamation, "Acesso Negado"
        Exit Sub
    End If
    Call ModProdutosServicos.LimparCampos
End Sub

Private Sub BtExcluir_Click()
    Call ModProdutosServicos.ExcluirProdutoServico
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    bCalculandoPreco = True
    ModProdutosServicos.vIDProduto = LstDados.List(LstDados.ListIndex, 0)
    CboTipo.Text = LstDados.List(LstDados.ListIndex, 1)
    TxtDescricao.Text = LstDados.List(LstDados.ListIndex, 2)
    TxtPrecoCusto.Text = Trim(Replace(LstDados.List(LstDados.ListIndex, 3), "R$", ""))
    TxtMargem.Text = Trim(Replace(LstDados.List(LstDados.ListIndex, 4), "%", ""))
    TxtPreco.Text = Trim(Replace(LstDados.List(LstDados.ListIndex, 5), "R$", ""))
    TxtEstoque.Text = LstDados.List(LstDados.ListIndex, 6)
    bCalculandoPreco = False
    
    On Error Resume Next
    Dim wsP As Worksheet
    Set wsP = ModProdutosServicos.PlanProdutosServicos()
    Dim cel As Range
    Set cel = wsP.Range("A:A").Find(What:=ModProdutosServicos.vIDProduto, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cel Is Nothing Then
        TxtEstoqueMin.Text = cel.Offset(0, 7).Text
    End If
    
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        BtSalvar.Caption = "SALVAR"
        BtSalvar.Enabled = False
        BtExcluir.Enabled = False
    Else
        BtSalvar.Caption = "ALTERAR"
        BtSalvar.Enabled = True
        BtExcluir.Enabled = True
    End If
    Call AplicarRestricoesPerfil
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

