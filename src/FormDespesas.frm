VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDespesas 
   Caption         =   "Painel Financeiro - Gestão de Despesas, DRE e Fluxo de Caixa"
   ClientHeight    =   8946.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   13496
   OleObjectBlob   =   "FormDespesas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Call ModFinanceiro.InicializarFormDespesas(Me)
End Sub

Private Sub CboCategoriaFluxo_Change()
    Call ModFinanceiro.AoMudarCategoriaFluxo(Me)
End Sub

Private Sub CboStatus_Change()
    Call ModFinanceiro.AoMudarStatusDespesa(Me)
End Sub

Private Sub LstDespesas_Click()
    Call ModFinanceiro.SelecionarItemListaDespesas(Me)
End Sub

Private Sub BtnSalvar_Click()
    Call ModFinanceiro.SalvarDespesaForm(Me)
End Sub

Private Sub BtnDarBaixa_Click()
    Call ModFinanceiro.DarBaixaDespesaForm(Me)
End Sub

Private Sub BtnLimpar_Click()
    Call ModFinanceiro.LimparFormDespesas(Me)
End Sub

Private Sub BtnExcluir_Click()
    Call ModFinanceiro.ExcluirDespesaForm(Me)
End Sub

Private Sub BtnFiltrar_Click()
    Call ModFinanceiro.CarregarListaDespesas(Me)
End Sub

Private Sub BtnVerTodos_Click()
    Me.CboFiltroMes.Text = "TODOS"
    Me.CboFiltroAno.Text = "TODOS"
    Me.CboFiltroStatus.Text = "TODOS"
    Call ModFinanceiro.CarregarListaDespesas(Me)
End Sub

Private Sub BtnIrDRE_Click()
    Unload Me
    Call ModFinanceiro.IrParaDRE
End Sub

Private Sub BtnIrFluxo_Click()
    Unload Me
    Call ModFinanceiro.IrParaFluxoCaixa
End Sub

Private Sub BtnFechar_Click()
    Unload Me
End Sub

