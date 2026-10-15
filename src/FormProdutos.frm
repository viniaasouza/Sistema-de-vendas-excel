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
    Call ModProdutos.LimparCampos
    Call ModProdutos.BuscarProdutos
End Sub

Private Sub BtSalvar_Click()
    Call ModProdutos.SalvarProdutos
End Sub

Private Sub CommandButton1_Click()
    Call ModProdutos.SalvarProdutos
End Sub

Private Sub BtNovo_Click()
    Call ModProdutos.LimparCampos
End Sub

Private Sub BtExcluir_Click()
    Call ModProdutos.ExcluirProduto
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    ' Guarda o ID do produto selecionado
    ModProdutos.vIDProduto = LstDados.List(LstDados.ListIndex, 0)
    
    ' Carrega descrição e preço nos campos
    TxtDescricao.Text = LstDados.List(LstDados.ListIndex, 2)
    
    Dim pStr As String
    pStr = Replace(LstDados.List(LstDados.ListIndex, 3), "R$", "")
    pStr = Trim(pStr)
    TxtPreco.Text = pStr
    
    ' Altera botão para Alterar e habilita Excluir
    BtSalvar.Caption = "ALTERAR"
    On Error Resume Next
    BtExcluir.Enabled = True
    On Error GoTo 0
End Sub

Private Sub TxtPreco_AfterUpdate()
    Dim valTxt As String
    valTxt = Replace(TxtPreco.Text, "R$", "")
    valTxt = Trim(valTxt)
    If IsNumeric(valTxt) Then
        TxtPreco.Text = Format(CDbl(valTxt), "#,##0.00")
    End If
End Sub

Private Sub TxtPesquisa_Change()
    If TxtPesquisa.Text <> Empty Then
        LblPesquisa.Visible = False
    Else
        LblPesquisa.Visible = True
    End If
    Call ModProdutos.BuscarProdutos
End Sub

Private Sub LblPesquisa_Click()
    TxtPesquisa.SetFocus
End Sub

Private Sub CommandButton2_Click()
    TxtPesquisa.Text = ""
    LblPesquisa.Visible = True
    TxtPesquisa.SetFocus
    Call ModProdutos.BuscarProdutos
End Sub

