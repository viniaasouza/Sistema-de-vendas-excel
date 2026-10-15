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
    Call ModProdutosServicos.LimparCampos
    Call ModProdutosServicos.BuscarProdutos
    
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) = "RECEPCAO" Then
        TxtPreco.Locked = True
        TxtPreco.BackColor = 15790320
        Me.Caption = "Cadastro de Produtos e Serviços (Modo Consulta - Recepção)"
    Else
        TxtPreco.Locked = False
        TxtPreco.BackColor = &H80000005
        Me.Caption = "Cadastro de Produtos e Serviços"
    End If
End Sub

Private Sub BtSalvar_Click()
    Call ModProdutosServicos.SalvarProdutos
End Sub

Private Sub BtNovo_Click()
    Call ModProdutosServicos.LimparCampos
End Sub

Private Sub BtExcluir_Click()
    Call ModProdutosServicos.ExcluirProduto
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    ModProdutosServicos.vIDProduto = LstDados.List(LstDados.ListIndex, 0)
    TxtDescricao.Text = LstDados.List(LstDados.ListIndex, 2)
    
    Dim pStr As String
    pStr = Replace(LstDados.List(LstDados.ListIndex, 3), "R$", "")
    pStr = Trim(pStr)
    TxtPreco.Text = pStr
    
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
    Call ModProdutosServicos.BuscarProdutos
End Sub

Private Sub LblPesquisa_Click()
    TxtPesquisa.SetFocus
End Sub

Private Sub CommandButton2_Click()
    TxtPesquisa.Text = ""
    LblPesquisa.Visible = True
    TxtPesquisa.SetFocus
    Call ModProdutosServicos.BuscarProdutos
End Sub

