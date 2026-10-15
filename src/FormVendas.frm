VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVendas 
   Caption         =   "UserForm1"
   ClientHeight    =   2863
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "FormVendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Call ModVendas.IniciarNovaVenda
End Sub

Private Sub CboProduto_Change()
    Call ModVendas.ProdutoSelecionado
End Sub

Private Sub TxtQtd_Change()
    Call ModVendas.AtualizarSubtotalItem
End Sub

Private Sub TxtPrecoUnit_Change()
    Call ModVendas.AtualizarSubtotalItem
End Sub

Private Sub BtAdicionarItem_Click()
    Call ModVendas.AdicionarItem
End Sub

Private Sub LstItens_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ModVendas.CarregarItemParaEdicao
End Sub

Private Sub LstItens_Click()
    On Error Resume Next
    BtRemoverItem.Enabled = (LstItens.ListIndex >= 0)
    On Error GoTo 0
End Sub

Private Sub BtRemoverItem_Click()
    Call ModVendas.RemoverItem
End Sub

Private Sub TxtDesconto_Change()
    Call ModVendas.RecalcularTotaisVenda
End Sub

Private Sub BtFinalizarVenda_Click()
    Call ModVendas.FinalizarVenda
End Sub

Private Sub BtNovaVenda_Click()
    If MsgBox("Deseja cancelar esta venda e iniciar uma nova?", vbQuestion + vbYesNo, "Nova Venda") = vbYes Then
        Call ModVendas.IniciarNovaVenda
    End If
End Sub

