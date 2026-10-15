VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormVendas 
   Caption         =   "UserForm1"
   ClientHeight    =   10703
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10521
   OleObjectBlob   =   "FormVendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bBloqueiaBuscaCli As Boolean
Private bBloqueiaBuscaProd As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Emissão de Vendas (PDV)"
    Call ModVendas.IniciarNovaVenda
End Sub

' === BUSCA INTELIGENTE DE CLIENTES AO DIGITAR ===
Private Sub CboCliente_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyEscape Then Exit Sub
    If bBloqueiaBuscaCli Then Exit Sub
    
    Dim textoDigitado As String
    textoDigitado = CboCliente.Text
    
    bBloqueiaBuscaCli = True
    Call ModVendas.FiltrarClientes(textoDigitado)
    CboCliente.Text = textoDigitado
    bBloqueiaBuscaCli = False
    
    On Error Resume Next
    If CboCliente.ListCount > 0 Then CboCliente.DropDown
    On Error GoTo 0
End Sub

' === BUSCA INTELIGENTE DE PRODUTOS AO DIGITAR ===
Private Sub CboProduto_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyEscape Then Exit Sub
    If bBloqueiaBuscaProd Then Exit Sub
    
    Dim textoDigitado As String
    textoDigitado = CboProduto.Text
    
    bBloqueiaBuscaProd = True
    Call ModVendas.FiltrarProdutos(textoDigitado)
    CboProduto.Text = textoDigitado
    bBloqueiaBuscaProd = False
    
    On Error Resume Next
    If CboProduto.ListCount > 0 Then CboProduto.DropDown
    On Error GoTo 0
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

' === CONTROLE DE FORMA DE PAGAMENTO E CARTÃO ===
Private Sub CboFormaPagto_Change()
    Call ModVendas.FormaPagtoAlterada
End Sub

Private Sub CboParcelas_Change()
    Call ModVendas.RecalcularTotaisVenda
End Sub

Private Sub TxtTaxaCartao_Change()
    Call ModVendas.RecalcularTotaisVenda
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

