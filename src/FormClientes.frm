VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormClientes 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   8904.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6818
   OleObjectBlob   =   "FormClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Call ModClientes.LimparCampos
    Call ModClientes.BuscarClientes
End Sub

Private Sub BtSalvar_Click()
    Call ModClientes.SalvarClientes
End Sub

Private Sub BtNovo_Click()
    Call ModClientes.LimparCampos
End Sub

Private Sub BtExcluir_Click()
    Call ModClientes.ExcluirCliente
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    ' Guarda o ID do cliente selecionado
    ModClientes.vIDCliente = LstDados.List(LstDados.ListIndex, 0)
    
    ' Preenche os campos
    TxtNome.Text = LstDados.List(LstDados.ListIndex, 1)
    TxtCPF.Text = LstDados.List(LstDados.ListIndex, 2)
    TxtTelefone.Text = LstDados.List(LstDados.ListIndex, 3)
    TxtEmail.Text = LstDados.List(LstDados.ListIndex, 4)
    TxtEndereco.Text = LstDados.List(LstDados.ListIndex, 5)
    
    ' Altera botão para Alterar e habilita Excluir
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
    Call ModClientes.BuscarClientes
End Sub

Private Sub LblPesquisa_Click()
    TxtPesquisa.SetFocus
End Sub

Private Sub CommandButton2_Click()
    TxtPesquisa.Text = ""
    LblPesquisa.Visible = True
    TxtPesquisa.SetFocus
    Call ModClientes.BuscarClientes
End Sub

