VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRelatorios 
   Caption         =   "UserForm1"
   ClientHeight    =   9702.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   12817
   OleObjectBlob   =   "FormRelatorios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Consulta e Relatório de Vendas"
    Call ModRelatorios.InicializarFormRelatorios
End Sub

Private Sub BtFiltrar_Click()
    Call ModRelatorios.FiltrarVendas
End Sub

Private Sub TxtCliente_Change()
    Call ModRelatorios.FiltrarVendas
End Sub

Private Sub CboStatus_Change()
    Call ModRelatorios.FiltrarVendas
End Sub

Private Sub CboFormaPagto_Change()
    Call ModRelatorios.FiltrarVendas
End Sub

Private Sub BtExportarPDF_Click()
    Call ModRelatorios.GerarRelatorioPDF
End Sub

Private Sub BtCupomPDF_Click()
    Call ModRelatorios.GerarComprovanteVendaPDF
End Sub

Private Sub BtAlternarStatus_Click()
    Call ModRelatorios.AlternarStatusVenda
End Sub

Private Sub LstVendas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ModRelatorios.GerarComprovanteVendaPDF
End Sub

Private Sub BtFechar_Click()
    Unload Me
End Sub

