VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Call ModSeguranca.ProtegerAbaUsuarios
    ThisWorkbook.Sheets("INÍCIO").Activate
    ActiveWindow.DisplayGridlines = False
    Call ModDashboard.AtualizarDashboard
    
    ' Exigir autenticacao obrigatoria ao abrir a planilha
    ModSeguranca.vUsuarioAtivo = ""
    ModSeguranca.vNomeOperador = ""
    ModSeguranca.vPerfilAtivo = "RECEPCAO"
    ModSeguranca.vIDUsuarioAtivo = 0
    Call ModProdutosServicos.AplicarProtecaoInventario
    
    ' Exibir tela de Login imediatamente ao abrir
    FormLogin.bModoTroca = True
    FormLogin.Show
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name <> "USUARIOS" Then
        Call ModSeguranca.ProtegerAbaUsuarios
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Call ModSeguranca.ProtegerAbaUsuarios
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Call ModProdutosServicos.ProcessarEdicaoPlanilha(Sh, Target)
    On Error GoTo 0
End Sub

