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
    ' Inicia com perfil restrito e oculta todas as planilhas restritas
    ModSeguranca.vUsuarioAtivo = ""
    ModSeguranca.vNomeOperador = ""
    ModSeguranca.vPerfilAtivo = "RECEPCAO"
    ModSeguranca.vIDUsuarioAtivo = 0
    Call ModSeguranca.AplicarVisibilidadeAbasPorPerfil
    
    ThisWorkbook.Sheets("INÍCIO").Activate
    ActiveWindow.DisplayGridlines = False
    Call ModDashboard.AtualizarDashboard
    Call ModFinanceiro.AtualizarDemonstrativos
    
    ' Exibe tela de Login imediatamente
    FormLogin.bModoTroca = True
    FormLogin.Show
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If UCase(Trim(ModSeguranca.vPerfilAtivo)) <> "ADMIN" Then
        If Sh.Name <> "INÍCIO" Then
            Call ModSeguranca.AplicarVisibilidadeAbasPorPerfil
        End If
    End If
    If Sh.Name <> "USUARIOS" Then
        Call ModSeguranca.ProtegerAbaUsuarios
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' Ao fechar, garante que todas as abas restritas fiquem VeryHidden
    ModSeguranca.vPerfilAtivo = "RECEPCAO"
    Call ModSeguranca.AplicarVisibilidadeAbasPorPerfil
    Call ModSeguranca.ProtegerAbaUsuarios
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Call ModProdutosServicos.ProcessarEdicaoPlanilha(Sh, Target)
    If Sh.Name = "APURACAO_RESULTADOS" Or Sh.Name = "FLUXO_CAIXA" Then
        If Not Intersect(Target, Sh.Range("E5")) Is Nothing Then
            Call ModFinanceiro.AtualizarDemonstrativos
        End If
    End If
    On Error GoTo 0
End Sub

