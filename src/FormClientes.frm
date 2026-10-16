VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormClientes 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   9261.001
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9338.001
   OleObjectBlob   =   "FormClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConfigurarLayoutClientes()
    On Error Resume Next
    Me.Width = 476
    Me.Height = 488
    
    LblTitulo.Left = 16: LblTitulo.Top = 8: LblTitulo.Width = 436: LblTitulo.Height = 24
    LblNome.Left = 16: LblNome.Top = 36: LblNome.Width = 200: LblNome.Height = 12: LblNome.Caption = "Nome / Razão Social *"
    TxtNome.Left = 16: TxtNome.Top = 48: TxtNome.Width = 436: TxtNome.Height = 20
    
    LblCPF.Left = 16: LblCPF.Top = 74: LblCPF.Width = 120: LblCPF.Height = 12: LblCPF.Caption = "CPF / CNPJ:"
    TxtCPF.Left = 16: TxtCPF.Top = 86: TxtCPF.Width = 210: TxtCPF.Height = 20
    
    LblTelefone.Left = 242: LblTelefone.Top = 74: LblTelefone.Width = 120: LblTelefone.Height = 12: LblTelefone.Caption = "Telefone / Celular:"
    TxtTelefone.Left = 242: TxtTelefone.Top = 86: TxtTelefone.Width = 210: TxtTelefone.Height = 20
    
    LblEmail.Left = 16: LblEmail.Top = 112: LblEmail.Width = 120: LblEmail.Height = 12: LblEmail.Caption = "E-mail:"
    TxtEmail.Left = 16: TxtEmail.Top = 124: TxtEmail.Width = 210: TxtEmail.Height = 20
    
    LblBairro.Left = 242: LblBairro.Top = 112: LblBairro.Width = 120: LblBairro.Height = 12: LblBairro.Caption = "Bairro:": LblBairro.ForeColor = 0: LblBairro.BackStyle = 0
    TxtBairro.Left = 242: TxtBairro.Top = 124: TxtBairro.Width = 210: TxtBairro.Height = 20
    
    LblEndereco.Left = 16: LblEndereco.Top = 150: LblEndereco.Width = 240: LblEndereco.Height = 12: LblEndereco.Caption = "Endereço Completo (Rua, Nº, Compl.):"
    TxtEndereco.Left = 16: TxtEndereco.Top = 162: TxtEndereco.Width = 436: TxtEndereco.Height = 20
    
    LblDataNasc.Left = 16: LblDataNasc.Top = 188: LblDataNasc.Width = 105: LblDataNasc.Height = 12: LblDataNasc.Caption = "Data Nascimento:": LblDataNasc.ForeColor = 0: LblDataNasc.BackStyle = 0
    TxtDataNasc.Left = 16: TxtDataNasc.Top = 200: TxtDataNasc.Width = 105: TxtDataNasc.Height = 20
    
    LblIdade.Left = 132: LblIdade.Top = 188: LblIdade.Width = 80: LblIdade.Height = 12: LblIdade.Caption = "Idade (Auto):": LblIdade.ForeColor = 0: LblIdade.BackStyle = 0
    TxtIdade.Left = 132: TxtIdade.Top = 200: TxtIdade.Width = 85: TxtIdade.Height = 20
    TxtIdade.Locked = True: TxtIdade.BackColor = 15790320
    
    BtSalvar.Left = 336: BtSalvar.Top = 196: BtSalvar.Width = 116: BtSalvar.Height = 26
    
    TxtPesquisa.Left = 16: TxtPesquisa.Top = 232: TxtPesquisa.Width = 404: TxtPesquisa.Height = 20
    LblPesquisa.Left = 22: LblPesquisa.Top = 235: LblPesquisa.Width = 240: LblPesquisa.Height = 14
    CommandButton2.Left = 424: CommandButton2.Top = 232: CommandButton2.Width = 28: CommandButton2.Height = 20
    
    LblColunas.Left = 18: LblColunas.Top = 258: LblColunas.Width = 434: LblColunas.Height = 12
    LblColunas.Caption = "Cód   Nome do Cliente                            CPF/CNPJ          Telefone         Bairro           Idade"
    
    LstDados.Left = 16: LstDados.Top = 272: LstDados.Width = 436: LstDados.Height = 140
    LstDados.ColumnCount = 7
    LstDados.ColumnWidths = "28 pt;135 pt;72 pt;68 pt;68 pt;45 pt;0 pt"
    
    BtNovo.Left = 16: BtNovo.Top = 420: BtNovo.Width = 80: BtNovo.Height = 26
    BtExcluir.Left = 104: BtExcluir.Top = 420: BtExcluir.Width = 80: BtExcluir.Height = 26
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    Call ConfigurarLayoutClientes
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

Public Sub AtualizarIdadeAutomatica()
    Dim s As String
    s = Trim(TxtDataNasc.Text)
    If s = "" Then
        TxtIdade.Text = ""
        Exit Sub
    End If
    
    ' Se ja tem barras e ano com 4 digitos (ex: DD/MM/AAAA ou D/M/AAAA)
    If Len(s) >= 8 And InStr(s, "/") > 0 Then
        Dim partes() As String
        partes = Split(s, "/")
        If UBound(partes) = 2 Then
            If Len(partes(2)) = 4 And IsDate(s) Then
                TxtIdade.Text = ModClientes.CalcularIdade(s) & " anos"
                Exit Sub
            End If
        End If
    End If
    
    ' Se digitou 8 digitos seguidos sem barra (ex: 15081995)
    Dim soNumeros As String, i As Long, ch As String
    soNumeros = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then soNumeros = soNumeros & ch
    Next i
    
    If Len(soNumeros) = 8 Then
        Dim dtFormatada As String
        dtFormatada = Left(soNumeros, 2) & "/" & Mid(soNumeros, 3, 2) & "/" & Mid(soNumeros, 5, 4)
        If IsDate(dtFormatada) Then
            TxtIdade.Text = ModClientes.CalcularIdade(dtFormatada) & " anos"
            Exit Sub
        End If
    End If
End Sub

Private Sub TxtDataNasc_Change()
    Static formatando As Boolean
    If formatando Then Exit Sub
    
    Dim s As String, limpo As String, i As Long, ch As String
    s = TxtDataNasc.Text
    If s = "" Then
        TxtIdade.Text = ""
        Exit Sub
    End If
    
    ' Se o usuario esta digitando apenas numeros sem barras, aplica mascara automatica DD/MM/AAAA
    If InStr(s, "/") = 0 And InStr(s, "-") = 0 Then
        For i = 1 To Len(s)
            ch = Mid(s, i, 1)
            If ch >= "0" And ch <= "9" Then limpo = limpo & ch
        Next i
        If Len(limpo) > 2 And Len(limpo) <= 4 Then
            formatando = True
            TxtDataNasc.Text = Left(limpo, 2) & "/" & Mid(limpo, 3)
            TxtDataNasc.SelStart = Len(TxtDataNasc.Text)
            formatando = False
        ElseIf Len(limpo) > 4 And Len(limpo) <= 8 Then
            formatando = True
            TxtDataNasc.Text = Left(limpo, 2) & "/" & Mid(limpo, 3, 2) & "/" & Mid(limpo, 5, 4)
            TxtDataNasc.SelStart = Len(TxtDataNasc.Text)
            formatando = False
        End If
    End If
    
    Call AtualizarIdadeAutomatica
End Sub

Private Sub TxtDataNasc_AfterUpdate()
    Call AtualizarIdadeAutomatica
End Sub

Private Sub LstDados_Click()
    If LstDados.ListIndex = -1 Then Exit Sub
    
    ModClientes.vIDCliente = LstDados.List(LstDados.ListIndex, 0)
    TxtNome.Text = LstDados.List(LstDados.ListIndex, 1)
    TxtCPF.Text = LstDados.List(LstDados.ListIndex, 2)
    TxtTelefone.Text = LstDados.List(LstDados.ListIndex, 3)
    
    On Error Resume Next
    TxtBairro.Text = LstDados.List(LstDados.ListIndex, 4)
    TxtIdade.Text = LstDados.List(LstDados.ListIndex, 5)
    TxtEndereco.Text = LstDados.List(LstDados.ListIndex, 6)
    
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Sheets("CLIENTES")
    Dim cel As Range
    Set cel = wsC.Range("A:A").Find(What:=ModClientes.vIDCliente, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cel Is Nothing Then
        TxtEmail.Text = cel.Offset(0, 4).Text
        TxtDataNasc.Text = cel.Offset(0, 7).Text
        Call AtualizarIdadeAutomatica
    End If
    On Error GoTo 0
    
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

