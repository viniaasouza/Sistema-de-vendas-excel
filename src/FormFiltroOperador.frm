VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormFiltroOperador 
   Caption         =   "Filtrar Vendas por Operador"
   ClientHeight    =   4305
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5817
   OleObjectBlob   =   "FormFiltroOperador.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormFiltroOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    CboOperador.Clear
    CboOperador.AddItem "? TODOS (Ver Todas as Vendas)"
    
    Dim wsU As Worksheet
    On Error Resume Next
    Set wsU = ThisWorkbook.Sheets("USUARIOS")
    On Error GoTo 0
    
    If Not wsU Is Nothing Then
        Dim rInicio As Long, r As Long, ultLin As Long
        Dim celID As Range
        Set celID = wsU.Range("A:A").Find(What:="ID", LookIn:=xlValues, LookAt:=xlWhole)
        If Not celID Is Nothing Then rInicio = celID.Row + 1 Else rInicio = 5
        
        ultLin = wsU.Cells(wsU.Rows.Count, "A").End(xlUp).Row
        For r = rInicio To ultLin
            Dim n As String, u As String
            n = Trim(wsU.Cells(r, "B").Text)
            u = Trim(wsU.Cells(r, "C").Text)
            If n <> "" Then
                CboOperador.AddItem n & " (" & u & ")"
            End If
        Next r
    End If
    
    If Trim(ModDashboard.gFiltroOperador) = "" Or UCase(Trim(ModDashboard.gFiltroOperador)) = "TODOS" Then
        CboOperador.Text = "? TODOS (Ver Todas as Vendas)"
    Else
        CboOperador.Text = ModDashboard.gFiltroOperador
    End If
End Sub

Private Sub BtnTodos_Click()
    CboOperador.Text = "? TODOS (Ver Todas as Vendas)"
    Call BtnAplicar_Click
End Sub

Private Sub BtnVitoria_Click()
    CboOperador.Text = "Vitória"
    Call BtnAplicar_Click
End Sub

Private Sub BtnGeovana_Click()
    CboOperador.Text = "Geovana"
    Call BtnAplicar_Click
End Sub

Private Sub BtnAdmin_Click()
    CboOperador.Text = "Administrador"
    Call BtnAplicar_Click
End Sub

Private Sub BtnAplicar_Click()
    Dim txt As String, opNome As String
    txt = Trim(CboOperador.Text)
    
    If txt = "" Or InStr(1, txt, "TODOS", vbTextCompare) > 0 Then
        ModDashboard.gFiltroOperador = "TODOS"
    Else
        If InStr(txt, "(") > 0 Then
            opNome = Trim(Left(txt, InStr(txt, "(") - 1))
        Else
            opNome = txt
        End If
        ModDashboard.gFiltroOperador = opNome
    End If
    
    Unload Me
    Call ModDashboard.AtualizarDashboard
End Sub

Private Sub BtnCancelar_Click()
    Unload Me
End Sub

