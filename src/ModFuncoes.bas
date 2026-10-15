Attribute VB_Name = "ModFuncoes"

' ========================================================
' MOTOR DE CONSULTAS SQL COM ADODB (AULAS 10 A 15)
' ========================================================
Public Function ConexaoExcel() As Object
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")
    Dim strCon As String
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
             ";Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";"
    con.Open strCon
    Set ConexaoExcel = con
End Function

Public Function ExecutarConsultaSQL(ByVal sqlQuery As String) As Object
    On Error GoTo TrataErro
    Dim con As Object, rs As Object
    Set con = ConexaoExcel()
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, con, 3, 1
    Set ExecutarConsultaSQL = rs
    Exit Function
TrataErro:
    Set ExecutarConsultaSQL = Nothing
End Function

Public Function AcSQL(Valor As String) As String

    Dim N, t, v
    t = ""

    For N = 1 To VBA.Len(Valor)

        v = VBA.Asc(VBA.Mid(Valor, N, 1))

        Select Case v

            Case 39: t = t & "''"
            Case 65: t = t & "[AÁÀÂÃÄ]"
            Case 67: t = t & "[CÇ]"
            Case 69: t = t & "[EÉÈÊË]"
            Case 73: t = t & "[IÍÌÎÏ]"
            Case 79: t = t & "[OÓÒÔÕÖ]"
            Case 85: t = t & "[UÚÙÛÜ]"
            Case 97: t = t & "[aáàâãä]"
            Case 99: t = t & "[cç]"
            Case 101: t = t & "[eéèêë]"
            Case 105: t = t & "[iíìîï]"
            Case 111: t = t & "[oóòôõö]"
            Case 117: t = t & "[uúùûü]"

            Case Else

                If v > 31 And v < 127 Then
                    t = t & VBA.Chr(v)
                Else
                    t = t & "_"
                End If

        End Select

    Next

    AcSQL = t

End Function
