'HASH: 7EA636155747C514DAB4FCF1CEBAE539
 
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT HANDLE           ")
Consulta.Add("  FROM OFF_SITE_TABELAS ")
Consulta.Add(" WHERE TABELA = :TABELA ")
Consulta.Add("   AND OFFSITE = :OFF  ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("OFF").AsInteger = CurrentQuery.FieldByName("OFFSITE").AsInteger
Consulta.ParamByName("TABELA").AsInteger = CurrentQuery.FieldByName("TABELA").AsInteger
Consulta.Active = True

If (Not Consulta.FieldByName("HANDLE").IsNull) Then
  bsShowMessage("Tabela já cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If

End Sub
