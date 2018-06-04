'HASH: DDCAB92F9B30A7F0CD135E9B042DC580
'Macro: R_GRUPOSRELATORIOS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT HANDLE")
  qSQL.Add("  FROM R_GRUPOSRELATORIOS")
  qSQL.Add(" WHERE HANDLE <> :HANDLE")
  qSQL.Add("   AND NOME = :NOME")
  qSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSQL.ParamByName("NOME").AsString = CurrentQuery.FieldByName("NOME").AsString
  qSQL.Active = True

  If qSQL.FieldByName("HANDLE").AsInteger > 0 Then
    bsShowMessage("Grupo de relatório já existente com esse nome, tente outro.", "E")
    CanContinue = False
  End If

  Set qSQL = Nothing
End Sub
