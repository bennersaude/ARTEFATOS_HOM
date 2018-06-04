'HASH: 1AEBE7973F711D7784F6D85B3617CB54
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSql As Object
  Set qSql = NewQuery

  qSql.Clear
  qSql.Add("SELECT COUNT(HANDLE)QTD FROM SFN_CLASSEGERENCIAL_CONTACORP WHERE CLASSEGERENCIAL = :CLASSEGERENCIAL AND TIPODOCUMENTO = :TIPODOCUMENTO And HANDLE <> :HANDLE")
  qSql.ParamByName("CLASSEGERENCIAL").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
  qSql.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
  qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSql.Active = True

  If qSql.FieldByName("QTD").AsInteger > 0 Then
    bsShowMEssage("Já existe um cadastro para este tipo de documento!","E")
    CanContinue = False
  End If


End Sub
