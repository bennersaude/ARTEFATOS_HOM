'HASH: F31C1568FB9748A71B7D2F6B6E7F5F97
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'sms 50287
  Dim qLotacao As Object
  Set qLotacao = NewQuery

  qLotacao.Active = False
  qLotacao.Clear
  qLotacao.Add("SELECT HANDLE")
  qLotacao.Add("  FROM SAM_MIGRACAO_EQLOTACAO")
  qLotacao.Add(" WHERE CONTRATOORIGEM = :CONTRATOORIGEM AND CONTRATODESTINO = :CONTRATODESTINO ")
  qLotacao.Add("       AND LOTACAOORIGEM = :LOTACAOORIGEM AND HANDLE <> :HANDLE")
  qLotacao.ParamByName("CONTRATOORIGEM").AsInteger = CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger
  qLotacao.ParamByName("CONTRATODESTINO").AsInteger = CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger
  qLotacao.ParamByName("LOTACAOORIGEM").AsInteger = CurrentQuery.FieldByName("LOTACAOORIGEM").AsInteger
  qLotacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qLotacao.Active = True

  If Not qLotacao.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Equivalência de lotação existente para a lotação origem informada", "E")
    CanContinue = False
    Set qLotacao = Nothing
    Exit Sub
  End If


  Set qLotacao = Nothing

End Sub

