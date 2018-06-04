'HASH: 6A61E8AAFD7DE52A180610FC8A71145D
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSecao As Object
  Set qSecao = NewQuery

  qSecao.Active = False
  qSecao.Clear
  qSecao.Add("SELECT HANDLE")
  qSecao.Add("  FROM SAM_MIGRACAO_EQSECAO")
  qSecao.Add(" WHERE CONTRATOORIGEM = :CONTRATOORIGEM AND CONTRATODESTINO = :CONTRATODESTINO ")
  qSecao.Add("       AND SECAOORIGEM = :SECAOORIGEM AND HANDLE <> :HANDLE")
  qSecao.ParamByName("CONTRATOORIGEM").AsInteger = CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger
  qSecao.ParamByName("CONTRATODESTINO").AsInteger = CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger
  qSecao.ParamByName("SECAOORIGEM").AsInteger = CurrentQuery.FieldByName("SECAOORIGEM").AsInteger
  qSecao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSecao.Active = True

  If Not qSecao.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Equivalência de seção existente para a seção origem informada", "E")
    CanContinue = False
    Set qSecao = Nothing
    Exit Sub
  End If


  Set qSecao = Nothing

End Sub

