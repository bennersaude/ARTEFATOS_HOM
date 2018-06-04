'HASH: A8053D4F3B78969F8797822011282EC3
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qSql As Object
  Set qSql = NewQuery

  qSql.Active = False
  qSql.Add("SELECT COUNT(1) ACHOU       ")
  qSql.Add("  FROM SAM_TIPOCOMPLEMENTO  ")
  qSql.Add(" WHERE CODIGO = :pCODIGO    ")

  If CurrentQuery.State = 2 Then  ' Edição
    qSql.Add(" and handle <> :pHandleAtual ")
    qSql.ParamByName("pHandleAtual").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If
  qSql.ParamByName("pCODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger

  qSql.Active = True

  If qSql.FieldByName("ACHOU").AsInteger > 0 Then
    MsgBox("Código já existe.")
    CanContinue = False
    Set qSql = Nothing
    Exit Sub
  End If

  Set qSql = Nothing

End Sub
