'HASH: AA6B24AA5A9EA572910692B38D477E7A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE ")
  SQL.Add("  FROM SAM_MODEQUIVALENCIA  ")
  SQL.Add(" WHERE HANDLE <> :HANDLE    ")
  SQL.Add("   AND CODIGO = :CODIGO     ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Código informado já cadastrado.", "E")
    CanContinue = False
    Exit Sub
  End If

  Set SQL = Nothing

End Sub
