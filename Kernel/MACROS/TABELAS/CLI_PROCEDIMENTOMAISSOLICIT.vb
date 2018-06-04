'HASH: 0AAB8ABB5CCBE8F69F21FB8F1934F18D
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1                              ")
  SQL.Add("  FROM CLI_PROCEDIMENTOMAISSOLICIT    ")
  SQL.Add(" WHERE HANDLE       <> :HANDLE        ")
  SQL.Add("   AND ESPECIALIDADE = :ESPECIALIDADE ")
  SQL.Add("   AND EVENTO        = :EVENTO        ")
  SQL.ParamByName("HANDLE").AsInteger        = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  SQL.ParamByName("EVENTO").AsInteger        = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Já existe um registro para o mesmo evento e especialidade!","E")
  End If
  SQL.Active = False
  Set SQL = Nothing
End Sub
