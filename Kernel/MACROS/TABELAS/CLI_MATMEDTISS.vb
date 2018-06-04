'HASH: 939DE136F6DD3D9BF6F9164B98BC8BA5
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT 1 ")
  SQL.Add("  FROM CLI_MATMEDTISS C")
  SQL.Add(" WHERE C.TISSTABELAPRECO = :TABELAPRECO")
  SQL.Add("   AND C.VERSAOTISS = :VERSAOTISS")
  SQL.Add("   AND C.HANDLE <> :HANDLE")

  SQL.ParamByName("TABELAPRECO").AsInteger = CurrentQuery.FieldByName("TISSTABELAPRECO").AsInteger
  SQL.ParamByName("VERSAOTISS").AsInteger = CurrentQuery.FieldByName("VERSAOTISS").AsInteger
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Esta tabela de preço já está cadastrada para esta versão.", "E")
    Exit Sub
  End If
  Set SQL = Nothing

End Sub
