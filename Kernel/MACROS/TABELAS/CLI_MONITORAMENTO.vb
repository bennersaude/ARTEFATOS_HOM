'HASH: E2BCA1483147CEFB9158258CF6242D86
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qSql As Object
	Set qSql = NewQuery

	qSql.Clear
	qSql.Add("SELECT COUNT(1) EVDUP         ")
	qSql.Add("  FROM CLI_MONITORAMENTO      ")
	qSql.Add(" WHERE PROCEDIMENTO = :EVENTO ")
	qSql.Add("   AND HANDLE <> :HANDLE      ")

	qSql.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("PROCEDIMENTO").AsInteger
	qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	qSql.Active = True

	If qSql.FieldByName("EVDUP").AsInteger > 0 Then
		bsShowMessage("Procedimento já exitente!", "E")
		CanContinue = False
		Exit Sub
	End If

	Set qSql = Nothing

End Sub
