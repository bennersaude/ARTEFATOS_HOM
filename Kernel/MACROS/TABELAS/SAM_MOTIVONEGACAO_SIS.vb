'HASH: F7F87EBBD89EB2A4E3A9FDE12734CDA2
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT HANDLE                              ")
    SQL.Add("  FROM SAM_MOTIVONEGACAO_SIS               ")
    SQL.Add(" WHERE SISMOTIVONEGACAO = :SISMOTIVONEGACAO")
    SQL.Add("   AND HANDLE  <> :HANDLE")

    SQL.ParamByName("SISMOTIVONEGACAO").AsInteger = CurrentQuery.FieldByName("SISMOTIVONEGACAO").AsInteger
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

 	If Not SQL.EOF Then
		bsShowMessage("Já existe uma negação com o mesmo 'Motivo de negação do Sistema'" + Chr(13) + "Operação cancelada.", "E")
		CanContinue = False
		Exit Sub
	End If

	SQL.Active = False

	SQL.Clear
	SQL.Add("SELECT HANDLE                              ")
    SQL.Add("  FROM SAM_MOTIVONEGACAO_SIS               ")
    SQL.Add(" WHERE SAMMOTIVONEGACAO = :SAMMOTIVONEGACAO")
    SQL.Add("   AND HANDLE <> :HANDLE")

    SQL.ParamByName("SAMMOTIVONEGACAO").AsInteger = CurrentQuery.FieldByName("SAMMOTIVONEGACAO").AsInteger
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

	If Not SQL.EOF Then
		bsShowMessage("Já existe uma negação com o mesmo 'Motivo de negação do Cliente'" + Chr(13) + "Operação cancelada.", "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQL = Nothing


End Sub
